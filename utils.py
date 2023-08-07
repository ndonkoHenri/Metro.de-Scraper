import os
import openpyxl
from splinter import Browser
import time
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Color

failures = []
error_counter = 0
logs_cache = []

# todo: use OOP for cleaner code?


def get_item_info(browser: Browser, iid, link):
    """
    Uses the xpaths provided to find the name of the product, its price (if it is on sale),
    and whether it is on sale. It returns these values as well as its own ID.

    Args:
        browser: Browser: the browser object
        iid: the id from the excel file
        link: direct link to the article's page

    Returns:
        A tuple with the item's id, name, price and a boolean that indicates if it is on promotion/sale
    """
    item_id, item_name, item_price, is_promotion = iid, None, None, False

    name_xpath = "//div[@class='mfcss_article-detail--title']//child::span"
    id_xpath = "(//div[@class='mfcss_article-detail--articlenumber']//descendant::span)[1]"
    price_xpath = "(//span[@class='mfcss_article-detail--price-breakdown primary']//child::span)[2]"
    price_if_promo_xpath = "//span[@class='mfcss_article-detail--price-breakdown strike']//span//span"
    # price_2_xpath = "(//span[@class='mfcss_article-detail--price-breakdown primary promotion']//child::span)[2]"

    try:
        item_name = browser.find_by_xpath(name_xpath).first.value
        item_id = browser.find_by_xpath(id_xpath).first.value

        # if no price is found using price_xpath, then the article is certainly on promotion, hence we try price_if_promo_xpath
        try:
            item_price = browser.find_by_xpath(price_xpath).first.value
        except Exception as e:
            item_price = browser.find_by_xpath(price_if_promo_xpath).first.value
            is_promotion = True

        # the id on the webpage and that from the source excel file are different
        if item_id != str(iid):
            # write_log(f"! ids are different: Online={item_id} - Local={iid=} - {link=}")
            pass

    except Exception as e:
        write_log(f"Error: {iid} ~ {link} ~ {e} ")

    return item_id, item_name, item_price, is_promotion


def read_source(path):
    """
    Reads the source Excel file of this project and stores the data in a pandas dataframe.
    The first column of the dataframe with header row 'source' contains the needed data.
    """
    df = pd.read_excel(path)
    article_number_col = df["Metro articlenumber"]
    links_col = df["Link"]
    return article_number_col, links_col


def summarize(td):
    """
    Reports the total execution time, number of errors, and failures in to the log file.

    Args:
        td: Pass in the total execution time of the program
    """
    global failures, error_counter
    write_log(f'Total execution time: {td} seconds or {round(td / 60, 1)} minutes')
    write_log(f'Number of errors: {error_counter}')
    # write_log(f"Failures:\n{failures}")


def reject_cookies(browser: Browser):
    """Clicks on the button in the cookie banner to reject cookies. The banner is found in the shadow DOM."""
    time.sleep(3)

    try:
        shadow = browser.find_by_css("cms-cookie-disclaimer").first.shadow_root
    except:
        write_log("Cookie Refusal failed. Retrying..")
        time.sleep(5)
        shadow = browser.find_by_css("cms-cookie-disclaimer").first.shadow_root

    shadow.find_by_css(".btn-secondary.reject-btn.field-reject-button-name").click()


def prepare_result_worksheet(result_path):
    """
    Prepares the result's workbook.
    If the file already exists, it will be opened/used and a new sheet will be created.
    Otherwise, a new Workbook is created with an active Worksheet.

    Return:
        wb: the result's Workbook
        ws: the active Worksheet in the result's workbook
    """
    sheet_name = f"{time.strftime('%d.%m.%Y - %Hh%M')}"

    if os.path.exists(result_path):
        wb = openpyxl.load_workbook(result_path)
        ws = wb.create_sheet(sheet_name)
        wb.active = ws
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

    # add header row
    ws.append(['ID', 'Name', 'Preis'])

    # customize header row
    font = Font(bold=True, size=11)
    for r in ws["A1:C1"]:
        for s in r:
            s.font = font

    return wb, ws


def save_results(wb: Workbook, path):
    """
    Saves the results of the scraping to an excel file.

    Args:
        wb: openpyxl.Workbook: the result's workbook
        path: Specify the path where the results should be saved
    """
    try:
        wb.save(path)
    except PermissionError:
        new_path = path.replace(path.split("\\")[-1], f"Result {time.strftime('%Hh%M')}.xlsx")
        wb.save(new_path)
        write_log(
            f"{'Permission errors were encountered while trying to save the results at the chosen destination. Please, always make sure to close the excel files in use.'.upper()} Result file was instead saved at {new_path}")


def write_log(text):
    """Takes a string as an argument and writes it to the logs.txt file."""
    global logs_cache
    with open("logs.txt", "a", encoding="utf-8") as f:
        f.write(f"\n- {text}")

    logs_cache.append(f"- {text}\n")


def start_automation(browser: Browser, source_path: str = "source.xlsx", destination_path: str = "Result.xlsx"):
    """
    Main start point.
    It calls the necessary functions, starts the scraping, saves and summarizes.
    """
    global error_counter, logs_cache
    start_time = time.time()
    articles_in_promotion = []

    try:
        write_log(f"Automation Initialization on the {time.strftime('%d.%m.%Y - %H:%M %p')}")
        write_log(f"{source_path=} | {destination_path=}")

        article_number_col, links_col = read_source(source_path)

        browser.visit("https://produkte.metro.de/shop/")

        reject_cookies(browser)

        wb, ws = prepare_result_worksheet(destination_path)

        for count, source_info in enumerate(zip(article_number_col, links_col), start=2):
            s_number, s_link = source_info
            if s_link.startswith("https://produkte.metro.de/shop/"):
                browser.visit(s_link)

                iid, name, price, is_promotion = get_item_info(browser, s_number, s_link)

                # if name or price is None (was not scraped successfully), retry several times before giving up
                loop = 0
                while ((name is None) or (price is None)) and (loop != 10):
                    time.sleep(5)
                    iid, name, price, is_promotion = get_item_info(browser, s_number, s_link)

                    # refresh the web page at certain intervals (useful when the page didn't initially load completely)
                    if (loop in [1, 4, 7]) and ((name is None) or (price is None)):
                        browser.reload()

                    loop += 1

                    time.sleep(2.5)

                # while loop was called, and the error was then resolved
                if (name is not None) and (price is not None) and loop > 0:
                    write_log("RESOLVED error above!")
                    ws.append([iid, str(name), str(price)])

                # while loop was called, and the error was NOT resolved
                elif ((name is None) or (price is None)) and loop > 0:
                    error_counter += 1
                    write_log(f"FAILED after while loop: {iid, s_link}")
                    failures.append((iid, s_link))
                    ws.append([iid, str(name), str(price), s_link])

                # went smoothly with no errors - while loop not called (loop==0 is True)
                else:
                    ws.append([iid, str(name), str(price)])

                if is_promotion:
                    articles_in_promotion.append(count)

            # if the link is not a metro link, ignore
            else:
                ws.append([s_number, "IGNORED", "IGNORED"])

        # if there are some articles in sales/promotion -> change the color of the price to red (kind of notifier)
        if articles_in_promotion:
            for i in articles_in_promotion:
                cell = ws.cell(row=i, column=3)
                cell.font = Font(color=Color(rgb="FF0000"))

        save_results(wb, destination_path)

        time_delta = round(time.time() - start_time, None)  # the total duration of the execution/automation
        summarize(time_delta)

    except Exception as e:
        write_log(e)
        write_log(f"{time.strftime('%H:%M %p')} | Etwas ist leider schiefgelaufen! Versuch es noch mal bitte.")
    finally:
        browser.quit()  # close browser session
        return logs_cache


if __name__ == "__main__":
    b = Browser("chrome")
    start_automation(b)
