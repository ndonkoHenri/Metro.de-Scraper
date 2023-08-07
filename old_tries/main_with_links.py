import time
from selenium.webdriver.common.by import By
from seleniumbase import BaseCase
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

BaseCase.main(__name__, __file__)


class Project(BaseCase):
    my_timeout = 25
    error_counter = 0
    failures = []

    def get_item_info(self, iid, link):
        """
        Fetches/scrapes the name and price of the item webpage in focus.
        It returns a tuple of strings containing the item's ID, name and price respectively.
        """
        self.open(link)
        item_id, item_name, item_price = None, None, None

        name_xpath = "//div[@class='mfcss_article-detail--title']//child::span"
        price_xpath = "(//span[@class='mfcss_article-detail--price-breakdown primary']//child::span)[2]"
        price_2_xpath = "(//span[@class='mfcss_article-detail--price-breakdown primary promotion']//child::span)[2]"
        id_xpath = "(//div[@class='mfcss_article-detail--articlenumber']//descendant::span)[1]"

        try:
            item_name = self.get_text(name_xpath, By.XPATH, timeout=self.my_timeout)
            item_id = self.get_text(id_xpath, By.XPATH)
            # some items have a different xpath for their prices, so try both xpaths
            try:
                item_price = self.get_text(price_xpath, By.XPATH, timeout=self.my_timeout)
            except:
                item_price = self.get_text(price_2_xpath, By.XPATH, timeout=self.my_timeout)

            if item_id != str(iid):
                self.write_log(f"! ids are different: {self.get_text(id_xpath, By.XPATH)=} - {iid=} - {link=}")

        except Exception as e:
            print(f"An exception was raised while fetching Item Informations. {e}")
            self.error_counter += 1

        return str(item_id if item_id is not None else iid), str(item_name), str(item_price)

    def read_source(self):
        """
        Reads the source Excel file of this project and stores the data in a pandas dataframe.
        The first column of the dataframe with header row 'source' contains the needed data.
        """
        df = pd.read_excel("source.xlsx")
        self.artikel_nummer_col = df["Metro Artikelnummer"]
        self.links_col = df["Link"]

    def test_start(self):
        """
        The main function of this test.
        """
        self.start_time = time.time()

        self.read_source()

        self.prepare_ergebnis_worksheet()

        # Open website
        self.open("https://produkte.metro.de/shop/")
        self.maximize_window()

        self.reject_cookies()

        for count, source_info in enumerate(zip(self.artikel_nummer_col, self.links_col), start=1):
            s_number, s_link = source_info

            if s_link.startswith("https://produkte.metro.de/shop/"):
                iid, name, price = self.get_item_info(s_number, s_link)

                # save failures and retry them
                if iid is None or name is None or price is None:
                    iid, name, price = self.get_item_info(s_number, s_link)
                    if iid is None or name is None or price is None:
                        self.failures.append((s_number, s_link))

                self.worksheet.append([iid, name, price])
            else:
                self.worksheet.append([s_number, "IGNORED", "IGNORED"])

            if count % 25 == 0:
                self.save_ergebnis()

        self.write_log(str(self.failures))
        if self.failures:
            self.worksheet.append("**", "**", "**")

        self.save_ergebnis()

        self.end_time = time.time()

        # summary
        self.write_log(f'\n\nTotal execution time: {round(self.end_time - self.start_time, None)} seconds')
        self.write_log(f'Number of errors: {self.error_counter}')
        self.write_log(f'Number of Articles checked: {count}')

        # pytest.set_trace()

    def reject_cookies(self):
        """
        The reject_cookies function is used to reject the cookie banner.
        It will click on the button with class 'btn-secondary reject-btn field-reject-button-name'
        """
        self.wait_for_element_clickable(
            "cms-cookie-disclaimer::shadow button[class='btn-secondary reject-btn field-reject-button-name']", By.XPATH,
            timeout=self.my_timeout)
        self.click("cms-cookie-disclaimer::shadow button[class='btn-secondary reject-btn field-reject-button-name']")

    def prepare_ergebnis_worksheet(self):
        """
        Creates a new workbook and activates a worksheet, which will be used when saving the scraped results.
        """
        self.wb = Workbook()
        self.worksheet = self.wb.active

        # set sheet's title
        self.worksheet.title = "Ergebnis"

        # add header row
        self.worksheet.append(['ID', 'Name', 'Preis'])

        # customize header row
        font = Font(bold=True, size=11)
        for r in self.worksheet["A1:C1"]:
            for s in r:
                s.font = font

    def save_ergebnis(self):
        """
        Saves the workbook and closes it.
        """
        self.wb.save("Ergebnis.xlsx")

    def write_log(self, text):
        with open("../Metro.de Scraper/log.txt", "a", encoding="utf-8") as f:
            f.write(f"{text}\n")
