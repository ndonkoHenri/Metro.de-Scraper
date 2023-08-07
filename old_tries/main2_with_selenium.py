import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())
driver.implicitly_wait(15)

# driver = webdriver.Chrome(executable_path="C:\\BrowserDrivers\\chromedriver_win32\\chromedriver.exe")

driver.get("https://produkte.metro.de/shop/search?q=")


# driver.maximize_window()


# item name: title-wrapper
# item price: price-display-main-row

def get_item_name_and_price(item_num):
    global driver
    item_name, item_price = None, None

    #
    try:
        x = driver.find_element(By.XPATH, "(//a[@class='title'])[1]")
    except Exception as e:
        print(f"An exception was raised while fetching Item Name info for Item_{item_num}. {e}")
    else:
        item_name = x.get_attribute('description')

    #
    try:
        y = driver.find_element(By.XPATH, "((//span[contains(@class,'primary')])[1]//child::span)[2]")
    except Exception as e:
        print(f"An exception was raised while fetching Item Price info for Item_{item_num}. {e}")
    else:
        item_price = y.text

    print(item_name, item_price)


search_field = driver.find_element(By.XPATH,
                                   "//input[@aria-controls='react-autowhatever-autosuggest-search-input']")


def enter_item_num(item_num):
    global driver, search_field
    search_field.click()
    search_field.clear()
    search_field.send_keys(str(item_num))
    search_field.send_keys(Keys.RETURN)

    get_item_name_and_price(item_num)


def close_cookie_banner():
    global driver
    # xpath: //*[@id="main"]/div/div[2]/div/cms-cookie-disclaimer//div/div/div/div/div/div/div/div[2]/button[1]
    # parent_div = driver.find_element(By.CLASS_NAME, "buttons")
    # print(f'{parent_div.find_element(By.CLASS_NAME, "accept-btn")=}')
    # print(f'{driver.find_element(By.CLASS_NAME, "accept-btn")=}')
    # cookie_btns = driver.find_elements(By.TAG_NAME, "button")
    # print([(btn.get_property("class"), btn.get_attribute("class"), btn.get_dom_attribute('class')) for btn in cookie_btns])
    # cookie_btn.click()
    ablehnen_btn = driver.find_element(By.CSS_SELECTOR, ".btn-secondary.reject-btn.field-reject-button-name")
    ablehnen_btn.click()


test_list = [823873, 19986, 929368, 62728]
for i in test_list:
    enter_item_num(i)

# enter_item_num(823873)
# close_cookie_banner()
time.sleep(1000)
