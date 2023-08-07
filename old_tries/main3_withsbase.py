import time
from selenium.webdriver.common.by import By
from seleniumbase import BaseCase
import pytest
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

BaseCase.main(__name__, __file__)


class Project(BaseCase):
    def get_item_info(self):
        item_name, item_price = None, None

        #
        try:
            x = self.find_element("(//a[@class='title'])[1]", By.XPATH)
            y = self.find_element("((//span[contains(@class,'primary')])[1]//child::span)[2]", By.XPATH)
        except Exception as e:
            print(f"An exception was raised while fetching Item Informations. {e}")
        else:
            item_name = self.get_text("(//a[@class='title'])[1]", By.XPATH)
            item_price = self.get_text("((//span[contains(@class,'primary')])[1]//child::span)[2]", By.XPATH)

        print(item_name, item_price)
        return item_name, item_price

    def test_start(self):

        df = pd.read_excel("source.xlsx")
        source_col = df["source"]
        final = dict()

        self.open_worksheet()

        self.open("https://produkte.metro.de/shop/search?q=")
        time.sleep(10)
        self.cookies_ablehnen()

        for i in source_col:
            self.type("//input[@aria-controls='react-autowhatever-autosuggest-search-input']", f'{str(i)}\n')
            final[i] = (self.get_item_info())

            self.worksheet.append([i, *self.get_item_info()])

        pytest.set_trace()
        self.save_and_close_workbook()
        time.sleep(200)

    def cookies_ablehnen(self):
        # self.assert_true(self.is_element_clickable(
        #     "cms-cookie-disclaimer::shadow button[class='btn-secondary reject-btn field-reject-button-name']"),
        #     'Cookie-Banner button is not clickable.')
        self.click("cms-cookie-disclaimer::shadow button[class='btn-secondary reject-btn field-reject-button-name']")
        
    def open_worksheet(self):
        self.wb = Workbook()
        self.worksheet = self.wb.active

        # little customisations
        self.worksheet.title = "Sheet_1"
        self.worksheet.append(['Produkt_ID', 'Produkt_Name', 'Produkt_Preis'])
        font = Font(bold=True, size=20)
        for r in self.worksheet["A1:A3"]:
            for s in r:
                s.font = font

    def save_and_close_workbook(self):
        # save the workbook
        self.wb.save("Ergebnis.xlsx")
