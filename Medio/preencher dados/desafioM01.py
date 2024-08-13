from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import openpyxl

class Desafio01:

    def __init__(self) -> None:
        self.nav = webdriver.Chrome()
        self.nav.get('https://rpachallenge.com/')
        self.nav.maximize_window()

    def iniciar(self) -> None:
        arq = openpyxl.load_workbook('challenge.xlsx')
        arq_work = arq['Sheet1']

        self.nav.find_element(By.XPATH, "//button[contains(text(),'Start')]").click()
        for i in arq_work.iter_rows(min_row=2, values_only=True):
            if not i[0]:
                continue
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']").send_keys(i[0])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']").send_keys(i[1])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']").send_keys(i[2])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']").send_keys(i[3])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']").send_keys(i[4])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']").send_keys(i[5])
            self.nav.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']").send_keys(i[6])
            self.nav.find_element(By.XPATH,"/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input").click()

        input()


if __name__ == '__main__':
    Desafio01().iniciar()