import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class SalesforceManager:
    def __init__(self, url=r"https://login.salesforce.com/"):
        self.driver = webdriver.Chrome(options=Options())
        self.driver.get(url)

    def login(self):
        credentials = json.loads(open(r"C:\Users\abu89\PycharmProjects\SAPAutomation\keys\salesforce.json").read())
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='username']"))).send_keys(
            credentials["username"])
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='password']"))).send_keys(
            credentials["password"])
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='submit']"))).click()

    def search_case(self, case_number):
        self.driver.get(r"https://autelroboticsusa.lightning.force.com/lightning/o/Case/list?filterName=Recent")
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[@class='slds-button slds-button_neutral search-button slds-truncate']")))
        self.driver.find_elements(By.XPATH,
                                  "//button[@class='slds-button slds-button_neutral search-button slds-truncate']")[
            0].click()

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Cases and more...']")))
        search_box = self.driver.find_elements(By.XPATH, "//input[@placeholder='Search Cases and more...']")[0]
        search_box.send_keys(case_number)

        time.sleep(1.5)
        search_box.send_keys(Keys.ENTER)
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//a[@class='scopesItem slds-nav-vertical__action'][@title='Cases']")))
        self.driver.find_elements(By.XPATH, "//a[@class='scopesItem slds-nav-vertical__action'][@title='Cases']")[
            0].click()

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "(//tbody)[2]//tr//th")))
        self.driver.find_elements(By.XPATH, "(//tbody)[2]//tr//th")[0].click()
