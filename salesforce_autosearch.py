import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import usaddress
from sap_manager import InboundPackageInfo


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

    def extract_case_info(self):
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//records-record-layout-item[@field-label='Received by']//div//div//div["
                           "@class='slds-form-element__control'][span]//slot[@name='outputField']")))

        first_name = ""
        last_name = ""
        house_number = ""
        street = ""
        city = ""
        state = ""
        zipcode = ""
        try:
            name = self.driver.find_elements(By.XPATH,
                                                           "//records-record-layout-item[@field-label='Received "
                                                           "by']//div//div//div[@class='slds-form-element__control']["
                                                           "span]//slot["
                                                           "@name='outputField']//lightning-formatted-text")[0].text
            split = name.strip().split(" ")
            first_name = split[0]
            last_name = split[1]
            address = self.driver.find_elements(
                By.XPATH, "//records-record-layout-item["
                          "@field-label='Address info']//div//div//div["
                          "@class='slds-form-element__control']["
                          "span]//slot["
                          "@name='outputField']//lightning-formatted-text")[0].text
            address = usaddress.tag(address)[0]
            house_number = address.get("AddressNumber")
            street = (f"{address.get("StreetNamePreDirectional", "")} "
                      f"{address.get("StreetNamePreModifier", "")} "
                      f"{address.get("StreetNamePreType", "")} "
                      f"{address.get("StreetName", "")} "
                      f"{address.get("StreetNamePostDirectional", "")} "
                      f"{address.get("StreetNamePostModifier", "")} "
                      f"{address.get("StreetNamePostType", "")}"
                      )
            street = ' '.join(street.split())
            city = address.get("PlaceName", "")
            state = address.get("StateName", "")
            zipcode = address.get("ZipCode", "")

        except Exception:
            pass

        package_info = InboundPackageInfo(
            first_name=first_name,
            last_name=last_name,
            box_name="",
            street=street,
            house_number=house_number,
            city=city,
            state=state,
            zipcode=zipcode,
            date_received="",
            case_number="",
            product_sku="",
            zt01_number="",
            vl01n_number="",
            notes=""
        )
        return package_info

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


#salesforce_manager = SalesforceManager()
#salesforce_manager.login()
#salesforce_manager.search_case("00060582")
#salesforce_manager.extract_case_info()

