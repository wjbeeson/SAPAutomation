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
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime


class SalesforceManager:
    def __init__(self, url=r"https://login.salesforce.com/"):
        chrome_options = Options()
        chrome_options.add_argument("--disable-notifications")
        #chrome_options.add_argument('--headless=new')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.get(url)

    def login(self):
        credentials = json.loads(open(r"keys\salesforce.json").read())
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
            street = (f'{address.get("StreetNamePreDirectional", "")} '
                      f'{address.get("StreetNamePreModifier", "")} '
                      f'{address.get("StreetNamePreType", "")} '
                      f'{address.get("StreetName", "")} '
                      f'{address.get("StreetNamePostDirectional", "")} '
                      f'{address.get("StreetNamePostModifier", "")} '
                      f'{address.get("StreetNamePostType", "")}'
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

    def execute_on_web_element(self, xpath, action, index=-1):
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        if index == -1:
            index = len(self.driver.find_elements(By.XPATH, xpath)) - 1
        element = self.driver.find_elements(By.XPATH, xpath)[index]
        return action(element)
    def search_case(self, case_number):
        self.driver.get(f"https://autelroboticsusa.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?str={case_number}#!/fen=500&initialViewMode=detail&str={case_number}")
        id = self.execute_on_web_element("(//div[@class='pbBody']//tr)[2]//th//a", lambda f: f.get_attribute("data-seclki"))
        self.driver.get(f"https://autelroboticsusa.lightning.force.com/lightning/r/Case/{id}/view")


    def send_replacement_invoice(self, case_number, material_sku, price):
        self.search_case(case_number)
        self.execute_on_web_element('html', lambda f: f.send_keys(Keys.END))
        time.sleep(1)
        self.execute_on_web_element('html', lambda f: f.send_keys(Keys.HOME))
        account_name = self.execute_on_web_element("((//span[.='Account Information']/../../../div/div/slot"
                                                   "/records-record-layout-row)[1]//slot//records-record-layout-"
                                                   "item[@field-label='Account Name']//span[contains(.,'@gmail')"
                                                   "])[2]", lambda f: f.text)

        # Click new order button
        self.execute_on_web_element("//article[@aria-label='Internal Work Order']//lst-related-list-view-"
                                    "manager//lst-common-list-internal//lst-list-view-manager-header//div//div//"
                                    "div//div//div//h2//a", lambda f: f.click())

        # Click new button
        self.execute_on_web_element("//button[@name='New']", lambda f: f.click())

        # Click Replacement Button
        self.execute_on_web_element("(//label[@class='slds-radio topdown-radio-container slds-clearfix'])[7]//span"
                                    , lambda f: f.click())

        # Click Next Button
        self.execute_on_web_element("//div[@class='inlineFooter']//button", lambda f: f.click(), 1)

        # Select Current Step
        self.execute_on_web_element("//button[@aria-label='Current step, --None--']", lambda f: f.click())
        self.execute_on_web_element("//button[@aria-label='Current step, --None--']/../..//div[@role='listbox']//"
                                    "lightning-base-combobox-item", lambda f: f.click(),2)

        # Select Region
        self.execute_on_web_element("//button[@aria-label='Region, --None--']", lambda f: f.click())
        self.execute_on_web_element("//button[@aria-label='Region, --None--']/../..//div["
                                    "@role='listbox']//lightning-base-combobox-item", lambda f: f.click(), 2)

        # Select account name
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Accounts...']")))
        combo_box = self.driver.find_elements(By.XPATH, "//input[@placeholder='Search Accounts...']")[0]
        _ = combo_box.location_once_scrolled_into_view
        combo_box.click()
        combo_box.send_keys(account_name)
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{account_name}']/../../../../div["
                                                      f"@role='listbox']//lightning-base-combobox-item//span"
                                                      f"//lightning-icon[@icon-name='utility:search']")))
        self.driver.find_elements(By.XPATH,
                                  f"//input[@data-value='{account_name}']/../../../../div["
                                  f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

        # Click Save Button
        self.execute_on_web_element("//button[@name='SaveEdit']", lambda f: f.click())

        # Click invoice button
        self.execute_on_web_element("//span[@title='Proforma Invoice']/..", lambda f: f.click())

        # Click new button
        self.execute_on_web_element("(//button[@name='New'])[2]", lambda f: f.click())

        # Click Overseas Button
        self.execute_on_web_element(f"//span[.='Overseas']", lambda f: f.click())

        # Click Next Button
        self.execute_on_web_element(f"//span[.='Next']/..", lambda f: f.click())

        # Select account name
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Accounts...']")))
        combo_box = self.driver.find_elements(By.XPATH, "//input[@placeholder='Search Accounts...']")[0]
        _ = combo_box.location_once_scrolled_into_view
        combo_box.click()
        combo_box.send_keys(account_name)
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{account_name}']/../../../../div["
                                                      f"@role='listbox']//lightning-base-combobox-item//span"
                                                      f"//lightning-icon[@icon-name='utility:search']")))
        self.driver.find_elements(By.XPATH,
                                  f"//input[@data-value='{account_name}']/../../../../div["
                                  f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

        # Select Price Book
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Price Books...']")))
        combo_box = self.driver.find_elements(By.XPATH, "//input[@placeholder='Search Price Books...']")[0]
        _ = combo_box.location_once_scrolled_into_view
        combo_box.click()
        price_book = "North America Price Book"
        combo_box.send_keys("North America Price Book")

        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{price_book}']/../../../../div["
                                                      f"@role='listbox']//lightning-base-combobox-item//span"
                                                      f"//lightning-icon[@icon-name='utility:search']")))
        self.driver.find_elements(By.XPATH,
                                  f"//input[@data-value='{price_book}']/../../../../div["
                                  f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

        # Click Save Button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@name='SaveEdit']")))
        time.sleep(2)
        self.driver.find_elements(By.XPATH, "//button[@name='SaveEdit']")[0].click()
        time.sleep(2)
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@role='alertdialog']//div//div//div//span//a//div")))
        def get_active_bar():
            while True:
                self.driver.refresh()
                time.sleep(3)
                # Check to see if add product is active
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
                                                              "[@class='slds-col slds-no-flex slds-grid slds-"
                                                              "grid_vertical-align-center horizontal "
                                                              "actionsContainer']//div//"
                                                              "runtime_platform_actions-actions-ribbon//ul//li")))
                result = self.driver.find_elements(By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
                                                             "[@class='slds-col slds-no-flex slds-grid slds-"
                                                             "grid_vertical-align-center horizontal "
                                                             "actionsContainer']//div//"
                                                             "runtime_platform_actions-actions-ribbon//ul//li")
                time.sleep(3)
                check = self.driver.find_elements(By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
                                                             "[@class='slds-col slds-no-flex slds-grid slds-"
                                                             "grid_vertical-align-center horizontal "
                                                             "actionsContainer']//div//"
                                                             "runtime_platform_actions-actions-ribbon//ul//li")
                if len(check) == len(result):
                    break
            return result

        # Check to see if add product button is active
        active_bar = get_active_bar()
        if len(active_bar) < 4:
            # Click Dropdown if not in active bar
            time.sleep(1)
            active_bar[len(active_bar) - 1].click()
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                          "down-trigger slds-dropdown-trigger_click slds-is-open']")))
            self.driver.find_elements(By.XPATH,
                                      "//runtime_platform_actions-action-renderer[@title='Add Product']")[
                len(self.driver.find_elements(By.XPATH, "//runtime_platform_actions-action-renderer"
                                                        "[@title='Add Product']"))].click()
        else:
            # Click button if in active bar
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"(//button[.='Add Product'])[1]")))
            self.driver.find_elements(By.XPATH, f"//button[.='Add Product']")[len(self.driver.find_elements(By.XPATH, f"//button[.='Add Product']"))-1].click()

        # Row is 0 based index. Chooses which material line item to add
        def add_material(current_material, price):
            # Click Add Product Button in window
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//lightning-button[@variant='brand']//button[@class='slds-button "
                               "slds-button_brand'][.='Add Product']")))
            self.driver.find_elements(By.XPATH, "//lightning-button[@variant='brand']//button[@class='slds-button "
                                                "slds-button_brand'][.='Add Product']")[
                len(self.driver.find_elements(By.XPATH,
                                              "//lightning-button[@variant='brand']//button[@class='slds-button "
                                              "slds-button_brand'][.='Add Product']")) - 1].click()

            # Send material to combo box
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search by Material']")))
            search_box = self.driver.find_elements(By.XPATH, "//input[@placeholder='Search by Material']")[
                len(self.driver.find_elements(By.XPATH, "//input[@placeholder='Search by Material']")) - 1]
            search_box.click()
            search_box.send_keys(current_material)

            # Click search button
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//button[.='Search']")))
            self.driver.find_elements(By.XPATH, "//button[.='Search']")[
                len(self.driver.find_elements(By.XPATH, "//button[.='Search']")) - 1].click()

            # Select material
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//span[.='Material Number']/../../../../../../../tbody/tr)"
                                                          "[1]""//td//lightning-primitive-cell-checkbox")))
            self.driver.find_elements(By.XPATH, "(//span[.='Material Number']/../../../../../../../tbody/tr)[1]"
                                                "//td//lightning-primitive-cell-checkbox")[0].click()

            # Click the next button
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//button[.='Next']")))
            self.driver.find_elements(By.XPATH, "//button[.='Next']")[
                len(self.driver.find_elements(By.XPATH, "//button[.='Next']")) - 1].click()

            # Calculate Price Deduction
            current_price = float(self.driver.find_elements(By.XPATH, f"//lightning-base-formatted-text"
                                                                      f"[.='{current_material}']/../../../../../td")
                                  [7].text.split(" ")[1].replace(",", ""))
            discount = current_price - price

            # Click the discount field
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"(//lightning-base-formatted-text[.='{material_sku}']"
                                                          f"/../../../../../td)"
                                                          f"[9]//lightning-primitive-cell-factory//span//button")))
            self.driver.find_elements(By.XPATH,
                                      f"(//lightning-base-formatted-text[.='{current_material}']/../../../../../td)[9]"
                                      f"//lightning-primitive-cell-factory//span//button")[0].click()

            # Send discount
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//input[@name='dt-inline-edit-currency']")))
            discount_field = self.driver.find_elements(By.XPATH, f"//input[@name='dt-inline-edit-currency']")[0]
            discount_field.send_keys(discount)
            discount_field.send_keys(Keys.ENTER)

        add_material(material_sku, price)
        add_material("901030163", 0)  # This is the SKU for 2nd level labor

        # Click "Comfirm" button *rolls eyes*
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
        self.driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()


        # Check to see if payment request button is active
        active_bar = get_active_bar()
        active_bar_len = len(active_bar)
        if active_bar_len < 5:
            # Click Dropdown if not in active bar
            time.sleep(1)
            active_bar[active_bar_len - 1].click()
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                          "down-trigger slds-dropdown-trigger_click slds-is-open']")))
            self.driver.find_elements(By.XPATH,
                                      "//runtime_platform_actions-action-renderer[@title='Create Payment Request']")[
                0].click()
        else:
            # Click button if in active bar
            WebDriverWait(self.driver, 10).until(
                 EC.presence_of_element_located((By.XPATH, f"(//button[.='Create Payment Request'])[1]")))
            self.driver.find_elements(By.XPATH, f"//button[.='Create Payment Request']")[0].click()

        # Click "Comfirm" button *rolls eyes all the way around*
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
        self.driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()

        # Check to see if generate button is active
        active_bar = get_active_bar()
        active_bar_len = len(active_bar)
        if active_bar_len < 7:
            # Click Dropdown if not in active bar
            active_bar[active_bar_len - 1].click()
            time.sleep(1)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                          "down-trigger slds-dropdown-trigger_click slds-is-open']")))
            self.driver.find_elements(By.XPATH,
                                      "//runtime_platform_actions-action-renderer[@title='Genreate "
                                      "and Send Quotation']")[0].click()
        else:
            # Click button if in active bar
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//button[.='Genreate and Send Quotation']")))
            self.driver.find_elements(By.XPATH, f"//button[.='Genreate and Send Quotation']")[0].click()

        # Click "Comfirm" button *rolls eyes all the way around*
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
        self.driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()

        # Wait until the PDF Saves
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, f"//div[.='Saving quote PDF']")))

        # Load payment URL by putting it on screen
        self.driver.find_element(By.TAG_NAME, 'html').send_keys(Keys.END)
        time.sleep(5)

        # Get Payment Name
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//img[@src='https://autelroboticsusa.my.salesforce.com/img/icon/t4v35/custom/"
                           f"custom48_120.png']/../../../../../../../../../div")))
        split = self.driver.find_elements(By.XPATH,
                                                 f"//img[@src='https://autelroboticsusa.my.salesforce.com/img/icon"
                                                 f"/t4v35/custom/"
                                                 f"custom48_120.png']/../../../../../../../../../div")[0].text.split(
            " ")
        payment_name = ""
        for text in split:
            if datetime.today().strftime('%Y-%m-%d') in text:
                payment_name = text

        # Get Payment URL
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, f"//span[.='{payment_name}']/../../../../a")))
        payment_page_url = (self.driver.find_elements(By.XPATH, f"//span[.='{payment_name}']/../../../../a")[0]
                            .get_attribute("href"))

        self.driver.get(payment_page_url)
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, f"//span[.='Payment Link']/../../../div/div[@class='slds-"
                                                      f"form-element__control']//span//slot//records-formula-output//"
                                                      f"slot//lightning-formatted-text//a")))
        return self.driver.find_elements(By.XPATH,
                                         f"//span[.='Payment Link']/../../../div/div[@class='slds-form-element__control']"
                                         f"//span//slot//records-formula-output//slot//"
                                         f"lightning-formatted-text//a")[0].get_attribute("href")


manager = SalesforceManager()
manager.login()
case_number = "00057550"
nano = manager.send_replacement_invoice(case_number, "102000750", 400)
lite = manager.send_replacement_invoice(case_number, "102000722", 800)
manager.search_case(case_number)
print(f"EVO Nano+ Premium Bundle [$400.00, Normally $719.00]: {nano}")
print(f"EVO Lite+ Premium Bundle [$800.00, Normally $999.00]: {lite}")
input()
