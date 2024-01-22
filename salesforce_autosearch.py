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


class SalesforceManager:
    def __init__(self, url=r"https://login.salesforce.com/"):
        chrome_options = Options()
        chrome_options.add_argument("--disable-notifications")
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

    def search_case(self, case_number):
        while True:
            try:
                self.driver.get(r"https://autelroboticsusa.lightning.force.com/lightning/o/Case/list?filterName=00B4x00000IadxUEAR")

                # Search the cases
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//input[@aria-label='Search 西安全部未处理北美个案 list view.']")))
                search_input = self.driver.find_elements(By.XPATH, f"//input[@aria-label='Search 西安全部未处理北美个案 list view.']")[0]
                search_input.send_keys(case_number)
                search_input.send_keys(Keys.ENTER)

                # Wait until search goes through
                while True:
                    menu_text = self.driver.find_elements(By.XPATH, f"//span[contains(.,'Updated')]//span/..")[0].text
                    results = ""
                    numbers = "0123456789"
                    for letter in menu_text:
                        if letter in numbers:
                            results = results + letter
                        else:
                            break
                    if int(results) < 50:
                        break
                    else:
                        time.sleep(0.1)

                # Click the top result
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"((//span[.='邮件状态']/../../../../../../../table/tbody/tr)[1]/th)[1]")))
                self.driver.find_elements(By.XPATH, f"((//span[.='邮件状态']/../../../../../../../table/tbody/tr)[1]/th)[1]")[
                    0].click()
                break
            except Exception:
                pass

    def send_replacement_invoice(self, case_number, material_sku, price):
        action = ActionChains(self.driver)
        self.search_case(case_number)
        account_name = "fabianlinda1961@gmail.com" #TODO: Change this to the account name
        # Click new order button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//article[@aria-label='Internal Work "
                                                      "Order']//lst-related-list-view-manager//lst-common-list-internal//lst"
                                                      "-list-view-manager-header//div//div//div//div//div//h2//a")))
        self.driver.find_elements(By.XPATH, "//article[@aria-label='Internal Work "
                                            "Order']//lst-related-list-view-manager//lst-common-list-internal//lst"
                                            "-list-view-manager-header//div//div//div//div//div//h2//a")[0].click()
        # Click new button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@name='New']")))
        self.driver.find_elements(By.XPATH, "//button[@name='New']")[0].click()

        # Click Replacement Button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//label[@class='slds-radio topdown-radio-container slds-clearfix']")))
        self.driver.find_elements(By.XPATH,
                                  "(//label[@class='slds-radio topdown-radio-container slds-clearfix'])[7]//span")[
            0].click()

        # Click Next Button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='inlineFooter']//button")))
        self.driver.find_elements(By.XPATH, "//div[@class='inlineFooter']//button")[1].click()

        # Select Current Step
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Current step, --None--']")))
        self.driver.find_elements(By.XPATH, "//button[@aria-label='Current step, --None--']")[0].click()

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Current step, --None--']/../..//div["
                                                      "@role='listbox']//lightning-base-combobox-item")))
        self.driver.find_elements(By.XPATH,
                                  "//button[@aria-label='Current step, --None--']/../..//div["
                                  "@role='listbox']//lightning-base-combobox-item")[2].click()

        # Select Region
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Region, --None--']")))
        self.driver.find_elements(By.XPATH, "//button[@aria-label='Region, --None--']")[0].click()

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Region, --None--']/../..//div["
                                                      "@role='listbox']//lightning-base-combobox-item")))
        self.driver.find_elements(By.XPATH,
                                  "//button[@aria-label='Region, --None--']/../..//div["
                                  "@role='listbox']//lightning-base-combobox-item")[2].click()

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
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@name='SaveEdit']")))
        self.driver.find_elements(By.XPATH, "//button[@name='SaveEdit']")[0].click()

        # Get the work order number
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, f"//div[.='Saving quote PDF']")))
        pass

        # Click invoice button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@title='Proforma Invoice']/..")))
        self.driver.find_elements(By.XPATH, "//span[@title='Proforma Invoice']/..")[0].click()

        # Click new button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@name='New']")))
        self.driver.find_elements(By.XPATH, "//button[@name='New']")[
            len(self.driver.find_elements(By.XPATH, "//button[@name='New']")) - 1].click()

        # Click Overseas Button
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//span[.='Overseas']")))
        self.driver.find_elements(By.XPATH, f"//span[.='Overseas']")[len(self.driver.find_elements(By.XPATH, f"//span[.='Overseas']"))-1].click()

        # Click Next Button
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH,f"//span[.='Next']/..")))
        self.driver.find_elements(By.XPATH, f"//span[.='Next']")[
            len(self.driver.find_elements(By.XPATH, f"//span[.='Next']")) - 1].click()

        self.driver.refresh()
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
        self.driver.find_elements(By.XPATH, "//button[@name='SaveEdit']")[0].click()

        # Click Add Product Button on main screen
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[.='Add Product']")))
        self.driver.find_elements(By.XPATH, f"//button[.='Add Product']")[
            len(self.driver.find_elements(By.XPATH, f"//button[.='Add Product']")) - 1].click()

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
                len(self.driver.find_elements(By.XPATH, "//input[@placeholder='Search by Material']"))-1]
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
            self.driver.find_elements(By.XPATH,"(//span[.='Material Number']/../../../../../../../tbody/tr)[1]"
                                               "//td//lightning-primitive-cell-checkbox")[0].click()

            # Click the next button
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//button[.='Next']")))
            self.driver.find_elements(By.XPATH, "//button[.='Next']")[
                len(self.driver.find_elements(By.XPATH, "//button[.='Next']")) - 1].click()

            # Calculate Price Deduction
            current_price = float(self.driver.find_elements(By.XPATH, f"//lightning-base-formatted-text"
                                                                      f"[.='{current_material}']/../../../../../td")
                                  [7].text.split(" ")[1].replace(",",""))
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

        # Click "Genreate" Invoice Button *rolls eyes harder*
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
        pass

manager = SalesforceManager()
manager.login()
manager.send_replacement_invoice("00057550", "102000722", 700)
