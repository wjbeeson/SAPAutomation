import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import usaddress
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from selenium.webdriver.support.ui import Select
from package_info import PackageInfo
import sku_def_class


def login(driver):
    url = r"https://login.salesforce.com/"
    driver.get(url)
    credentials = json.loads(open(r"keys\salesforce.json").read())
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='username']"))).send_keys(
        credentials["username"])
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='password']"))).send_keys(
        credentials["password"])
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='submit']"))).click()


def execute_on_web_element(driver, xpath, action, index=-1):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    if index == -1:
        index = len(driver.find_elements(By.XPATH, xpath)) - 1
    element = driver.find_elements(By.XPATH, xpath)[index]
    return action(element)


def generate_payment_link(driver, case_number, material_sku, price):
    search_case(driver, case_number)
    execute_on_web_element(driver, 'html', lambda f: f.send_keys(Keys.END))
    time.sleep(1)
    execute_on_web_element(driver, 'html', lambda f: f.send_keys(Keys.HOME))
    account_name = execute_on_web_element(driver, "((//span[.='Account Information']/../../../div/div/slot"
                                                  "/records-record-layout-row)[1]//slot//records-record-layout-"
                                                  "item[@field-label='Account Name']//span[contains(.,'@')"
                                                  "])[2]", lambda f: f.text)

    # Click new order button
    execute_on_web_element(driver, "//article[@aria-label='Internal Work Order']//lst-related-list-view-"
                                   "manager//lst-common-list-internal//lst-list-view-manager-header//div//div//"
                                   "div//div//div//h2//a", lambda f: f.click())

    # Click new button
    execute_on_web_element(driver, "//button[@name='New']", lambda f: f.click())

    # Click Replacement Button
    execute_on_web_element(driver, "(//label[@class='slds-radio topdown-radio-container slds-clearfix'])[7]//span"
                           , lambda f: f.click())

    # Click Next Button
    execute_on_web_element(driver, "//div[@class='inlineFooter']//button", lambda f: f.click(), 1)

    # Select Current Step
    execute_on_web_element(driver, "//button[@aria-label='Current step, --None--']", lambda f: f.click())
    execute_on_web_element(driver, "//button[@aria-label='Current step, --None--']/../..//div[@role='listbox']//"
                                   "lightning-base-combobox-item", lambda f: f.click(), 2)

    # Select Region
    execute_on_web_element(driver, "//button[@aria-label='Region, --None--']", lambda f: f.click())
    execute_on_web_element(driver, "//button[@aria-label='Region, --None--']/../..//div["
                                   "@role='listbox']//lightning-base-combobox-item", lambda f: f.click(), 2)

    # Select account name
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Accounts...']")))
    combo_box = driver.find_elements(By.XPATH, "//input[@placeholder='Search Accounts...']")[0]
    _ = combo_box.location_once_scrolled_into_view
    combo_box.click()
    combo_box.send_keys(account_name)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{account_name}']/../../../../div["
                                                  f"@role='listbox']//lightning-base-combobox-item//span"
                                                  f"//lightning-icon[@icon-name='utility:search']")))
    driver.find_elements(By.XPATH,
                         f"//input[@data-value='{account_name}']/../../../../div["
                         f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

    # Click Save Button
    execute_on_web_element(driver, "//button[@name='SaveEdit']", lambda f: f.click())

    # Click invoice button
    execute_on_web_element(driver, "//span[@title='Proforma Invoice']/..", lambda f: f.click())

    # Click new button
    execute_on_web_element(driver, "(//button[@name='New'])[2]", lambda f: f.click())

    # Click Overseas Button
    execute_on_web_element(driver, f"//span[.='Overseas']", lambda f: f.click())

    # Click Next Button
    execute_on_web_element(driver, f"//span[.='Next']/..", lambda f: f.click())

    # Select account name
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Accounts...']")))
    combo_box = driver.find_elements(By.XPATH, "//input[@placeholder='Search Accounts...']")[0]
    _ = combo_box.location_once_scrolled_into_view
    combo_box.click()
    combo_box.send_keys(account_name)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{account_name}']/../../../../div["
                                                  f"@role='listbox']//lightning-base-combobox-item//span"
                                                  f"//lightning-icon[@icon-name='utility:search']")))
    driver.find_elements(By.XPATH,
                         f"//input[@data-value='{account_name}']/../../../../div["
                         f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

    # Select Price Book
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Price Books...']")))
    combo_box = driver.find_elements(By.XPATH, "//input[@placeholder='Search Price Books...']")[0]
    _ = combo_box.location_once_scrolled_into_view
    combo_box.click()
    price_book = "North America Price Book"
    combo_box.send_keys("North America Price Book")

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//input[@data-value='{price_book}']/../../../../div["
                                                  f"@role='listbox']//lightning-base-combobox-item//span"
                                                  f"//lightning-icon[@icon-name='utility:search']")))
    driver.find_elements(By.XPATH,
                         f"//input[@data-value='{price_book}']/../../../../div["
                         f"@role='listbox']//ul//li//lightning-base-combobox-item")[0].click()

    # Click Save Button
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//button[@name='SaveEdit']")))
    time.sleep(2)
    driver.find_elements(By.XPATH, "//button[@name='SaveEdit']")[0].click()
    time.sleep(2)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@role='alertdialog']//div//div//div//span//a//div")))

    def get_active_bar():
        while True:
            driver.refresh()
            time.sleep(3)
            # Check to see if add product is active
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
                                                          "[@class='slds-col slds-no-flex slds-grid slds-"
                                                          "grid_vertical-align-center horizontal "
                                                          "actionsContainer']//div//"
                                                          "runtime_platform_actions-actions-ribbon//ul//li")))
            result = driver.find_elements(By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
                                                    "[@class='slds-col slds-no-flex slds-grid slds-"
                                                    "grid_vertical-align-center horizontal "
                                                    "actionsContainer']//div//"
                                                    "runtime_platform_actions-actions-ribbon//ul//li")
            time.sleep(3)
            check = driver.find_elements(By.XPATH, "//div[.='Proforma Invoice']/../../../..//div"
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
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                      "down-trigger slds-dropdown-trigger_click slds-is-open']")))
        driver.find_elements(By.XPATH,
                             "//runtime_platform_actions-action-renderer[@title='Add Product']")[
            len(driver.find_elements(By.XPATH, "//runtime_platform_actions-action-renderer"
                                               "[@title='Add Product']"))].click()
    else:
        # Click button if in active bar
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"(//button[.='Add Product'])[1]")))
        driver.find_elements(By.XPATH, f"//button[.='Add Product']")[
            len(driver.find_elements(By.XPATH, f"//button[.='Add Product']")) - 1].click()

    # Row is 0 based index. Chooses which material line item to add
    def add_material(current_material, price):
        # Click Add Product Button in window
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//lightning-button[@variant='brand']//button[@class='slds-button "
                           "slds-button_brand'][.='Add Product']")))
        driver.find_elements(By.XPATH, "//lightning-button[@variant='brand']//button[@class='slds-button "
                                       "slds-button_brand'][.='Add Product']")[
            len(driver.find_elements(By.XPATH,
                                     "//lightning-button[@variant='brand']//button[@class='slds-button "
                                     "slds-button_brand'][.='Add Product']")) - 1].click()

        # Send material to combo box
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search by Material']")))
        search_box = driver.find_elements(By.XPATH, "//input[@placeholder='Search by Material']")[
            len(driver.find_elements(By.XPATH, "//input[@placeholder='Search by Material']")) - 1]
        search_box.click()
        search_box.send_keys(current_material)

        # Click search button
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[.='Search']")))
        driver.find_elements(By.XPATH, "//button[.='Search']")[
            len(driver.find_elements(By.XPATH, "//button[.='Search']")) - 1].click()

        # Select material
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "(//span[.='Material Number']/../../../../../../../tbody/tr)"
                                                      "[1]""//td//lightning-primitive-cell-checkbox")))
        driver.find_elements(By.XPATH, "(//span[.='Material Number']/../../../../../../../tbody/tr)[1]"
                                       "//td//lightning-primitive-cell-checkbox")[0].click()

        # Click the next button
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[.='Next']")))
        driver.find_elements(By.XPATH, "//button[.='Next']")[
            len(driver.find_elements(By.XPATH, "//button[.='Next']")) - 1].click()

        # Calculate Price Deduction
        current_price = float(driver.find_elements(By.XPATH, f"//lightning-base-formatted-text"
                                                             f"[.='{current_material}']/../../../../../td")
                              [7].text.split(" ")[1].replace(",", ""))
        discount = current_price - price

        # Click the discount field
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"(//lightning-base-formatted-text[.='{material_sku}']"
                                                      f"/../../../../../td)"
                                                      f"[9]//lightning-primitive-cell-factory//span//button")))
        driver.find_elements(By.XPATH,
                             f"(//lightning-base-formatted-text[.='{current_material}']/../../../../../td)[9]"
                             f"//lightning-primitive-cell-factory//span//button")[0].click()

        # Send discount
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//input[@name='dt-inline-edit-currency']")))
        discount_field = driver.find_elements(By.XPATH, f"//input[@name='dt-inline-edit-currency']")[0]
        discount_field.send_keys(discount)
        discount_field.send_keys(Keys.ENTER)

    add_material(material_sku, price)
    add_material("901030163", 0)  # This is the SKU for 2nd level labor

    # Click "Comfirm" button *rolls eyes*
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
    driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()

    # Check to see if payment request button is active
    active_bar = get_active_bar()
    active_bar_len = len(active_bar)
    if active_bar_len < 5:
        # Click Dropdown if not in active bar
        time.sleep(1)
        active_bar[active_bar_len - 1].click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                      "down-trigger slds-dropdown-trigger_click slds-is-open']")))
        driver.find_elements(By.XPATH,
                             "//runtime_platform_actions-action-renderer[@title='Create Payment Request']")[
            0].click()
    else:
        # Click button if in active bar
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"(//button[.='Create Payment Request'])[1]")))
        driver.find_elements(By.XPATH, f"//button[.='Create Payment Request']")[0].click()

    # Click "Comfirm" button *rolls eyes all the way around*
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
    driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()

    # Check to see if generate button is active
    active_bar = get_active_bar()
    active_bar_len = len(active_bar)
    if active_bar_len < 7:
        # Click Dropdown if not in active bar
        active_bar[active_bar_len - 1].click()
        time.sleep(1)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//lightning-button-menu[@class='menu-button-item slds-drop"
                                                      "down-trigger slds-dropdown-trigger_click slds-is-open']")))
        driver.find_elements(By.XPATH,
                             "//runtime_platform_actions-action-renderer[@title='Genreate "
                             "and Send Quotation']")[0].click()
    else:
        # Click button if in active bar
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//button[.='Genreate and Send Quotation']")))
        driver.find_elements(By.XPATH, f"//button[.='Genreate and Send Quotation']")[0].click()

    # Click "Comfirm" button *rolls eyes all the way around*
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//button[.='Comfirm']")))
    driver.find_elements(By.XPATH, f"//button[.='Comfirm']")[0].click()

    # Wait until the PDF Saves
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, f"//div[.='Saving quote PDF']")))

    # Load payment URL by putting it on screen
    driver.find_element(By.TAG_NAME, 'html').send_keys(Keys.END)
    time.sleep(5)

    # Get Link
    payment_page_url = (driver.find_elements(By.XPATH, f'//img[@src="https://autelroboticsusa.my.salesforce.com/img/'
                                                       f'icon/t4v35/custom/custom48_120.png"]/../../../../../../../../'
                                                       f'../../../lst-related-list-view-manager/lst-common-list-internal'
                                                       f'/div/div/lst-primary-display-manager/div/lst-primary-display/'
                                                       f'lst-primary-display-card/lst-customized-template-list/div/'
                                                       f'lst-template-list-item-factory/lst-related-preview-card/article'
                                                       f'/div/div/h3/lst-template-list-field/lst-output-lookup/'
                                                       f'force-lookup/div/records-hoverable-link/div/a')[0]
                        .get_attribute("Href"))

    pass
    driver.get(payment_page_url)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, f"//span[.='Payment Link']/../../../div/div[@class='slds-"
                                                  f"form-element__control']//span//slot//records-formula-output//"
                                                  f"slot//lightning-formatted-text//a")))
    return driver.find_elements(By.XPATH,
                                f"//span[.='Payment Link']/../../../div/div[@class='slds-form-element__control']"
                                f"//span//slot//records-formula-output//slot//"
                                f"lightning-formatted-text//a")[0].get_attribute("href")


def send_nano_lite_replacement_email(driver, case_number, sku_prices, has_care):
    sku_def_dict = {}
    sku_def_dict["102000750"] = sku_def_class.SkuDef("102000750", "EVO Nano+ Premium Bundle", 719)
    sku_def_dict["102000722"] = sku_def_class.SkuDef("102000750", "EVO Lite+ Premium Bundle", 999)
    links = ""
    for sku in list(sku_prices.keys()):
        link = generate_payment_link(driver, case_number, sku, sku_prices[sku])
        sku_def: sku_def_class = sku_def_dict[sku]
        links = links + f"{sku_def.name} [${sku_prices[sku]}, Normally ${sku_def.standard_price}]: {link}\n"

    package_info = extract_case_info(driver, case_number)

    driver.get(f"https://autelroboticsusa.my.salesforce.com/_ui/core/email/author/EmailAuthor")
    template = 'templates/nano_lite_replacement.txt'
    if has_care:
        template = 'templates/autel_care_replacement.txt'
    with open(template, 'r') as file:
        body = file.read()
        body = body.replace("{name}", package_info.first_name + " " + package_info.last_name)
        body = body.replace("{case_number}", package_info.case_number)
        body = body.replace("{links}", links)
        subject = f"Autel Robotics - Case {package_info.case_number} Received"

    # From
    Select(execute_on_web_element(driver, "//select[@id='p26']", lambda f: f)).select_by_index(1)

    # Subject
    execute_on_web_element(driver, "//input[@id='p6']", lambda f: f.send_keys(subject), 0)

    # Body
    execute_on_web_element(driver, "//textarea[@id='p7']", lambda f: f.send_keys(body), 0)

    # Email
    execute_on_web_element(driver, "//textarea[@id='p24']", lambda f: f.send_keys(package_info.account_name), 0)

    # Related To
    Select(execute_on_web_element(driver, "//select[@id='p3_mlktp']", lambda f: f)).select_by_index(15)

    # Case Number
    execute_on_web_element(driver, "//input[@id='p3']", lambda f: f.send_keys(case_number), 0)
    pass
    # Click the send button
    #execute_on_web_element(driver, "//input[@title='Send']", lambda f: f.click(), 0)


def send_confirmation_email(driver, case_number):
    package_info = extract_case_info(driver, case_number)
    driver.get(f"https://autelroboticsusa.my.salesforce.com/_ui/core/email/author/EmailAuthor")

    with open('templates/confirmation_email.txt', 'r') as file:
        body = file.read()
        body = body.replace("{name}", package_info.first_name + " " + package_info.last_name)
        body = body.replace("{case_number}", package_info.case_number)
        subject = f"Autel Robotics - Case {package_info.case_number} Received"

    # From
    Select(execute_on_web_element(driver, "//select[@id='p26']", lambda f: f)).select_by_index(1)

    # Subject
    execute_on_web_element(driver, "//input[@id='p6']", lambda f: f.send_keys(subject), 0)

    # Body
    execute_on_web_element(driver, "//textarea[@id='p7']", lambda f: f.send_keys(body), 0)

    # Email
    execute_on_web_element(driver, "//textarea[@id='p24']", lambda f: f.send_keys(package_info.account_name), 0)

    # Related To
    Select(execute_on_web_element(driver, "//select[@id='p3_mlktp']", lambda f: f)).select_by_index(15)

    # Case Number
    execute_on_web_element(driver, "//input[@id='p3']", lambda f: f.send_keys(case_number), 0)

    # Click the send button
    execute_on_web_element(driver, "//input[@title='Send']", lambda f: f.click(), 0)


def search_case(driver, case_number):
    driver.get(
        f"https://autelroboticsusa.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?str={case_number}#!/fen=500&initialViewMode=detail&str={case_number}")
    case_id = execute_on_web_element(driver, "(//div[@class='pbBody']//tr)[2]//th//a",
                                     lambda f: f.get_attribute("data-seclki"))
    driver.get(f"https://autelroboticsusa.lightning.force.com/lightning/r/Case/{case_id}/view")


def extract_case_info(driver, case_number):
    search_case(driver, case_number)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "//records-record-layout-item[@field-label='Received by']//div//div//div["
                       "@class='slds-form-element__control'][span]//slot[@name='outputField']")))

    case_number = ""
    first_name = ""
    last_name = ""
    house_number = ""
    street = ""
    city = ""
    state = ""
    zipcode = ""
    account_name = ""
    try:
        case_number = driver.find_elements(By.XPATH, "(//records-record-layout-item[@field-label="
                                                     "'Case Number']//div//div//div)[2]")[0].text
        account_name = driver.find_elements(By.XPATH, "//button[@title='Edit Account Name']/..//span//a//span")[0].text
        name = driver.find_elements(By.XPATH,
                                    "//records-record-layout-item[@field-label='Received "
                                    "by']//div//div//div[@class='slds-form-element__control']["
                                    "span]//slot["
                                    "@name='outputField']//lightning-formatted-text")[0].text
        split = name.strip().split(" ")
        first_name = split[0]
        last_name = split[1]
        address = driver.find_elements(
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

    package_info = PackageInfo(
        case_number=case_number,
        first_name=first_name,
        last_name=last_name,
        street=street,
        house_number=house_number,
        city=city,
        state=state,
        zipcode=zipcode,
        account_name=account_name
    )
    return package_info
