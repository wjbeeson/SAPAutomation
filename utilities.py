
from salesforce import *
from sap_manager import *


def update_package_info(original: PackageInfo, new: PackageInfo, override=False):
    if override:
        for field in original.__dict__:
            if new.__dict__[field] is not None:
                original.__dict__[field] = new.__dict__[field]
    else:
        for field in original.__dict__:
            if original.__dict__[field] is None:
                original.__dict__[field] = new.__dict__[field]
    return original

def start_externals(form_object, driver, sap_manager):
    form_object.driver = driver
    form_object.sap_manager = sap_manager
    if sap_manager is None:
        sap_manager = SapManager()
        sap_manager.login()
        form_object.sap_manager = sap_manager

    if driver is None:
        # Start Driver
        chrome_options = Options()
        chrome_options.add_argument("--disable-notifications")
        # chrome_options.add_argument('--headless=new')
        form_object.driver = webdriver.Chrome(options=chrome_options)
        login(form_object.driver)