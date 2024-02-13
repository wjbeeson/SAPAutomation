import receive_form
import repair_form
from salesforce import *
from sap_manager import SapManager


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

def switch_to_form(self, form_name):
    form_dict = {
        "RECEIVE": receive_form.ReceiveForm,
        "REPAIR": repair_form.RepairForm,
    }
    driver = self.driver
    sap_manager = self.sap_manager

    package_info = self.get_package_info()
    self.base.destroy()
    form_dict[form_name](driver, sap_manager, package_info)

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