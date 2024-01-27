import salesforce
from tkinter import *
from package_info import PackageInfo
import salesforce
import utilities
from salesforce import *


class FormBase:
    def __init__(self):

        self.base = None
        self.driver = None
        self.case_number = None
        self.product_sku = None
        self.date_received = None
        self.first_name = None
        self.last_name = None
        self.box_name = None
        self.street = None
        self.house_number = None
        self.city = None
        self.state = None
        self.zipcode = None
        self.zt01_number = None
        self.vl01n_number = None
        self.notes = None
        self.error_text = None

    def form_search_case(self):
        package_info: PackageInfo = salesforce.extract_case_info(self.driver, self.case_number.get())
        self.set_package_info(package_info)
        package_info: PackageInfo = self.get_package_info()
        self.set_package_info(package_info)

    def open_repair_form(self):
        utilities.switch_to_form(self, "REPAIR")

    def open_receive_form(self):
        utilities.switch_to_form(self, "RECEIVE")





    def get_package_info(self):
        def get_form_element(element):
            try:
                return element.get()
            except Exception:
                return ""
        case_number = get_form_element(self.case_number)
        if len(case_number) != 8:
            preceding_zeros = ""
            for i in range(8 - len(case_number)):
                preceding_zeros += "0"
            case_number = preceding_zeros + case_number

        if get_form_element(self.box_name) == "" or get_form_element(self.box_name) == " ":
            box_name = get_form_element(self.first_name).title() + " " + get_form_element(self.last_name).title()
        else:
            box_name = get_form_element(self.box_name).title()

        package_info = PackageInfo(
            first_name=get_form_element(self.first_name).title(),
            last_name=get_form_element(self.last_name).title(),
            box_name=box_name,
            street=get_form_element(self.street).title(),
            house_number=get_form_element(self.house_number),
            city=get_form_element(self.city).title(),
            state=get_form_element(self.state).upper().replace(".", ""),
            zipcode=get_form_element(self.zipcode),
            date_received=get_form_element(self.date_received),
            case_number=case_number,
            product_sku=get_form_element(self.product_sku),
            zt01_number=get_form_element(self.zt01_number),
            vl01n_number=get_form_element(self.vl01n_number),
            notes=get_form_element(self.notes).lower()
        )
        return package_info


    def set_package_info(self, package_info: PackageInfo):
        def set_form_element(element, value):
            try:
                if value is not None:
                    element.delete(0, END)
                    element.insert(0, value)
            except Exception:
                pass

        set_form_element(self.first_name, package_info.first_name)
        set_form_element(self.last_name, package_info.last_name)
        set_form_element(self.box_name, package_info.box_name)
        set_form_element(self.street, package_info.street)
        set_form_element(self.house_number, package_info.house_number)
        set_form_element(self.city, package_info.city)
        set_form_element(self.state, package_info.state)

        set_form_element(self.zipcode, package_info.zipcode)
        set_form_element(self.date_received, package_info.date_received)
        set_form_element(self.case_number, package_info.case_number)
        set_form_element(self.product_sku, package_info.product_sku)
        set_form_element(self.zt01_number, package_info.zt01_number)
        set_form_element(self.vl01n_number, package_info.vl01n_number)
        set_form_element(self.notes, package_info.notes)


