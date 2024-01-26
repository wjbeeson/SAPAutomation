from tkinter import *
import salesforce
from sap_manager import InboundPackageInfo


def form_search_case(self):
    return salesforce.extract_case_info(self.driver, self.case_number.get())


def get_package_info(self):
    case_number = self.case_number.get()
    if len(case_number) != 8:
        preceding_zeros = ""
        for i in range(8 - len(case_number)):
            preceding_zeros += "0"
        case_number = preceding_zeros + case_number
    else:
        case_number = self.case_number.get()

    if self.box_name.get() == "" or self.box_name.get() == " ":
        box_name = self.first_name.get().title() + " " + self.last_name.get().title()
    else:
        box_name = self.box_name.get()

    package_info = InboundPackageInfo(
        first_name=self.first_name.get().title(),
        last_name=self.last_name.get().title(),
        box_name=box_name,
        street=self.street.get().title(),
        house_number=self.house_number.get(),
        city=self.city.get().title(),
        state=self.state.get().upper().replace(".", ""),
        zipcode=self.zipcode.get(),
        date_received=self.date_received.get(),
        case_number=case_number,
        product_sku=self.product_sku.get(),
        zt01_number=self.zt01_number.get(),
        vl01n_number=self.vl01n_number.get(),
        notes=self.notes.get().lower()
    )
    return package_info


def set_package_info(self, package_info: InboundPackageInfo):
    def set_form_element(element, value):
        if value is not None:
            element.delete(0, END)
            element.insert(0, value)

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


def update_package_info(original: InboundPackageInfo, new: InboundPackageInfo, override=False):
    if override:
        for field in original.__dict__:
            if new.__dict__[field] is not None:
                original.__dict__[field] = new.__dict__[field]
    else:
        for field in original.__dict__:
            if original.__dict__[field] is None:
                original.__dict__[field] = new.__dict__[field]
    return original
