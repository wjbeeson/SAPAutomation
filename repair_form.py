from tkinter import *

import utilities
from form_base import FormBase
from form_parts import add_menu_bar
from salesforce import *


class RepairForm(FormBase):
    def __init__(self, driver=None, sap_manager=None, package_info: PackageInfo = PackageInfo()):
        super().__init__()
        self.color = "Blue"
        utilities.start_externals(self, driver, sap_manager)

        self.base = Tk()
        self.base.geometry('500x500')
        self.base.title("Repair Form")
        current_row = add_menu_bar(self, 0, self.color)
        current_row += 1

        # Customer Information
        Label(self.base, text="Customer Information", font=("bold", 15)).grid(row=current_row, column=0, columnspan=3,
                                                                              sticky=W)
        current_row += 1

        Label(self.base, text="Account", font=("bold", 10)).grid(row=current_row, column=0)
        self.account_name = Entry(self.base)
        self.account_name.grid(row=current_row, column=1)
        current_row += 1

        current_row += 1

        Label(self.base, text="First Name", font=("bold", 10)).grid(row=current_row, column=0)
        self.first_name = Entry(self.base)
        self.first_name.grid(row=current_row, column=1)

        Label(self.base, text="Last Name", font=("bold", 10)).grid(row=current_row, column=2)
        self.last_name = Entry(self.base)
        self.last_name.grid(row=current_row, column=3)
        current_row += 1

        Label(self.base, text="Autel Care", font=("bold", 15)).grid(row=current_row, column=0, columnspan=3,
                                                                    sticky=W)
        self.autel_care = IntVar()
        self.autel_care.set(0)
        Checkbutton(self.base, text='Autel Care', variable=self.autel_care, onvalue=1, offvalue=0).grid(row=current_row,
                                                                                                        column=1)

        current_row += 1

        Label(self.base, text="Model", font=("bold", 10)).grid(row=current_row, column=0)
        variable = StringVar(self.base)
        variable.set("Nano+")  # default value
        self.model = OptionMenu(self.base, variable, "Nano+", "Lite+")
        self.model.grid(row=current_row, column=1, sticky=W + E)

        Label(self.base, text="Count", font=("bold", 10)).grid(row=current_row, column=2)
        variable = StringVar(self.base)
        variable.set("1")  # default value
        self.replacement_count = OptionMenu(self.base, variable, "1", "2")
        self.replacement_count.grid(row=current_row, column=3, sticky=W + E)
        current_row += 1

        Button(self.base, text='Send Nano/Lite Replacement Email', bg=self.color, fg="white",
               command=self.send_replacement_email).grid(row=current_row, column=0, columnspan=3, sticky=W + E)
        Button(self.base, text='Clear', bg='white', fg=self.color, command=self.reset_form).grid(row=current_row,
                                                                                                 column=3,
                                                                                                 rowspan=1,
                                                                                                 sticky=W + E + S + N)
        current_row += 1

        self.error_text = Label(self.base, text=f"", font=("bold", 10), foreground="red")
        current_row += 1

        self.error_text.grid(row=current_row, column=0, columnspan=4, sticky=W + E)
        # it will be used for displaying the registration form onto the window
        self.set_package_info(package_info)
        self.base.mainloop()

    def get_replacement_info(self):
        result = {}
        autel_care = self.autel_care.get()
        if autel_care == 0:
            result["102000750"] = 400
            result["102000722"] = 800
            return result

        model = self.model.cget("text")
        if model == "Nano+":
            if self.replacement_count.cget("text") == "1":
                result["102000750"] = 79
            else:
                result["102000750"] = 99
        else:
            if self.replacement_count.cget("text") == "1":
                result["102000722"] = 149
            else:
                result["102000722"] = 159
        pass
        return result

    def send_replacement_email(self):
        sku_prices = self.get_replacement_info()
        package_info = self.get_package_info()
        if not self.validate_form(package_info):
            return
        has_care = (self.autel_care.get() == 1)
        send_nano_lite_replacement_email(self.driver, self.case_number.get(), sku_prices, has_care)

    def reset_form(self):
        self.case_number.delete(0, END)
        self.first_name.delete(0, END)
        self.last_name.delete(0, END)
        self.account_name.delete(0, END)
        self.error_text.config(text="")

    def search_case(self):
        search_case(self.driver, self.case_number.get())
        package_info = extract_case_info(self.driver, self.case_number.get())
        self.set_package_info(package_info)
        self.format_form()

    def validate_form(self, package_info: PackageInfo):
        self.error_text.config(text="")
        if package_info.case_number == "":
            self.error_text.config(text="Case Number is required")
            return False
        if package_info.account_name == "":
            self.error_text.config(text="Account is required")
            return False
        if package_info.account_name.find("@") == -1:
            self.error_text.config(text="Account needs to be in Email Format")
            return False
        return True

    def format_form(self):
        package_info = self.get_package_info()
        self.set_package_info(package_info)
