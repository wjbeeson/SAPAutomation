from tkinter import *
from salesforce_autosearch import SalesforceManager
from sap_manager import SapManager, InboundPackageInfo
from template_printer import TemplatePrinter


class InboundPackageForm:
    def __init__(self):
        self.logged_zt01_number = False
        self.recorded_zt01_number = False
        self.logged_vl01n_number = False
        self.recorded_vl01n_number = False
        self.zt01_number = ""
        self.vl01n_number = ""
        self.package_info = None

        sap_manager = SapManager()
        sap_manager.login()
        self.sap_manager = sap_manager

        sf_manager = SalesforceManager()
        sf_manager.login()
        self.sf_manager = sf_manager

        base = Tk()
        base.geometry('500x500')
        base.title("Inbound Package Form")

        # Product Information
        Label(base, text="Product Information", font=("bold", 15)).grid(row=0, column=0, columnspan=3, sticky=W)
        Label(base, text="Case Number", font=("bold", 10)).grid(row=1, column=0)
        self.case_number = Entry(base)
        self.case_number.grid(row=1, column=1)
        Button(base, text='Search', bg='brown', fg='white', command=self.search_case).grid(row=1, column=2, sticky=W)

        Label(base, text="Product SKU", font=("bold", 10)).grid(row=2, column=0)
        self.product_sku = Entry(base)
        self.product_sku.grid(row=2, column=1)

        Label(base, text="Date Received", font=("bold", 10)).grid(row=2, column=2)
        self.date_received = Entry(base)
        self.date_received.grid(row=2, column=3)

        # Customer Information
        Label(base, text="Customer Information", font=("bold", 15)).grid(row=3, column=0, columnspan=3, sticky=W)

        Label(base, text="First Name", font=("bold", 10)).grid(row=4, column=0)
        self.first_name = Entry(base)
        self.first_name.grid(row=4, column=1)

        Label(base, text="Last Name", font=("bold", 10)).grid(row=4, column=2)
        self.last_name = Entry(base)
        self.last_name.grid(row=4, column=3)

        Label(base, text="Box Name", font=("bold", 10)).grid(row=5, column=0)
        self.box_name = Entry(base)
        self.box_name.grid(row=5, column=1)

        # Address Information
        Label(base, text="Address Information", font=("bold", 15)).grid(row=6, column=0, columnspan=3, sticky=W)

        Label(base, text="House Number", font=("bold", 10)).grid(row=7, column=0)
        self.house_number = Entry(base)
        self.house_number.grid(row=7, column=1)

        Label(base, text="Street", font=("bold", 10)).grid(row=7, column=2)
        self.street = Entry(base)
        self.street.grid(row=7, column=3)

        Label(base, text="City", font=("bold", 10)).grid(row=8, column=0)
        self.city = Entry(base)
        self.city.grid(row=8, column=1)

        Label(base, text="State", font=("bold", 10)).grid(row=8, column=2)
        self.state = Entry(base)
        self.state.grid(row=8, column=3)

        Label(base, text="Zipcode", font=("bold", 10)).grid(row=9, column=0)
        self.zipcode = Entry(base)
        self.zipcode.grid(row=9, column=1)

        Label(base, text="Notes", font=("bold", 10)).grid(row=10, column=0)
        self.notes = Entry(base)
        self.notes.grid(row=10, column=1, columnspan=3, rowspan=3, sticky=W + E)

        # Receive Package
        Button(base, text='Receive Package', bg='brown', fg='white', command=self.receive_package).grid(row=14,
                                                                                                        column=0,
                                                                                                        columnspan=4,
                                                                                                        sticky=W + E)
        # Process Logger
        self.step_1 = Label(base, text=f"1. Log ZT01 Number", font=("bold", 10))
        self.step_1.grid(row=15, column=0, columnspan=4, sticky=W)

        self.step_2 = Label(base, text=f"2. Record ZT01 Number", font=("bold", 10))
        self.step_2.grid(row=16, column=0, columnspan=4, sticky=W)

        self.step_3 = Label(base, text=f"3. Log VL01N Number", font=("bold", 10))
        self.step_3.grid(row=17, column=0, columnspan=4, sticky=W)

        self.step_4 = Label(base, text=f"4. Record VL01N Number", font=("bold", 10))
        self.step_4.grid(row=18, column=0, columnspan=4, sticky=W)

        Button(base, text='Print Form', bg='brown', fg='white', command=self.print_form).grid(row=19, column=0,
                                                                                              columnspan=3,
                                                                                              sticky=W + E)
        Button(base, text='Clear', bg='white', fg='brown', command=self.delete_contents).grid(row=19, column=3,
                                                                                              sticky=W + E)
        # it will be used for displaying the registration form onto the window
        self.base = base
        base.mainloop()

    def delete_contents(self):
        self.case_number.delete(0, END)
        self.product_sku.delete(0, END)
        self.date_received.delete(0, END)
        self.first_name.delete(0, END)
        self.last_name.delete(0, END)
        self.box_name.delete(0, END)
        self.house_number.delete(0, END)
        self.street.delete(0, END)
        self.city.delete(0, END)
        self.state.delete(0, END)
        self.zipcode.delete(0, END)
        self.notes.delete(0, END)
        self.step_1.config(text=f"1. Log ZT01 Number")
        self.step_2.config(text=f"2. Record ZT01 Number")
        self.step_3.config(text=f"3. Log VL01N Number")
        self.step_4.config(text=f"4. Record VL01N Number")
        self.logged_zt01_number = False
        self.recorded_zt01_number = False
        self.logged_vl01n_number = False
        self.recorded_vl01n_number = False

    def search_case(self):
        self.sf_manager.search_case(self.case_number.get())
        package_info = self.sf_manager.extract_case_info()
        self.first_name.delete(0, END)
        self.first_name.insert(0, package_info.first_name)

        self.last_name.delete(0, END)
        self.last_name.insert(0, package_info.last_name)

        self.house_number.delete(0, END)
        self.house_number.insert(0, package_info.house_number)

        self.street.delete(0, END)
        self.street.insert(0, package_info.street)

        self.city.delete(0, END)
        self.city.insert(0, package_info.city)

        self.state.delete(0, END)
        self.state.insert(0, package_info.state)

        self.zipcode.delete(0, END)
        self.zipcode.insert(0, package_info.zipcode)

    def receive_package(self):
        case_number = self.case_number.get()
        if len(case_number) != 8:
            preceding_zeros = ""
            for i in range(8 - len(case_number)):
                preceding_zeros += "0"
            case_number = preceding_zeros + case_number
        else:
            case_number = self.case_number.get()

        if self.box_name.get() == "":
            box_name = self.first_name.get() + " " + self.last_name.get()
        else:
            box_name = self.box_name.get()

        self.package_info = InboundPackageInfo(
            first_name=self.first_name.get(),
            last_name=self.last_name.get(),
            box_name=box_name,
            street=self.street.get(),
            house_number=self.house_number.get(),
            city=self.city.get(),
            state=self.state.get(),
            zipcode=self.zipcode.get(),
            date_received=self.date_received.get(),
            case_number=case_number,
            product_sku=self.product_sku.get(),
            zt01_number=self.zt01_number,
            vl01n_number=self.vl01n_number,
            notes=self.notes.get()
        )
        if not self.logged_zt01_number:
            try:
                self.sap_manager.log_zt01_number(self.package_info)
                self.logged_zt01_number = True
                self.step_1.config(text=f"1. Log ZT01 Number: Done")
            except Exception:
                pass

        if not self.recorded_zt01_number:
            try:
                self.zt01_number = self.sap_manager.record_zt01_number(self.package_info.case_number)
                self.package_info.zt01_number = self.zt01_number
                print("zt01_number: " + self.zt01_number)
                self.recorded_zt01_number = True
                self.step_2.config(text=f"2. Record ZT01 Number: {self.zt01_number}")
            except Exception:
                pass

        if not self.logged_vl01n_number:
            try:
                self.sap_manager.log_vl01n_number(self.zt01_number)
                self.logged_vl01n_number = True
                self.step_3.config(text=f"3. Log VL01N Number: Done")
            except Exception:
                pass

        if not self.recorded_vl01n_number:
            try:
                self.vl01n_number = self.sap_manager.record_vl01n_number(self.zt01_number)
                self.package_info.vl01n_number = self.vl01n_number
                print("vl01_number: " + self.vl01n_number)
                self.recorded_vl01n_number = True
                self.step_4.config(text=f"4. Record VL01N Number: {self.vl01n_number}")
                self.print_form()
            except Exception:
                pass

    def print_form(self):
        template_path = r"assets/InboundTemplate.xlsx"
        printer = TemplatePrinter(template_path)
        printer.write_workbook(self.package_info)
        printer.save_workbook()
        printer.print_workbook()