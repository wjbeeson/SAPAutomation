from tkinter import *
import form_utilities
from inbound_form_printer import InboundFormPrinter
from inbound_label_printer import InboundLabelPrinter
from salesforce import *
from sap_manager import SapManager
from form_menu_bar import add_menu_bar


class InboundPackageForm:
    def __init__(self, package_info: InboundPackageInfo = InboundPackageInfo()):
        # Start Driver
        chrome_options = Options()
        chrome_options.add_argument("--disable-notifications")
        # chrome_options.add_argument('--headless=new')
        self.driver = webdriver.Chrome(options=chrome_options)

        for method in filter(callable, form_utilities.__dict__.values()):
            setattr(self, method.__name__, method)
        add_menu_bar(self, 0)

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
        time = datetime.now().strftime("%m/%d")
        self.date_received.insert(0, time)

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
        Label(base, text=f"ZT01 Number", font=("bold", 10)).grid(row=16)
        self.zt01_number = Entry(base)
        self.zt01_number.grid(row=16, column=1)

        Label(base, text=f"VL01N Number", font=("bold", 10)).grid(row=18)
        self.vl01n_number = Entry(base)
        self.vl01n_number.grid(row=18, column=1)

        Button(base, text='Print Form', bg='brown', fg='white', command=self.print_form).grid(row=19, column=0,
                                                                                              columnspan=3,
                                                                                              sticky=W + E)
        Button(base, text='Clear', bg='white', fg='brown', command=self.reset_form).grid(row=19, column=3,
                                                                                         sticky=W + E)
        Button(base, text='Print Label', bg='brown', fg='white', command=self.print_label).grid(row=20, column=0,
                                                                                                columnspan=3,
                                                                                                sticky=W + E)
        Button(base, text='Send Nano/Lite Email', bg='brown', fg='white', pady=35,
               command=self.send_nano_lite_replacement_email).grid(row=21, column=0,
                                                                   columnspan=4, sticky=E)

        self.error_text = Label(base, text=f"", font=("bold", 10), foreground="red")
        self.error_text.grid(row=100, column=0, columnspan=4, sticky=W + E)
        # it will be used for displaying the registration form onto the window
        self.base = base
        base.mainloop()

    def send_nano_lite_replacement_email(self):
        self.format_form()
        self.sf_manager.send_nano_lite_replacement_email(self.case_number.get())

    def reset_form(self):
        self.case_number.delete(0, END)
        self.product_sku.delete(0, END)
        time = datetime.now().strftime("%m/%d")
        self.date_received.delete(0, END)
        self.date_received.insert(0, time)
        self.first_name.delete(0, END)
        self.last_name.delete(0, END)
        self.box_name.delete(0, END)
        self.house_number.delete(0, END)
        self.street.delete(0, END)
        self.city.delete(0, END)
        self.state.delete(0, END)
        self.zipcode.delete(0, END)
        self.notes.delete(0, END)
        self.zt01_number.delete(0, END)
        self.vl01n_number.delete(0, END)
        self.error_text.config(text="")

    def send_email(self):
        self.format_form()
        self.sf_manager.send_confirmation_email(self.case_number.get())

    def search_case(self):
        self.sf_manager.search_case(self.case_number.get())
        package_info = self.sf_manager.extract_case_info()
        self.set_package_info(package_info)
        self.format_form()

    def validate_form(self, package_info: InboundPackageInfo):
        self.error_text.config(text="")
        if package_info.case_number == "":
            self.error_text.config(text="Case Number is required")
            return False
        if package_info.product_sku == "":
            self.error_text.config(text="Product SKU is required")
            return False
        if len(package_info.state) != 2:
            self.error_text.config(text="State must be two characters")
            return False
        return True

    def receive_package(self):
        package_info = self.get_package_info()

        if not self.validate_form(package_info):
            return

        # Print Label
        self.print_label()

        # Generate zt01 Number
        self.sap_manager.log_zt01_number(package_info)

        # Get zt01 Number
        zt01_number = self.sap_manager.record_zt01_number(package_info.case_number)
        self.zt01_number.delete(0, END)
        self.zt01_number.insert(0, zt01_number)

        # Generate vl01n Number
        self.sap_manager.log_vl01n_number(zt01_number)

        # Get vl01n Number
        vl01n_number = self.sap_manager.record_vl01n_number(zt01_number)
        self.vl01n_number.delete(0, END)
        self.vl01n_number.insert(0, vl01n_number)

        # Print Form
        self.print_form()
        self.send_email()

    def format_form(self):
        package_info = self.get_package_info()
        self.set_package_info(package_info)

    def print_form(self):
        printer = InboundFormPrinter()
        printer.write_workbook(self.get_package_info())
        printer.save_workbook()
        printer.print_workbook()

    def print_label(self):
        printer = InboundLabelPrinter()
        printer.write_workbook(self.get_package_info())
        printer.save_workbook()
        printer.print_workbook()

InboundPackageForm()