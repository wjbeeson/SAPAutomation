from tkinter import *
from package_info import PackageInfo
from salesforce import *
from form_parts import add_menu_bar
from form_base import FormBase
from sap_manager import SapManager


class RepairForm(FormBase):
    def __init__(self, driver=None, sap_manager=None, package_info: PackageInfo = PackageInfo()):
        super().__init__()
        self.driver = driver
        self.sap_manager = sap_manager
        if sap_manager is None:
            sap_manager = SapManager()
            sap_manager.login()
            self.sap_manager = sap_manager

        if driver is None:
            # Start Driver
            chrome_options = Options()
            chrome_options.add_argument("--disable-notifications")
            # chrome_options.add_argument('--headless=new')
            self.driver = webdriver.Chrome(options=chrome_options)
            login(self.driver)

        self.base = Tk()
        self.base.geometry('500x500')
        self.base.title("Repair Form")
        current_row = add_menu_bar(self, 0)
        current_row += 1

        # Product Information
        Label(self.base, text="Product Information", font=("bold", 15)).grid(row=current_row, column=0, columnspan=3,
                                                                             sticky=W)
        current_row += 1

        Label(self.base, text="Product SKU", font=("bold", 10)).grid(row=current_row, column=0)
        self.product_sku = Entry(self.base)
        self.product_sku.grid(row=current_row, column=1)

        Label(self.base, text="Date Received", font=("bold", 10)).grid(row=current_row, column=2)
        self.date_received = Entry(self.base)
        self.date_received.grid(row=current_row, column=3)
        time = datetime.now().strftime("%m/%d")
        self.date_received.insert(0, time)
        current_row += 1

        # Customer Information
        Label(self.base, text="Customer Information", font=("bold", 15)).grid(row=current_row, column=0, columnspan=3,
                                                                              sticky=W)
        current_row += 1

        Label(self.base, text="First Name", font=("bold", 10)).grid(row=current_row, column=0)
        self.first_name = Entry(self.base)
        self.first_name.grid(row=current_row, column=1)

        Label(self.base, text="Last Name", font=("bold", 10)).grid(row=current_row, column=2)
        self.last_name = Entry(self.base)
        self.last_name.grid(row=current_row, column=3)
        current_row += 1

        Label(self.base, text="Box Name", font=("bold", 10)).grid(row=current_row, column=0)
        self.box_name = Entry(self.base)
        self.box_name.grid(row=current_row, column=1)
        current_row += 1

        # Address Information
        Label(self.base, text="Address Information", font=("bold", 15)).grid(row=current_row, column=0, columnspan=3,
                                                                             sticky=W)
        current_row += 1

        Label(self.base, text="House Number", font=("bold", 10)).grid(row=current_row, column=0)
        self.house_number = Entry(self.base)
        self.house_number.grid(row=current_row, column=1)

        Label(self.base, text="Street", font=("bold", 10)).grid(row=current_row, column=2)
        self.street = Entry(self.base)
        self.street.grid(row=current_row, column=3)
        current_row += 1

        Label(self.base, text="City", font=("bold", 10)).grid(row=current_row, column=0)
        self.city = Entry(self.base)
        self.city.grid(row=current_row, column=1)

        Label(self.base, text="State", font=("bold", 10)).grid(row=current_row, column=2)
        self.state = Entry(self.base)
        self.state.grid(row=current_row, column=3)
        current_row += 1

        Label(self.base, text="Zipcode", font=("bold", 10)).grid(row=current_row, column=0)
        self.zipcode = Entry(self.base)
        self.zipcode.grid(row=current_row, column=1)
        current_row += 1

        Label(self.base, text="Notes", font=("bold", 10)).grid(row=current_row, column=0)
        self.notes = Entry(self.base)
        self.notes.grid(row=current_row, column=1, columnspan=3, rowspan=3, sticky=W + E)
        current_row += 3

        # Receive Package
        Button(self.base, text='Receive Package', bg='brown', fg='white', command=self.receive_package).grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        Button(self.base, text='Send Nano/Lite Email', bg='white', fg='brown',
               command=self.send_nano_lite_replacement_email).grid(row=current_row, column=3, sticky=W + E)
        current_row += 1

        self.error_text = Label(self.base, text=f"", font=("bold", 10), foreground="red")
        current_row += 1

        self.error_text.grid(row=current_row, column=0, columnspan=4, sticky=W + E)
        # it will be used for displaying the registration form onto the window
        self.set_package_info(package_info)
        self.base.mainloop()

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

    def validate_form(self, package_info: PackageInfo):
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

