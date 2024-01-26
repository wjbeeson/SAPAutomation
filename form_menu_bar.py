from tkinter import *

import salesforce
from salesforce import *
from sap_manager import SapManager, InboundPackageInfo
from inbound_form_printer import InboundFormPrinter
from inbound_label_printer import InboundLabelPrinter
from datetime import datetime

def add_menu_bar(self, start_row: int):
    current_row = start_row + 1

    # Case Search
    Label(self.self.base, text="Case Number", font=("bold", 10)).grid(row=current_row, column=0)
    self.case_number = Entry(self.self.base)
    self.case_number.grid(row=current_row, column=1)

    Button(self.self.base, text='Search', bg='brown', fg='white', command=self.form_search_case).grid(row=current_row, column=2, sticky=W)
    current_row += 1


