from tkinter import *
from form_base import FormBase


def add_menu_bar(self, start_row: int, color="brown") -> int:
    self: FormBase
    current_row = start_row + 1

    # Receive Package
    Button(self.base, text='Receive', bg=color, fg='white', command=self.open_receive_form).grid(row=current_row,
                                                                                                   column=0,
                                                                                                   sticky=W + E)
    Button(self.base, text='Repair', bg=color, fg='white', command=self.open_repair_form).grid(row=current_row,
                                                                                                 column=1, sticky=W + E)
    current_row += 1

    # Case Search
    Label(self.base, text="Case Number", font=("bold", 10)).grid(row=current_row, column=0)
    self.case_number = Entry(self.base)
    self.case_number.grid(row=current_row, column=1)

    Button(self.base, text='Search', bg=color, fg='white', command=self.form_search_case).grid(row=current_row,
                                                                                                 column=2, sticky=W)
    current_row += 1

    return current_row
