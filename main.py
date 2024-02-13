from tkinter import *
import receive_form
import repair_form


def start_receive_form():
    base.destroy()
    receive_form.ReceiveForm()

def start_repair_form():
    base.destroy()
    repair_form.RepairForm()


base = Tk()
base.geometry('500x500')
base.title("Start Form")
Button(base, text='Receive Form', bg='brown', fg='white', command=start_receive_form, font=("Arial", 40)).pack(expand=True,
                                                                                                     fill=BOTH, padx=25,
                                                                                                     pady=25)
Button(base, text='Repair Form', bg="blue", fg='white', command=start_repair_form, font=("Arial", 40)).pack(expand=True,
                                                                                                     fill=BOTH, padx=25,
                                                                                                     pady=25)
base.mainloop()
