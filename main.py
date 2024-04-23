from tkinter import *
import receive_form


def start_receive_form():
    base.destroy()
    receive_form.ReceiveForm()


base = Tk()
base.geometry('500x500')
base.title("Start Form")
Button(base, text='Receive Form', bg='brown', fg='white', command=start_receive_form, font=("Arial", 40)).pack(
    expand=True,
    fill=BOTH, padx=25,
    pady=25)
base.mainloop()
