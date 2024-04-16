import json
import subprocess
import time
from datetime import date
from tkinter import Tk
import win32com.client
import win32com.client
import win32gui
from pywinauto import Application
from pywinauto_recorder.player import *
import win32con

def multiply_symbol(symbol, number):
    final_symbol = ""

    for i in range(number):
        final_symbol += symbol
    return final_symbol

# This is needed because otherwise Tkinter opens up a bunch of empty windows while running. I believe it's because I
# call "TK()" too many times in my code. It's super hacky but it works.
def close_excess_windows():
    while True:
        handle = win32gui.FindWindow(None, r'tk')
        if handle != 0:
            win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
        else:
            break

# This is a class that manages the SAP GUI. It's a bit of a mess because I had to use a lot of trial and error to get
# the pywinauto_recorder to work reliably. None of this would be necessary if I could just use the SAP GUI scripting,
# but that's not an option because the SAP GUI scripting is disabled (thanks, China).
# Also, this should not have been a class. I should have just made it a bunch of functions. My only excuse is that I
# have never used the library before and planned on adding a bunch of fields before I knew what I was doing.
# (I still don't)
class SapManager:
    def __init__(self):
        pass

    # Logs in using keys stored in a JSON file.
    def login(self):
        try:
            subprocess.call("TASKKILL /F /IM saplogon.exe", shell=True)
            time.sleep(1)
        except Exception:
            pass
        exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        app = Application(backend="uia")
        app.start(exe_path)
        time.sleep(4)
        credentials = json.loads(open(r"keys\sap.json").read())

        win32com.client.GetObject('SAPGUI').GetScriptingEngine.OpenConnection(credentials["connection-name"], True)
        with UIPath(u"SAP||Window"):
            with UIPath(u"||Pane"):
                time.sleep(1)
                send_keys(credentials["username"] + "{TAB}")
                send_keys(credentials["password"] + "{ENTER}")

    # This function logs the ZT01 number into SAP. This is basically just the mouse path that the pywinauto_recorder
    # recorded when I manually entered the ZT01 number into SAP.
    def log_zt01_number(self, package_info):
        with UIPath("RegEx: SAP Easy Access  -  User Menu .*||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                # Step 1: go to appropriate window for entry
                click(u"||Edit")
                send_keys("va01""{ENTER}")
                time.sleep(2)

                # Step 1: enter in warehouse info
                send_keys("{TAB}")
                send_keys("{TAB}")
                send_keys("20""{TAB}")
                send_keys("00""{TAB}")
                send_keys("D003""{TAB}")
                send_keys("ARP""{ENTER}")
                time.sleep(2)

        with UIPath(u"Create Inbound RMA: Overview||Window"):
            #  Case Number
            send_keys("{DOWN}")
            send_keys("{DOWN}")

            send_keys(package_info.case_number)
            send_keys("{TAB}")
            send_keys("{TAB}")

            send_keys(str(date.today().strftime("%Y.%m.%d")))
            send_keys("{UP}")
            send_keys("{UP}")

            #  PO Date

            send_keys("10015")
            send_keys("{ENTER}")
            time.sleep(2)

            #  First Name
            send_keys("{TAB}")
            send_keys(package_info.first_name)

            #  Last Name
            send_keys("{TAB}")
            send_keys(package_info.last_name)

            #  Address
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys(package_info.street)

            send_keys("{TAB}")
            send_keys(package_info.house_number)

            send_keys("{TAB}")
            send_keys(package_info.zipcode)

            send_keys("{TAB}")
            send_keys(package_info.city)

            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys(package_info.state)

            send_keys("{TAB}")
            send_keys("{ENTER}")

            # Product Info
            time.sleep(2)
            send_keys("{ENTER}")
            time.sleep(2)
            send_keys("{ENTER}")
            time.sleep(2)
            send_keys(package_info.product_sku)
            send_keys("{TAB}")
            send_keys("1")
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys("2021")
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys("{TAB}")
            send_keys("2")
            send_keys("{TAB}")
            send_keys("{ENTER}")
            time.sleep(2)
            send_keys('^s')
            time.sleep(2)
            send_keys('{F3}')
            time.sleep(2)
            send_keys('{F3}')
            time.sleep(1)

    # This function records the ZT01 number from SAP.
    def record_zt01_number(self, case_number):
        with UIPath("RegEx: SAP Easy Access  -  User Menu .*||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                # Step 1: go to appropriate window for entry
                click(u"||Edit")
                send_keys("va03""{ENTER}")
                time.sleep(2)
            with UIPath(u"Display Sales Order: Initial Screen||Window"):
                send_keys("^a{DELETE}")
                send_keys("{TAB}")
                send_keys(case_number)
                send_keys(multiply_symbol("{TAB}", 5))
                send_keys(multiply_symbol("{ENTER}", 1))
                time.sleep(3)
                send_keys(multiply_symbol("{DOWN}", 20))

                # For some reason, the clipboard doesn't always get the value the first time, so I got mad and just
                # started spamming the clipboard with the value until it finally worked.
                send_keys("^c")
                send_keys("^c")
                send_keys("^c")
                send_keys("^c")
                send_keys("^c")
                time.sleep(2)
                send_keys("^c")
                send_keys("^c")
                send_keys("^c")
                send_keys("^c")
                clipboard_values = Tk().clipboard_get()
                zt01_number = clipboard_values.split("\t")[len(clipboard_values.split("\t")) - 2]
                send_keys("{ESC}")
                time.sleep(2)
                send_keys("{F3}")
                time.sleep(1)
                send_keys("{F3}")
                time.sleep(1)
                send_keys("{F3}")
                time.sleep(2)
                return zt01_number

    # This function logs the VL01N number into SAP.
    def log_vl01n_number(self, zt01_number):
        with UIPath("RegEx: SAP Easy Access  -  User Menu .*||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                click(u"||Edit")
                send_keys("vl01n""{ENTER}")
                time.sleep(2)
        with UIPath(u"Create Outbound Delivery with Order Reference||Window"):
            send_keys("{UP}""{UP}")
            send_keys("^a{DELETE}")
            send_keys("2021")
            send_keys("{TAB}""{TAB}")
            send_keys("^a{DELETE}")
            send_keys(zt01_number)
            send_keys("{ENTER}")
        with UIPath(u"Returns Delivery  Create: Overview||Window"):
            time.sleep(2)
            with UIPath(u"||Pane"):
                click(u"Post Goods Receipt||Button")
            time.sleep(2)
            send_keys("{F3}")
            time.sleep(1)

    # This function records the VL01N number from SAP.
    def record_vl01n_number(self, zt01_number):
        with UIPath("RegEx: SAP Easy Access  -  User Menu .*||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                click(u"||Edit")
                send_keys("va03""{ENTER}")
                time.sleep(2)
            with UIPath(u"Display Sales Order: Initial Screen||Window"):
                send_keys("^a{DELETE}")
                send_keys(zt01_number)
                send_keys("{ENTER}")
                time.sleep(2)
                send_keys("{F5}")
                time.sleep(2)
                send_keys("{DOWN}")
                send_keys("^c")
                clipboard_values = Tk().clipboard_get()
                vl01_number = clipboard_values.split(" ")[2]
                send_keys("{F3}")
                time.sleep(1)
                send_keys("{F3}")
                time.sleep(1)
                send_keys("{F3}")
                time.sleep(1)
                close_excess_windows()
                return vl01_number
