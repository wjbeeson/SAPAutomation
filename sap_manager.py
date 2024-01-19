import json
import subprocess
import time
from datetime import date
from tkinter import Tk
import win32com.client
import win32com.client
from pywinauto import Application
from pywinauto_recorder.player import *
import usaddress


class InboundPackageInfo:
    def __init__(self, first_name, last_name, street, house_number, city, state, zipcode, case_number, product_sku,
                 box_name, date_received, zt01_number, vl01n_number, notes):
        self.first_name = first_name
        self.last_name = last_name
        self.street = street
        self.house_number = house_number
        self.city = city
        self.state = state
        self.zipcode = zipcode
        self.case_number = case_number
        self.product_sku = product_sku
        self.box_name = box_name
        self.date_received = date_received
        self.zt01_number = zt01_number
        self.vl01n_number = vl01n_number
        self.notes = notes


def multiply_symbol(symbol, number):
    final_symbol = ""
    for i in range(number):
        final_symbol += symbol
    return final_symbol


class SapManager:
    def __init__(self):
        pass

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
        win32com.client.GetObject('SAPGUI').GetScriptingEngine.OpenConnection("sap", True)
        credentials = json.loads(open(r"C:\Users\abu89\PycharmProjects\SAPAutomation\keys\sap.json").read())
        with UIPath(u"SAP||Window"):
            with UIPath(u"||Pane"):
                time.sleep(1)
                send_keys(credentials["username"]+"{TAB}")
                send_keys(credentials["password"]+"{ENTER}")

    def log_zt01_number(self, package_info):
        with UIPath(u"SAP Easy Access  -  User Menu for Daniel||Window"):
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

    def record_zt01_number(self, case_number):
        with UIPath(u"SAP Easy Access  -  User Menu for Daniel||Window"):
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

    def log_vl01n_number(self, zt01_number):
        with UIPath(u"SAP Easy Access  -  User Menu for Daniel||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                # Step 1: go to appropriate window for entry
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

    def record_vl01n_number(self, zt01_number):
        with UIPath(u"SAP Easy Access  -  User Menu for Daniel||Window"):
            with UIPath(u"||Pane->||ComboBox"):
                # Step 1: go to appropriate window for entry
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
                return vl01_number