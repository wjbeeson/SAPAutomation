import os
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from sap_manager import InboundPackageInfo


class TemplatePrinter:
	def __init__(self, src):
		self.wb = load_workbook(src)
		# self.ws = self.wb.get_sheet_by_name("Sheet1") # Deprecated
		self.ws = self.wb["Template"]
		self.destination = Path(src).parent / "temp.xlsx"
	
	# Write the value in the cell defined by row_dest+column_dest
	def write_workbook(self, package_info):
		package_info: InboundPackageInfo
		self.ws.sheet_view.showGridLines = False
		self.ws["B1"] = package_info.case_number
		self.ws["B1"].font = Font(size=14, italic=True)
		
		self.ws["B3"] = package_info.zt01_number
		self.ws["B3"].font = Font(size=14, italic=True)
		
		self.ws["B5"] = package_info.vl01n_number
		self.ws["B5"].font = Font(size=14, italic=True)
		
		self.ws["F1"] = (package_info.first_name + " " + package_info.last_name)
		self.ws["F1"].font = Font(size=14, italic=True)
		
		self.ws["F3"] = package_info.box_name
		self.ws["F3"].font = Font(size=14, italic=True)
		
		self.ws["F5"] = package_info.date_received
		self.ws["F5"].font = Font(size=14, italic=True)
		
		self.ws["E10"] = package_info.notes
		self.ws["E10"].font = Font(size=14, italic=True)
		self.ws["E10"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
	
	def save_workbook(self):
		self.wb.save(self.destination)
	
	def print_workbook(self):
		os.startfile(self.destination, 'print')
	

"""
package_info = InboundPackageInfo(
	first_name="First",
	last_name="Last       t",
	box_name="Box",
	street="Street",
	house_number="First",
	city="City",
	state="State",
	zipcode="Zipcode",
	date_received="Date",
	case_number="Case",
	product_sku="SKU",
	zt01_number="zt01",
	vl01n_number="vl01n",
	notes="Notes Notes Notes Notes Notes Notes Notes NotesNotes  Notes Notes Notes Notes Notes Notes "
)

printer = TemplatePrinter("C:\\Users\\abu89\\PycharmProjects\\SAPAutomation\\scripts\\assets\\InboundTemplate.xlsx")
printer.write_workbook(package_info)
printer.save_workbook()
printer.print_workbook()
"""