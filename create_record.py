# import Workbook
from openpyxl import Workbook
from openpyxl import load_workbook
# create Workbook object
wb=Workbook()
# select demo.xlsx
sheet=wb.active
# set value
sheet['A1'] = "ID"
sheet['B1'] = "NAME"
sheet['C1'] = "ATTENDANCE"
sheet['D1'] = "TIME"
# save workbook
wb.save("Record.xlsx")
