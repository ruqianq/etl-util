import openpyxl 
from openpyxl.styles import Alignment
wbook=openpyxl.load_workbook("openpyxl_merge_unmerge.xlsx")
sheet=wbook["merge_sample"]
data=sheet['B4'].value
sheet.merge_cells('B4:E4')
sheet['B4']=data
sheet['B4'].alignment = Alignment(horizontal='center')
wbook.save("openpyxl_merge_unmerge.xlsx")
exit()
