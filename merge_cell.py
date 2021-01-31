import openpyxl 

# wbook=openpyxl.load_workbook("openpyxl_merge_unmerge.xlsx")
# sheet=wbook["merge_sample"]

def merge(sheet, start_cell, end_cell): 
  data=sheet['B4'].value
  sheet.merge_cells('B4:E4')
  sheet['B4']=data
#   wbook.save("openpyxl_merge_unmerge.xlsx")
  return sheet
