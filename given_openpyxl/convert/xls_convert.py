import win32com.client as win32

def convert_xlsx(current_path, extends):
  
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	wb = excel.Workbooks.Open(current_path)

	# extends xls --> xlsx , extends xlsb --> xlsx
	if extends == "xls":
		wb.SaveAs(current_path+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
	else:
		wb.SaveAs(current_path[0:len(current_path)-1]+"x", FileFormat = 51)
	wb.Close()   
	excel.Application.Quit()
