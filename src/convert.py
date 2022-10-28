import win32com.client as win32
excel = win32.gencache.EnsureDispatch("Excel.Application")
strFile = excel.GetOpenFilename()
wb = excel.Workbooks.Open(strFile)
wb.SaveAs(strFile+"x",FileFormat=51)
wb.Close()
excel.Application.Quit()