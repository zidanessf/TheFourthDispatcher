import win32com.client as win32
import os
excel = win32.gencache.EnsureDispatch("Excel.Application")
strFile = excel.GetOpenFilename()
wb = excel.Workbooks.Open(strFile)
wb.SaveAs(os.getcwd() + "\\tmp\\next_day_schedule.xlsx",FileFormat=51)
wb.Close()
excel.Application.Quit()