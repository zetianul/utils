import win32com.client as win32
import os

excel = win32.gencache.EnsureDispatch('Excel.Application')
#Before saving the file set DisplayAlerts to False to suppress the warning dialog:
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('./test1.xlsx'), False, False, None, '2102' )
# refer https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb214129(v=office.12)?redirectedfrom=MSDN
# FileFormat = 51 is for .xlsx extension
wb.SaveAs(os.path.abspath('./test2.xlsx'), 51, '')                                               
wb.Close() 
excel.Application.Quit()