#Se debe desarrollar un script que trabaje mejor con Excel

#https://medium.com/towards-data-science/automate-excel-with-python-7c0e8c7c6256

#https://towardsdatascience.com/automate-excel-with-python-pivot-table-899eab993966

#https://kahemchu.medium.com/automate-excel-chart-with-python-d7bec97df1e5#eb8b-f75b97c49234

#https://towardsdatascience.com/automatic-download-email-attachment-with-python-4aa59bc66c25

#http://www.icodeguru.com/webserver/Python-Programming-on-Win32/ch09.htm

#https://jpereiran.github.io/articles/2019/06/14/Excel-automation-with-pywin32.html


import win32com.client as wc

xl = wc.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True

wb = xl.Workbooks.Add()
sh = wb.Sheets[1]

sh.Range('A1:A10').Value = [[i] for i in range(10)]