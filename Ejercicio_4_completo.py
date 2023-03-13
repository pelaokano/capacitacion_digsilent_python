import sys

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
sys.path.append(ruta_interprete)

#se importa libreria de power factory y del objeto COM para crear objetos de Test Universe de Omicron
import powerfactory as pf
from win32com import client

variables = {'Nombre elemento':'b:loc_name', 
'Potencia activa':'m:P:bus1', 
'Potencia reactiva':'m:Q:bus1', 
'Cargabilidad':'c:loading'}

excel = client.Dispatch("Excel.Application")
excel.visible=True
libro = excel.Workbooks.Add()
libro.Worksheets[0].Name="Report.LDF"
hoja = libro.Worksheets[0]

hoja.Cells(1,1).Value = 'Nombre linea'
hoja.Cells(1,2).Value = 'Pi'
hoja.Cells(1,3).Value = 'Qi'

app=pf.GetApplication()
script = app.GetCurrentScript()

lineas=script.lineas.All()
ldf=app.GetFromStudyCase('ComLdf')
ldf.Execute()

for i, linea in enumerate(lineas):
    hoja.Cells(i+2,1).Value = linea.loc_name
    hoja.Cells(i+2,2).Value = linea.GetAttribute('m:P:bus1')
    hoja.Cells(i+2,3).Value = linea.GetAttribute('m:Q:bus1')