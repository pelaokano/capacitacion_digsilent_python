import sys

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
sys.path.append(ruta_interprete)

import powerfactory as pf
import pandas as pd
from win32com import client

excel = client.Dispatch("Excel.Application")
excel.visible = True
libro = excel.Workbooks.Add()
#libro.Name = 'Flujo de Potencia'
libro.Worksheets[0].Name = 'flujo_de_potencia'
hoja = libro.Worksheets[0]

hoja.Cells(1,1).Value = 'Nombre linea'
hoja.Cells(1,2).Value = 'Potencia activa'
hoja.Cells(1,3).Value = 'Potencia reactiva'
hoja.Cells(1,4).Value = 'Cargabilidad'

app=pf.GetApplication()
script = app.GetCurrentScript()

#extraigo todos los objetos dentro de el set lineas
lineas = script.lineas.All()

#acceder al comando flujo de potencia
ldf = app.GetFromStudyCase('ComLdf')

errorFlujo = ldf.Execute()

if errorFlujo == 0:

    for i, linea in enumerate(lineas):
        nombre = linea.GetAttribute('e:loc_name')
        Pi = linea.GetAttribute('m:P:bus1')
        Qi = linea.GetAttribute('m:Q:bus1')
        cargabilidad = linea.GetAttribute('c:loading')
        
        hoja.Cells(2+i,1).Value = nombre
        hoja.Cells(2+i,2).Value = Pi
        if Pi > 100:
            hoja.Cells(2+i,2).Font.Bold = True
        hoja.Cells(2+i,3).Value = Qi
        hoja.Cells(2+i,4).Value = cargabilidad

