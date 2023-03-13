import sys

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
sys.path.append(ruta_interprete)

import powerfactory as pf
import pandas as pd
import funciones

import importlib

importlib.reload(funciones)

#variable = ['e:loc_name','m:P:bus1', 'm:Q:bus1', 'c:loading']

app=pf.GetApplication()
script = app.GetCurrentScript()
#extraigo todos los objetos dentro de el set lineas
lineas = script.lineas.All()

#acceder al comando flujo de potencia
ldf = app.GetFromStudyCase('ComLdf')

errorFlujo = ldf.Execute()

#contenido = []

if errorFlujo == 0:
    contenido = [funciones.resultados(o, ['e:loc_name','m:P:bus1', 'm:Q:bus1', 'c:loading']) for o in lineas]
    #for linea in lineas:
    #    contenido.append(funciones.resultados(linea, ['e:loc_name','m:P:bus1', 'm:Q:bus1', 'c:loading']))

resultados = pd.DataFrame(contenido, columns = ['nombre linea', 'Potencia activa', 'Potencia reactiva', 'Cargabilidad'])

#filtro = resultados['Potencia activa'] == resultados['Potencia activa'].max()

app.PrintPlain(resultados)

#resultados.to_excel('C:\\Users\\56965\\Documents\\python\\digsilent\\capacitacionBBOCH\\resultados.xlsx')
#resultados.to_excel(r'C:\Users\56965\Documents\python\digsilent\capacitacionBBOCH\resultados.xlsx')
