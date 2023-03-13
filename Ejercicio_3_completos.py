import sys
import importlib
import funciones

importlib.reload(funciones)

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
sys.path.append(ruta_interprete)

#se importa libreria de power factory y del objeto COM para crear objetos de Test Universe de Omicron
import powerfactory as pf
import pandas as pd

app=pf.GetApplication()
script = app.GetCurrentScript()

lineas=script.lineas.All()
ldf=app.GetFromStudyCase('ComLdf')
ldf.Execute()

contenido = [funciones.var_resultado(['e:loc_name', 'm:P:bus1', 'm:Q:bus1'], linea) for linea in lineas ]

#app.PrintPlain(var_resultado(['e:loc_name', 'm:P:bus1', 'm:Q:bus1'], lineas[0]))
app.PrintPlain(contenido)


# contenido = []

# for linea in lineas:
#     nombre = linea.GetAttribute('e:loc_name')
#     Pi = linea.GetAttribute('m:P:bus1')
#     Qi = linea.GetAttribute('m:Q:bus1')

#     contenido.append([nombre, Pi, Qi])

app.PrintPlain(contenido)
data = pd.DataFrame(contenido, columns = ['nombre linea', 'Pi', 'Qi'])

app.PrintPlain(data[data['Pi'] >= 100])

