import sys

sys.path.append('C:\\Program Files\\DIgSILENT\\PowerFactory 2020 SP7\\Python\\3.8\\')
sys.path.append('C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\')

import powerfactory as pf

import requests
import json
import win32com.client
import time

inicio = time.time()

app = pf.GetApplicationExt()
user = app.GetCurrentUser()
proyecto = user.GetContents('39 Bus New England System3.IntPrj', 1)[0]
proyecto.Activate()

carpeta_script = app.GetProjectFolder('script',0)
script=carpeta_script.GetContents('Ejercicio_8.ComPython',0)[0]
#print(script.loc_name)
filtro=script.GetContents('filtro.SetFilt',0)[0]
listaFiltro = filtro.Get()
#print(listaFiltro)
ldf = app.GetFromStudyCase('ComLdf')

#encabezados de peticion 
payload = ""
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:106.0) Gecko/20100101 Firefox/106.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-ES,es;q=0.8,en-US;q=0.5,en;q=0.3",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://infotecnica.coordinador.cl/",
    "Origin": "https://infotecnica.coordinador.cl",
    "Connection": "keep-alive",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-site"
}

filename = 'C:\\Users\\56965\\Documents\\python\\digsilent\\capacitacionBBOCH\\centrales.xlsx'
#crear aplicación Excel
excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
#visible
excel.Visible = True 
#Abrir un excel
libro = excel.Workbooks.Open(filename)
hoja1 = libro.Sheets('Hoja1')

#print(proyecto.loc_name)

i = 2
while hoja1.Range('B' + str(i)).Value != None:
    id_central = str(int(hoja1.Range('B' + str(i)).Value))
    url1 = f"https://api-infotecnica.coordinador.cl/v1/unidades-generadoras/{id_central}/fichas-tecnicas/general"
    response1 = requests.request("GET", url1, data=payload, headers=headers)
    textoJson1 = json.loads(response1.text)
    potencia = float(textoJson1['7246']['valor_texto'].replace(',','.'))
    
    url2 = f"https://api-infotecnica.coordinador.cl/v1/unidades-generadoras/{id_central}"
    response2 = requests.request("GET", url2, data=payload, headers=headers)
    textoJson2 = json.loads(response2.text)
    
    nombre_central = textoJson2['central_nombre']
    
    hoja1.Range('C' + str(i)).Value = nombre_central
    hoja1.Range('D' + str(i)).Value = potencia
    
    nombre_digsilent = str(hoja1.Range('A' + str(i)).Value)
    
    app.PrintPlain(f'{nombre_central}, {potencia}, {id_central}')
    central = app.GetCalcRelevantObjects(f'{nombre_digsilent}.*', includeOutOfService = 0)
    if central[0].HasAttribute('e:pgini'):
        central[0].SetAttribute('e:pgini', potencia)
    i = i + 1

lineas = app.GetCalcRelevantObjects('*.ElmLne', includeOutOfService = 0)
errorFlujo = ldf.Execute()

hoja1.Range('F1').Value = 'Nombre linea'
hoja1.Range('G1').Value = 'Potencia Activa'
hoja1.Range('H1').Value = 'Cargabilidad'

if errorFlujo == 0:
    for i, linea in enumerate(lineas):
        nombre = linea.GetAttribute('e:loc_name')
        Pi = linea.GetAttribute('m:P:bus1')
        #Qi = linea.GetAttribute('m:Q:bus1')
        cargabilidad = linea.GetAttribute('c:loading')
        
        hoja1.Range('F' + str(i + 2)).Value = nombre
        hoja1.Range('G' + str(i + 2)).Value = Pi
        #hoja1.Range('H' + str(i + 1)).Value = Qi
        hoja1.Range('H' + str(i + 2)).Value = cargabilidad
        if cargabilidad > 80:
            hoja1.Range('H' + str(i + 2)).Font.ColorIndex  = 3
        else:
            hoja1.Range('H' + str(i + 2)).Font.ColorIndex  = 1
            

libro.Save()
libro.Close()
excel.Quit()

fin = time.time()

print(fin - inicio)