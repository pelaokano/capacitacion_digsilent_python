import sys

#ruta del interprete de python que estamos usando
sys.path.append("C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\")

#este script permite ejecutar las contingencias
import powerfactory as pf
import pandas as pd
import time

app = pf.GetApplication()#se crea una instancia de digsilen
scr=app.GetCurrentScript()#se crea una instancia del propio script
actPr=app.GetActiveProject()

lineas = scr.lineas.All()

cc = app.GetFromStudyCase('ComShc')

modeloTC = scr.modelos.GetContents('*.TypCt')[0]
modeloTP = scr.modelos.GetContents('*.TypVt')[0]
modeloRele = scr.modelos.GetContents('*.TypRelay')[0]

dict_rele = {}

for linea in lineas:
    
    cubiculo0 = linea.GetCubicle(0)
    cubiculo1 = linea.GetCubicle(1)

    rele0 = cubiculo0.CreateObject('ElmRelay','Rele0')
    TC0 = cubiculo0.CreateObject('StaCt','TC0')
    TP0 = cubiculo0.CreateObject('StaVt','TP0')

    rele1 = cubiculo1.CreateObject('ElmRelay','Rele1')
    TC1 = cubiculo1.CreateObject('StaCt','TC1')
    TP1 = cubiculo1.CreateObject('StaVt','TP1')

    TC1.typ_id = modeloTC
    TC1.ptapset = 1000
    TC1.stapset = 1

    TC0.typ_id = modeloTC
    TC0.ptapset = 1000
    TC0.stapset = 1

    TP1.typ_id = modeloTP
    TP1.ptapset = 20000
    TP1.stapset = 1

    TP0.typ_id = modeloTP
    TP0.ptapset = 20000
    TP0.stapset = 1

    rele0.typ_id = modeloRele
    rele1.typ_id = modeloRele

    ajustesFF = [(0.8, 0.8, 60, 0), (1.2, 1.2, 60, 0.6), (1.5, 1.5, 60, 1)]
    ajustesFN = [(0.8, 1.6, 60, 0), (1.2, 2.4, 60, 0.6), (1.5, 3, 60, 1)]

    for i in range(1,6):
        #Rele 0
        polyFF0 = rele0.GetContents(f'Ph-Ph Polygonal {i}.RelDispoly')[0]
        timerFF0 = rele0.GetContents(f'Ph-Ph Polygonal {i} Delay.RelTimer')[0]
        
        polyFN0 = rele0.GetContents(f'Ph-E Polygonal {i}.RelDispoly')[0]
        timerFN0 = rele0.GetContents(f'Ph-E Polygonal {i} Delay.RelTimer')[0]
        if i < 4:
            polyFF0.Xmax = ajustesFF[i-1][0] * linea.X1
            polyFF0.Rmax = ajustesFF[i-1][1] * linea.X1
            polyFF0.phi = ajustesFF[i-1][2]
            timerFF0.Tdelay = ajustesFF[i-1][3]
            
            polyFF0.outserv = 0
            timerFF0.outserv = 0

            polyFN0.Xmax = ajustesFN[i-1][0] * linea.X1
            polyFN0.Rmax = ajustesFN[i-1][1] * linea.X1
            polyFN0.phi = ajustesFN[i-1][2]
            timerFN0.Tdelay = ajustesFN[i-1][3]
            
            polyFN0.outserv = 0
            timerFN0.outserv = 0
        else:
            polyFF0.outserv = 1
            timerFF0.outserv = 1
            polyFN0.outserv = 1
            timerFN0.outserv = 1
        
        #Rele 1
        polyFF1 = rele1.GetContents(f'Ph-Ph Polygonal {i}.RelDispoly')[0]
        timerFF1 = rele1.GetContents(f'Ph-Ph Polygonal {i} Delay.RelTimer')[0]
        
        polyFN1 = rele1.GetContents(f'Ph-E Polygonal {i}.RelDispoly')[0]
        timerFN1 = rele1.GetContents(f'Ph-E Polygonal {i} Delay.RelTimer')[0]
        if i < 4:
            polyFF1.Xmax = ajustesFF[i-1][0] * linea.X1
            polyFF1.Rmax = ajustesFF[i-1][1] * linea.X1
            polyFF1.phi = ajustesFF[i-1][2]
            timerFF1.Tdelay = ajustesFF[i-1][3]
            
            polyFF1.outserv = 0
            timerFF1.outserv = 0

            polyFN1.Xmax = ajustesFN[i-1][0] * linea.X1
            polyFN1.Rmax = ajustesFN[i-1][1] * linea.X1
            polyFN1.phi = ajustesFN[i-1][2]
            timerFN1.Tdelay = ajustesFN[i-1][3]
            
            polyFN1.outserv = 0
            timerFN1.outserv = 0
        else:
            polyFF1.outserv = 1
            timerFF1.outserv = 1
            polyFN1.outserv = 1
            timerFN1.outserv = 1
    
    dict_rele[linea] = [rele0, rele1]

for linea in lineas:
    cc.shcobj = linea
    for i in range(1,10):
        cc.ppro = i * 10
        ie = cc.Execute()
        if ie == 0:
            app.PrintPlain(f'linea con falla: {linea.loc_name}, distancia cc {cc.ppro}')
            for key, value in dict_rele.items():
                if value[0].HasAttribute('c:yout') and value[1].HasAttribute('c:yout'):
                    top0 = value[0].GetAttribute('c:yout')
                    top1 = value[1].GetAttribute('c:yout')
                    app.PrintPlain(f'linea: {key.loc_name}, tiempo operacion Rele0: {top0}, tiempo operacion Rele1: {top1}')
            