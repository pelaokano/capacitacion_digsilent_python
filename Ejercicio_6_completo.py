import sys

#ruta del interprete de python que estamos usando
sys.path.append("C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\")

#este script permite ejecutar las contingencias
import powerfactory as pf
import pandas as pd
import time

ruta = 'C:\\Users\\56965\\Documents\\python\\digsilent\\capacitacionBBOCH'

app=pf.GetApplication()#se crea una instancia de digsilen
scr=app.GetCurrentScript()#se crea una instancia del propio script
actPr=app.GetActiveProject()#se crea una instancia del proyecto activo

#crea set de elementos para inyectar potencia, son lineas y barras
def elementosSet(clase,setElem,enServ):
    """
    enServ=1 implica que considera a los elementos que estan en servicio
    enServ=0 implica que considera a todos los elementos en servicio o no
    """
    listaElementos=setElem
    listaElementos=listaElementos.All()
    elementos=[]
    
    if enServ==1:
        elementos = [obj for obj in listaElementos if obj.GetClassName()==str(clase) and obj.outserv==0] 
    elif enServ==0:
        elementos = [obj for obj in listaElementos if obj.GetClassName()==str(clase)] 
    
    return elementos

def fVariableGrafo(elementos,variables,result,grafo):
    for e in elementos:
        for v in variables:
            grafo.AddResVars(oRes,e,str(v))

def fExportarImagen(pizarra,pagina,ruta):
    pizarra.Show(pagina)
    comExp=scr.comando
    comExp.iopt_rd = "png"
    comExp.iopt_savas = 0
    comExp.f=ruta + "\\" + str(pagina.loc_name) + ".png" 
    comExp.Execute ()

nomRes="Quasi-Dynamic Simulation DC.ElmRes"
oRes=app.GetFromStudyCase(nomRes)

#se rescata de un set general los tipos de elementos a monitorear
lineas=elementosSet("ElmLne",scr.sElementos,1)
trafo2=elementosSet("ElmTr2",scr.sElementos,1)
trafo3=elementosSet("ElmTr3",scr.sElementos,1)

trafo=trafo2+trafo3

pizarraGrafica=app.GetFromStudyCase('SetDesktop')

#grafico linea
pagina_linea=pizarraGrafica.GetPage('Quasi_Lineas',1)
oPlot_linea=pagina_linea.GetOrInsertPlot('Curve','VisPlot',1)

#grafico trafo2 y trafo3
pagina_trafo=pizarraGrafica.GetPage('Quasi_trafo',1)
oPlot_trafo=pagina_trafo.GetOrInsertPlot('Curve','VisPlot',1)

#fVariableGrafo(lineas,["c:loading","m:P:bus1","m:Q:bus1"],oRes,oPlot_linea)
#fVariableGrafo(trafo,["c:loading","m:P:bushv","m:Q:bushv"],oRes,oPlot_trafo)

fVariableGrafo(lineas,["c:loading"],oRes,oPlot_linea)
fVariableGrafo(trafo,["c:loading"],oRes,oPlot_trafo)

oPlot_linea.DoAutoScaleX()
oPlot_linea.DoAutoScaleY()

oPlot_linea.isteps=1
oPlot_linea.usedfor="ldf"
oPlot_linea.shw_leg=2

oPlot_trafo.DoAutoScaleX()
oPlot_trafo.DoAutoScaleY()

oPlot_trafo.isteps=1
oPlot_trafo.usedfor="ldf"
oPlot_trafo.shw_leg=2

fExportarImagen(pizarraGrafica,pagina_linea,ruta)
fExportarImagen(pizarraGrafica,pagina_trafo,ruta)


