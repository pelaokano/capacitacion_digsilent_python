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

nomRes = "Quasi-Dynamic Simulation DC.ElmRes"
oRes = app.GetFromStudyCase(nomRes)

#lineas = elementosSet('ElmLne', scr.sElementos, 1)
#trafo = elementosSet('ElmTr2', scr.sElementos, 1)

elementos = scr.sElementos.All()

lineas = []
trafo = []

for e in elementos:
    if e.GetClassName() == 'ElmLne' and e.outserv == 0:
        lineas.append(e)
    elif e.GetClassName() == 'ElmTr2' and e.outserv == 0:
        trafo.append(e)

pizarraGrafica = app.GetFromStudyCase('SetDesktop')

pagina_grafo_lineas = pizarraGrafica.GetPage('Quasi_Lineas', 1)
plot_lineas = pagina_grafo_lineas.GetOrInsertPlot('Quasi_Lineas', 'VisPlot',1)

pagina_grafo_trafo = pizarraGrafica.GetPage('Quasi_trafo', 1)
plot_trafo = pagina_grafo_trafo.GetOrInsertPlot('Quasi_trafo', 'VisPlot',1)

fVariableGrafo(lineas, ['c:loading'], oRes, plot_lineas)
fVariableGrafo(trafo, ['c:loading'], oRes, plot_trafo)

#plot_lineas.DoAutoScaleX() 
#plot_trafo.DoAutoScaleX()

pizarraGrafica.Show(plot_lineas)
comando = scr.comando
comando.iopt_rd = "png"
comando.iopt_savas = 0
comando.f = ruta + '\\' + 'grafo_linea.png'
comando.Execute() 

pizarraGrafica.Show(plot_trafo)
comando = scr.comando
comando.iopt_rd = "png"
comando.iopt_savas = 0
comando.f = ruta + '\\' + 'grafo_trafo.png'
comando.Execute() 

pagina_grid = pizarraGrafica.GetPage('Grid', 0)

pizarraGrafica.Show(pagina_grid)
comando = scr.comando
comando.iopt_rd = "png"
comando.iopt_savas = 0
comando.f = ruta + '\\' + 'grid.png'
comando.Execute() 