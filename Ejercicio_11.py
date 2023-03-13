import sys

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
#ruta_resultados = 'C:\\Users\\56965\\Documents\\python\\digsilent\\resultados\\'
sys.path.append(ruta_interprete)

#se importa libreria de power factory y del objeto COM para crear objetos de Test Universe de Omicron
import powerfactory as pf
#import win32com.client
import numpy as np

#Se definen las caracteristicas de las fallas que se simularan sobre la linea
tipoFalla = {0:'3f',1:'2f', 2:'1f', 3:'2fg'}
resistencia = [0, 5, 10, 25, 50]
distancia = [1, 50, 99]

#Se definen los tiempos de simulacion, inicio, despeje y apertura de interruptores
t_ini = 1.5
t_despeje = 3
t_abrir = 3
t_stop = 4

#Se crea una instancia de digsilent
app = pf.GetApplication()

app.EchoOff()
#Se crea una instancia de proyecto activo
actPr = app.GetActiveProject()
#Se crea una instancia del propio script
scr = app.GetCurrentScript()
#Se crea una instancia del caso de esudio activo
stCase = app.GetActiveStudyCase()
#Se crea un instancia de objeto de resultados
resultados = stCase.GetContents('Resultado.ElmRes')
#Se crea una instancia del caso de eventos
eventos = stCase.GetContents('eventos.IntEvt')

if len(resultados) > 0:
    for r in resultados:
        r.Delete()
        
resultados = None
resultado = stCase.CreateObject('ElmRes', 'Resultado')

eventos = stCase.GetContents('eventos.IntEvt')

#Si existe el objeto de evento se elimina
if len(eventos) > 0:
    for e in eventos:
        e.Delete()
        
#Se crean los eventos de cortocircuito, despeje de falla y apertura de los interruptores
evento = stCase.CreateObject('IntEvt', 'eventos')

lineas = scr.lineas
lineas = lineas.All()
if lineas[0].GetClassName() == 'ElmLne':
    linea = lineas[0]

generadores = scr.generadores
generadores = generadores.All()
if generadores[0].GetClassName() == 'ElmSym':
    gen = generadores[0]

resultado.AddVariable(gen,'s:outofstep')
resultado.AddVariable(gen,'c:firel')

#s:outofstep
#s:outofstep

evtCC = evento.CreateObject('EvtShc', 'cortocircuito')
evtDes = evento.CreateObject('EvtShc', 'despejar')
evtOpen = evento.CreateObject('EvtSwitch', 'abrir')

#Se hacen ajustes de los eventos
#Distancia de linea
linea.ishclne = 1
linea.fshcloc = 50
#Se ajustan los parametros del evento de cortocircuito
evtCC.p_target = linea
evtCC.htime = 0
evtCC.mtime = 0
evtCC.time = t_ini
evtCC.X_f = 0
evtCC.R_f = 0
evtCC.i_shc = 0

#Se ajustan los parametros del evento de despeje de falla
evtDes.i_shc = 4
evtDes.p_target = linea
#evtDes.time = t_despeje
evtDes.htime = 0
evtDes.mtime = 0
evtDes.i_clearShc = 0

#Se ajustan los parametros del evento de apertura de falla
evtOpen.p_target = linea
evtOpen.htime = 0
evtOpen.mtime = 0
#evtOpen.time = t_abrir
evtOpen.i_switch = 0
evtOpen.i_allph = 1

#Se crean comandos de condiciones iniciales y de simulacion EMT
calIni = app.GetFromStudyCase('ComInc')
runSim = app.GetFromStudyCase('ComSim')

calIni.iopt_sim = 'rms'
calIni.iopt_net = 'sym'
calIni.p_resvar = resultado
calIni.p_event = evento

runSim.tstop = t_stop

# hacer for para recorrer la distancia de falla, resistencia de falla y tipo de falla
app.ResetCalculation()

for t_open in np.arange(1.6, 4, 0.1):
    evtDes.time = t_open
    evtOpen.time = t_open
            
    #Se ejecutan los comandos de condiciones iniciales y de simulacion EMT
    calIni.Execute()
    runSim.Execute()
    
    resultado.Load()
    NumVar=resultado.GetNumberOfColumns() 
    NumVal=resultado.GetNumberOfRows() 
    ColIndex=resultado.FindColumn(gen,'s:outofstep')
    ColIndex2=resultado.FindColumn(gen,'c:firel')
    #app.PrintPlain(NumVar)
    #app.PrintPlain(NumVal)
    #app.PrintPlain(ColIndex)
    
    for row in range(NumVal):
        #app.PrintPlain(f'prueba {t_open}')
        value=resultado.GetValue(row,ColIndex)[1]
        angulo=resultado.GetValue(row,ColIndex2)[1]
        #app.PrintPlain(value)
        if value > 0:
            app.PrintPlain(f'Sistema inestable con tiempo: {t_open}, el angulo del rotor es: {angulo}, nombre linea: {linea.loc_name}')
            break
         
app.EchoOn()         