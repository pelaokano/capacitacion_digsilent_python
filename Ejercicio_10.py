import powerfactory
app = powerfactory.GetApplication()
app.ClearOutputWindow()

#accede a la carpeta de la red
NetFolder=app.GetProjectFolder("netdat")

#accede a las path previamente definidas
allPaths=NetFolder.GetContents("*.SetPath",1)

#accede a la pizara grafica de digsilent
#Returns the currently active Graphics Board.
oGraph = app.GetGraphicsBoard()


for eachPath in allPaths:
  oViPage = oGraph.GetPage('TD_'+eachPath.loc_name,1)
  oVi = oViPage.GetOrInsertPlot('km','VisPlottz',1)
  oVi.pPath=eachPath
  #direccion del calculo
  oVi.iopt_dia="fwd"
  for index,eachRelay in enumerate(eachPath.AllProtectionDevices(0)):
    #AllProtectionDevices: Returns all protection devices in the path definition for a given direction.
    #parametros:
    #0 Return devices in forward direction.
    #1 Return devices in reverse direction.
    
    oVi.AddRelay(eachRelay,index+1,1,50)
    #Adds a relay to the plot and optionally sets the drawing style.
    #parametros:
    #relay The protection device to be added
    #colour (optional)
    #style (optional)

  oVi.DoAutoScaleX()
  oVi.DoAutoScaleY()
  
  oVi2= oViPage.GetOrInsertPlot('Sweep','VisPlottz',1)
  oVi2.pPath=eachPath
  #dirección del calculo
  oVi2.iopt_dia="fwd"
  #metodo calculo del diagrama, por defecto kilometrico
  oVi2.iopt_mod="iec"
  for index,eachRelay in enumerate(eachPath.AllProtectionDevices(0)):
    oVi2.AddRelay(eachRelay,index+1,1,50)
  oVi2.DoAutoScaleX()
  oVi2.DoAutoScaleY()
  
  oGraph.Close()
  oVi2.CreateObject("ElmRes","Results")
  oGraph.Show()

#se accede al asistente grafico de protecciones y se ejecuta el comando de short circuit sweep
graphAssis=app.GetFromStudyCase("ComProtgraphic")

#que accion ejecuta el comando: Update diagrams using a short circuit sweep
graphAssis.iopt_action=2
app.PrintPlain(graphAssis)
#se ejecuta el comando
graphAssis.Execute()