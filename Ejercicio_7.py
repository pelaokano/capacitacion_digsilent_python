import win32com.client
import pandas as pd
import sys

ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
sys.path.append(ruta_interprete)

#leer un archivo

# filename = 'C:\\Users\\56965\\Documents\\python\\digsilent\\capacitacionBBOCH\\resultados.xlsx'
# #crear aplicación Excel
# excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
# #visible
# excel.Visible = True 
# #Abrir un excel
# libro = excel.Workbooks.Open(filename)
# hoja1 = libro.Sheets('Sheet1')

# nombres_columnas = ['nombre', 'potencia_activa', 'potencia_reactiva']

# contenido = []

# i = 2
# while hoja1.Range('D' + str(i)).Value != None:
    # aux = [hoja1.Range('B' + str(i)).Value, hoja1.Range('C' + str(i)).Value, hoja1.Range('D' + str(i)).Value]
    # contenido.append(aux)
    # i = i + 1

# data = pd.DataFrame(contenido, columns = nombres_columnas)

# #data = pd.DataFrame(contenido)

# print(data)

#######################################################################################################

constWin32 = win32com.client.constants
#constantes https://learn.microsoft.com/es-es/office/vba/api/excel.constants

excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
#visible
excel.Visible = True 
#Crear un nuevo libro
libro = excel.Workbooks.Add()
#nombrar hoja
hoja1 = libro.Sheets('Hoja1')
hoja1.Range('A1').Value = 'Prueba 1'


for i in range(1,11):
    hoja1.Range('A' + str(i)).Value = f'Prueba {i}'
    if i % 2 == 0:
        hoja1.Range('B' + str(i)).Value = (-1)* i * 10
    else:
        hoja1.Range('B' + str(i)).Value = i * 10

hoja1.Columns('B').HorizontalAlignment = constWin32.xlLeft

#copiar valores
libro.Sheets('Hoja1').Range('A1:B10').Copy(Destination = libro.Sheets('Hoja1').Range('C1'))
#hoja1.Columns('D').HorizontalAlignment = constWin32.xlLeft
# hoja1.Range("A:D").Copy()

# #copiar valores a otro libro
# libro2 = excel.Workbooks.Add()
# hoja1_2 = libro2.Sheets('Hoja1')
# hoja1_2.Range("A1").PasteSpecial(Paste=constWin32.xlPasteValues)

hoja1.Range('E1').Value = 'Repetir'
hoja1.Range('E1:E10').FillDown()

# #contantes para insertar https://learn.microsoft.com/en-us/office/vba/api/excel.xlinsertshiftdirection
# hoja1.Rows('3:4').Insert(constWin32.xlShiftDown)
# hoja1.Rows('5:6').Insert(constWin32.xlShiftToRight)

#ancho columna
#hoja1.Columns("A").ColumnWidth = 30

# #color letra y color fondo https://learn.microsoft.com/es-es/office/vba/api/excel.colorindex
# hoja1.Columns('A:A').Font.ColorIndex  = 3 #rojo
# hoja1.Columns('B:B').Interior.ColorIndex = 5 #azul
# hoja1.Range("A2").Font.Size = 30
# hoja1.Range("A1").Font.Bold = True
# hoja1.Range("A1").Font.Name = "Calibri"

#crear una formula
# hoja1.Range('F1').Formula = '=B1/5'
# hoja1.Range('F1:F10').FillDown()

#formato de numeros https://support.microsoft.com/es-es/office/c%C3%B3digos-de-formato-de-n%C3%BAmero-5026bbd6-04bc-48cd-bf33-80f18b4eae68
#hoja1.Columns('B:B').NumberFormat = '#.##" capacitivo";[Rojo]-#.##" inductivo"'

#Graficos https://peltiertech.com/Excel/ChartsHowTo/ResizeAndMoveAChart.html
# chartShape = hoja1.Shapes.AddChart2()
# chart = chartShape.Chart
# chart.SetSourceData(Source=hoja1.Range("A1:B10"))
# chart.Chart.Title.Text = 'Grafico de prueba'
# chart.Parent.Top = 0
# chart.Parent.Left = 500
# chart.Parent.Height = 200 
# chart.Parent.Width = 300