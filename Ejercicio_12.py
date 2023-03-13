from tkinter import *
import powerfactory
app=powerfactory.GetApplication()
app.ClearOutputWindow()

res1=app.GetFromStudyCase("MyResults.ElmRes")
Line=app.GetCalcRelevantObjects("*.ElmLne")

def calc():
    warr=0
    SC=app.GetFromStudyCase("ComShc")
    SC.iopt_shc="3psc"
    SC.iopt_mde=1 
    try:
        LineStr=Lb.get(Lb.curselection()) 
    except:
        Warr1=Label(text="Parametro vacio:\n Seleccionar una linea").grid(row=0,column=2)
        warr=1
    Pos=Param.get().replace(",",".")

    if Pos=="":
        Warr1=Label(text="Parametro vacio:\n Seleccionar la ubicacion de la falla en %").grid(row=0,column=2)
        warr=1    

    if warr==0:
        for i in Line:
            if i.loc_name==LineStr:
                SC.shcobj=i
                obj=i
        SC.ppro=float(Pos)
        SC.Execute()
        value=obj.__getattr__("m:Ikss:busshc")
        lab_res1=Label(text=str(value)).grid(row=2,column=2)
    else:
        lab_res1=Label(text="Falta un parametro de entrada").grid(row=3,column=0)

    
Panel=Tk()
Panel.title("Calculo de Cortocircuito")
label0=Label(Panel,text="Seleccionar una linea:")
label0.grid(row=0,column=0)

Lb=Listbox(Panel,width=20,height=5)
j=0
for i in Line:
    Lb.insert(j,i.loc_name)
    j=j+1

Lb.grid(row=0,column=1)
label1=Label(Panel,text="Ubicación de la falla en %")
label1.grid(row=1,column=0)
Param=Entry(Panel,width=7)
Param.grid(row=1,column=1)
label2=Label(Panel,text="%")
label2.grid(row=1,column=3)
Button1=Button(Panel,text="Calcular",command=calc)
Button1.grid(row=2,column=0)
lab_res=Label(text="Corriente de cortocircuito:").grid(row=2,column=1)

Panel.mainloop()

