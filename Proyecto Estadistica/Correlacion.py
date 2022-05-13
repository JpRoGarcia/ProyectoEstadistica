import tkinter.constants
from matplotlib.pyplot import *
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.animation as anim
from tkinter import *
from tkinter import messagebox
import numpy as np
from math import *
import pandas as pd
from os import path


# ------------ Metodos Funcionales --------------
def MensajeCorrelacion(valor):
    if valor == 1:
        msg = " Entonces Existe una alineacion lineal Positiva Perfecta"
    elif valor == -1:
        msg = " Entonces Existe una alineacion lineal Negativa Perfecta"
    elif valor == 0:
        msg = " No Existe Asociación Lineal"
    elif valor < 1 and valor > 0.5:
        msg = " Entonces Existe una alineacion lineal Positiva Fuerte"
    elif valor <= 0.5 and valor > 0:
        msg = " Entonces Existe una alineacion lineal Positiva Debil"
    elif valor < 0 and valor >= -0.5:
        msg = " Entonces Existe una alineacion lineal Negativa Debil"
    elif valor < -0.5 and valor > -0.9:
        msg = " Entonces Existe una alineacion lineal Negativa Fuerte"


    return msg

def CorrelacionExcel():
    try:
        nombre = NombreExcel.get()+".xlsx"
        print(nombre)
        Hoja = NombreHoja.get()
        print(Hoja)
        ColX = NombreColUno.get()
        ColY = NombreColDos.get()
        direccion = (path.abspath(nombre))
        datos = pd.read_excel(direccion, sheet_name=Hoja)
        x = datos[ColX].to_numpy()
        y = datos[ColY].to_numpy()
        operador = regresion_lineal(x, y)
        diagramaDispersion(x, y, operador)
        valorCorrelacion = np.corrcoef(x, y)[0, 1]
        messagebox.showinfo("Resultado", "la Correlacion es: " + str(valorCorrelacion) + MensajeCorrelacion(valorCorrelacion))
    except:
        messagebox.showerror("Error", "Error al Ingresar los Datos.")

def CorrelacionManual():
    try:
        Auxx = ArrayUno.get()
        Auxy = ArrayDos.get()
        x = Auxx.split(',')
        y = Auxy.split(',')
        x = np.array(list(map(float, x)))
        y = np.array(list(map(float, y)))
        operador = regresion_lineal(x, y)
        diagramaDispersion(x, y, operador)
        valorCorrelacion = np.corrcoef(x, y)[0, 1]
        messagebox.showinfo("Resultado", "la Correlacion es: "+ str(valorCorrelacion) + MensajeCorrelacion(valorCorrelacion))
    except:
        messagebox.showerror("Error", "Error al Ingresar los Datos.")

def regresion_lineal(x, y):
    n = x.shape[0]
    sumaxy = 0
    sumax = 0
    sumay = 0
    sumax2 = 0

    for i in range(n):
        sumaxy += x[i] * y[i]
        sumax += x[i]
        sumay += y[i]
        sumax2 += x[i] ** 2

    mediax = sumax / n
    mediay = sumay / n

    a1 = ((n * sumaxy) - (sumax * sumay)) / ((n * sumax2) - (sumax ** 2))
    a0 = mediay - a1 * mediax

    yest = a0 + a1 * x

    return yest

def diagramaDispersion(x, y, f):
    W = Tk()
    W.title("Diagrama de Dispersión")
    fig = Figure()
    ax = fig.add_subplot(111)
    ax.scatter(x, y, color="#000000")
    ax.plot(x, f)
    cvs = FigureCanvasTkAgg(fig, W)
    cvs.draw()
    cvs.get_tk_widget().pack(side=TOP, fill=BOTH)
    tlb = NavigationToolbar2Tk(cvs, W)
    tlb.update()
    cvs.get_tk_widget().pack(side=TOP, fill=BOTH)

def clearCor():
    ArrayUno.set("")
    ArrayDos.set("")

def clearExcel():
    NombreExcel.set("")
    NombreHoja.set("")
    NombreColUno.set("")
    NombreColDos.set("")

#---------------------- Diseño ------------------------

#---------------------- Creacion Ventana ------------------------
ventana = Tk()
ventana.title("Coeficiente de Correlación")
ventana.geometry("550x450")
ventana.configure(background="#80DEEA")
ventana.resizable(0,0)

#---------------------- Creacion de Variables Aux ------------------------
C1 = ("#80ea8c")
C2 = ("#FF5050")
C3 = ("#80DEEA")
M = ("#eac180")

AB = 35
HB = 2

ArrayUno = StringVar()
ArrayDos = StringVar()

NombreExcel = StringVar()
NombreHoja = StringVar()
NombreColUno = StringVar()
NombreColDos = StringVar()


#------------------- Creacion Ventana Correlacion -----------------
LabelXC = Label(ventana, text="Columna X", bg=C3, font=('arial', 10), width=7, height=HB)
TxtXC = Entry(ventana, font=('arial', 20), width=33, textvariable=ArrayUno, bd=2, insertwidth=2)

LabelYC = Label(ventana, text="Columna Y", bg=C3, font=('arial', 10), width=7, height=HB)
TxtYC = Entry(ventana, font=('arial', 20), width=33, textvariable=ArrayDos, bd=2, insertwidth=2)

BtnEC = Button(ventana, text="Enviar", bg=C1, width=AB, height=HB, command=CorrelacionManual)
BtnCC = Button(ventana, text="Limpiar", bg=C2, width=AB, height=HB, command=clearCor)

LabelXC.place(x=20, y=125)
TxtXC.place(x=20, y=160)
LabelYC.place(x=20, y=205)
TxtYC.place(x=20, y=240)

BtnEC.place(x=275, y=400)
BtnCC.place(x=10, y=400)


#------------------- Creacion Ventana Excel -----------------
LabelNE = Label(ventana, text="Nombre del Archivo", bg=C3, font=('arial', 10), width=13, height=HB)
TxtNE = Entry(ventana, font=('arial', 20), width=33, textvariable=NombreExcel, bd=2, insertwidth=2)
LabelNH = Label(ventana, text="Nombre de la Hoja", bg=C3, font=('arial', 10), width=13, height=HB)
TxtNH = Entry(ventana, font=('arial', 20), width=33, textvariable=NombreHoja, bd=2, insertwidth=2)
LabelXE = Label(ventana, text="Nombre de la Primera Columna", bg=C3, font=('arial', 10), width=23, height=HB)
TXTXE = Entry(ventana, font=('arial', 20), width=33, textvariable=NombreColUno, bd=2, insertwidth=2)
LabelEY = Label(ventana, text="Nombre de la Segunda Columna", bg=C3, font=('arial', 10), width=23, height=HB)
TXTYE = Entry(ventana, font=('arial', 20), width=33, textvariable=NombreColDos, bd=2, insertwidth=2)

BtnCE = Button(ventana, text="Enviar", bg=C1, width=AB, height=HB, command=CorrelacionExcel)
BtnEE = Button(ventana, text="Limpiar", bg=C2, width=AB, height=HB, command=clearExcel)

#------------------- Ocultar Widgets ---------------------------

def CambiarM(event=None):
    LabelNE.place_forget()
    TxtNE.place_forget()
    LabelNH.place_forget()
    TxtNH.place_forget()
    LabelXE.place_forget()
    TXTXE.place_forget()
    LabelEY.place_forget()
    TXTYE.place_forget()
    BtnCE.place_forget()
    BtnEE.place_forget()

    LabelXC.place(x=20, y=125)
    TxtXC.place(x=20, y=160)
    LabelYC.place(x=20, y=205)
    TxtYC.place(x=20, y=240)

    BtnEC.place(x=275, y=400)
    BtnCC.place(x=10, y=400)

def CambiarE(event=None):
    LabelXC.place_forget()
    TxtXC.place_forget()
    LabelYC.place_forget()
    TxtYC.place_forget()
    BtnEC.place_forget()
    BtnCC.place_forget()

    LabelNE.place(x=20, y=50)
    TxtNE.place(x=20, y=85)
    LabelNH.place(x=20, y=130)
    TxtNH.place(x=20, y=165)
    LabelXE.place(x=20, y=205)
    TXTXE.place(x=20, y=240)
    LabelEY.place(x=20, y=280)
    TXTYE.place(x=20, y=315)

    BtnEE.place(x=10, y=400)
    BtnCE.place(x=275, y=400)

#------------------------ Menu -------------------------
ButtonC = Button(ventana, text="Correlación Manual", bg=M, width=AB, height=HB, command=CambiarM).place(x=10, y=9)
ButtonE = Button(ventana, text="Correlación con Excel", bg=M, width=AB, height=HB, command=CambiarE).place(x=275, y=9)

ventana.mainloop()
