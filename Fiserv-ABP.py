#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from LIB.Fiserv import Procesar_Fiserv

def Colaboraciones():
    # Abrir el navegador en "https://cafecito.app/abustos"
    import os
    os.system("start https://cafecito.app/abustos")



class GUI:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(background="#2e2e2e", height=310, width=260)
        Toplevel_1.iconbitmap("LIB/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.overrideredirect("False")
        Toplevel_1.resizable(False, False)
        Toplevel_1.title("Procesar Fiserv")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="LIB/ABP blanco en sin fondo.png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            state="normal",
            takefocus=False,
            text='\nProcesamiento de PDFs de Fiserv.\n',
            wraplength=325)
        Label_1.pack(expand="true", side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand="true", side="top")
        self.Seleccionar_carpeta = ttk.Button(Toplevel_1)
        self.Seleccionar_carpeta.configure(text='Seleccionar Carpeta' , command=Procesar_Fiserv)
        self.Seleccionar_carpeta.pack(expand="true", pady=4, side="top")

        self.Colaboraciones = ttk.Button(Toplevel_1)
        self.Colaboraciones.configure(text='Colaboraciones' , command=Colaboraciones)
        self.Colaboraciones.pack(expand="true", pady=4, side="top")
    

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = GUI()
    app.run()
