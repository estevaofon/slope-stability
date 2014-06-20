#!/usr/bin/python
from math import *
from xlrd import open_workbook
import sys
from Tkinter import *
import tkMessageBox
from tkFileDialog import askopenfilename
import ntpath

class CalcMethods:
    def loadSheet(self, sheet, opcao=2):
        book = open_workbook(sheet)
        sheet0 = book.sheet_by_index(0)

        self.a = sheet0.col_values(1,1)
        self.b = sheet0.col_values(2,1)
        self.c = sheet0.col_values(3,1)
        self.phi = sheet0.col_values(4,1)
        self.w = sheet0.col_values(5,1)
        self.u = sheet0.col_values(6,1)
        self.Q = sheet0.col_values(7,1)
        if opcao == 3:
            self.fo = sheet0.col_values(7,1)
            self.fo = self.fo[0]

        for i in range(len(self.a)):
            self.a[i]= radians(self.a[i])

        for i in range(len(self.phi)):
            self.phi[i]= radians(self.phi[i])

        self.n = len(self.a) 

    def bishop(self,F):
        try:
            soma1 = 0
            soma2 = 0
            
            if F == 0 or F < 0.5:
                F = 0.5

            for i in range (self.n):
                            
                Y=(self.w[i]+self.Q[i]-self.b[i]*self.u[i])*tan(self.phi[i])+self.b[i]*self.c[i]
                X = cos(self.a[i])*(1+(tan(self.a[i])*tan(self.phi[i]))/F)
                W =(Y)/(X)
                soma1 += W

                Z = (self.w[i]+self.Q[i])*sin(self.a[i])
                soma2 += Z
                  
            somarf = (soma1/soma2)
            delta = somarf - F
            discrepancia = (delta/somarf)*100
            self.FS = '%.3f' % somarf
            print 'F estimado = %.3f' % F
            print 'F = %.3f' % somarf
            print 'Discrepancia = %.3f ' %discrepancia,
            print'%\n'
            if (delta) >= 0.005:
                F+=0.005
                self.bishop(F)
            if (delta) <= -0.005:
                F-=0.005
                self.bishop(F)
        except:
            raise
    def fellenius(self):
        try:
            soma1 = 0
            soma2 = 0
            
            for i in range (self.n):
                Y = ((self.w[i]+self.Q[i])*cos(self.a[i])-(self.u[i]*self.b[i])/cos(self.a[i]))*tan(self.phi[i])
                X = (self.c[i]*self.b[i])/cos(self.a[i])
                W = Y+X
                soma1 += W

                Z =(self.w[i]+self.Q[i])*sin(self.a[i])
                soma2 += Z
                  
            somarf = soma1/soma2
            self.FS = '%.3f' % somarf
        except:
            raise
    def jambu(self,F):
        try:
            soma1 = 0
            soma2 = 0
            
            if F == 0 or F < 0.5:
                F = 0.5

            for i in range (self.n):
                            
                Y=(self.c[i]+((self.w[i]/self.b[i])-self.u[i])*tan(self.phi[i]))*self.b[i]
                X = (cos(self.a[i])**2)*(1+(tan(self.a[i])*tan(self.phi[i]))/F)
                W =(Y)/(X)
                soma1 += W

                Z = self.w[i]*tan(self.a[i])
                soma2 += Z
                
            somarf = self.fo*(soma1/soma2)
            delta = somarf - F
            discrepancia = (delta/somarf)*100
            self.FS = '%.3f' % somarf
            if (delta) >= 0.005:
                F+=0.005
                self.jambu(F)
                
            if (delta) <= -0.005:
                F-=0.005
                self.jambu(F)
        except:
            raise
class Gui:
    def __init__(self):
        self.opcao = 2
        self.filepath = ''
        self.calc = CalcMethods()
    def run(self):
        app = Tk()
        app.title("Estabilidade de Taludes")
        app.geometry('300x160')

        labelText = StringVar()
        labelText.set("Metodo de Calculo")
        label1 = Label(app, textvariable=labelText, height=2)
        label1.pack()

        self.var = IntVar()
        R1 = Radiobutton(app, text="Fellenius", variable=self.var, value=1,command=self.sel)
        R1.pack()
        R2 = Radiobutton(app, text="Bishop", variable=self.var, value=2,command=self.sel)
        R2.pack()
        R2.select()
        R3 = Radiobutton(app, text="Jambu", variable=self.var, value=3,command=self.sel)
        R3.pack()

        button1 = Button(app, text="Calcular", width=10, command = self.changeLabel)
        button1.pack()
        
        menubar = Menu(app)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Selecionar Planilha", command=self.filechoose)

        filemenu.add_separator()

        filemenu.add_command(label="Exit", command=app.quit)
        menubar.add_cascade(label="Arquivo", menu=filemenu)
        
        aboutmenu = Menu(menubar, tearoff=0)

        aboutmenu.add_command(label="O App", command=self.newWindow)
        menubar.add_cascade(label="Sobre", menu=aboutmenu)

        self.labelText2 = StringVar()
        self.labelText2.set("")
        label2 = Label(app, textvariable=self.labelText2, height=2, fg='blue')
        label2.pack()

        app.config(menu=menubar)

        app.mainloop()
    
    def changeLabel(self):
        try:
            if self.opcao == 0:
                self.notSelected()
            elif self.opcao ==1:
                self.calc.loadSheet(self.filepath)
                self.calc.fellenius()
                self.pronto()
            elif self.opcao ==2:
                self.calc.loadSheet(self.filepath)
                self.calc.bishop(1)
                self.pronto()
            elif self.opcao ==3:
                self.calc.loadSheet(self.filepath,self.opcao)
                self.calc.jambu(1)
                self.pronto()
        except IOError:
            self.missingFile()
        except:
            self.wrongSheet()
        return

    def sel(self):
        self.opcao = self.var.get()
    def pronto(self):
        tkMessageBox.showinfo("Status", "F.S = "+self.calc.FS)
    def missingFile(self):
        tkMessageBox.showinfo("Status", "Arquivo nao selecionado!")
    def unknowError(self):
        tkMessageBox.showinfo("Status", "Erro desconhecido!")
    def wrongSheet(self):
        erroMessage = sys.exc_info()[:2]
        tkMessageBox.showinfo("Status", "Planilha e metodo incompativeis!\n"+"Erro: \n"+str(erroMessage[0])+"\n"+str(erroMessage[1]))

    def filechoose(self):
        self.filepath = askopenfilename()
        if self.filepath != '':
            self.labelText2.set(ntpath.basename(self.filepath)+" selecionado")
    def newWindow(self):
        top = Toplevel()
        top.title("O App")
        top.geometry('380x80')
        labelText = StringVar()
        labelText.set("Aplicativo desenvolvido em Python por Estevao Fonseca")
        label1 = Label(top, textvariable=labelText, height=2)
        label1.pack(side='top', padx=10, pady=15)
        top.mainloop()
gui = Gui()
gui.run()
    
