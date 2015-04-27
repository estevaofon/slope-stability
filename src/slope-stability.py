#!/usr/bin/python
from math import *
from xlrd import open_workbook
import sys
from Tkinter import *
import tkMessageBox
from tkFileDialog import askopenfilename
import ntpath

class CalcMethods:
    def loadSheet(self, sheet, option=2):
        book = open_workbook(sheet)
        sheet0 = book.sheet_by_index(0)

        self.a = sheet0.col_values(1,1)
        self.b = sheet0.col_values(2,1)
        self.c = sheet0.col_values(3,1)
        self.phi = sheet0.col_values(4,1)
        self.w = sheet0.col_values(5,1)
        self.u = sheet0.col_values(6,1)
        self.Q = sheet0.col_values(7,1)
        if option == 3:
            self.fo = sheet0.col_values(7,1)
            self.fo = self.fo[0]

        for i in range(len(self.a)):
            self.a[i]= radians(self.a[i])

        for i in range(len(self.phi)):
            self.phi[i]= radians(self.phi[i])

        self.n = len(self.a)

    def bishop(self,F):
        try:
            sum1 = 0
            sum2 = 0

            if F == 0 or F < 0.5:
                F = 0.5

            for i in range (self.n):

                Y=(self.w[i]+self.Q[i]-self.b[i]*self.u[i])*tan(self.phi[i])+self.b[i]*self.c[i]
                X = cos(self.a[i])*(1+(tan(self.a[i])*tan(self.phi[i]))/F)
                W =(Y)/(X)
                sum1 += W

                Z = (self.w[i]+self.Q[i])*sin(self.a[i])
                sum2 += Z

            local_FS = (sum1/sum2)
            delta = local_FS - F
            discrepancy = (delta/local_FS)*100
            self.FS = '%.3f' % local_FS
            print('F estimated = %.3f' % F)
            print('F = %.3f' % local_FS)
            print('Discrepancy = {0:.3f}% \n'.format(discrepancy))
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
            sum1 = 0
            sum2 = 0

            for i in range (self.n):
                Y = ((self.w[i]+self.Q[i])*cos(self.a[i])-(self.u[i]*self.b[i])/cos(self.a[i]))*tan(self.phi[i])
                X = (self.c[i]*self.b[i])/cos(self.a[i])
                W = Y+X
                sum1 += W

                Z =(self.w[i]+self.Q[i])*sin(self.a[i])
                sum2 += Z

            local_FS = sum1/sum2
            self.FS = '%.3f' % local_FS
        except:
            raise
    def jambu(self,F):
        try:
            sum1 = 0
            sum2 = 0

            if F == 0 or F < 0.5:
                F = 0.5

            for i in range (self.n):

                Y=(self.c[i]+((self.w[i]/self.b[i])-self.u[i])*tan(self.phi[i]))*self.b[i]
                X = (cos(self.a[i])**2)*(1+(tan(self.a[i])*tan(self.phi[i]))/F)
                W =(Y)/(X)
                sum1 += W

                Z = self.w[i]*tan(self.a[i])
                sum2 += Z

            local_FS = self.fo*(sum1/sum2)
            delta = local_FS - F
            discrepancy = (delta/local_FS)*100
            self.FS = '%.3f' % local_FS
            print('F estimated = %.3f' % F)
            print('F = %.3f' % local_FS)
            print('Discrepancy = {0:.3f}% \n'.format(discrepancy))
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
        self.option = 2
        self.filepath = ''
        self.calc = CalcMethods()
    def run(self):
        app = Tk()
        app.title("Slope Stability")
        app.geometry('300x160')

        labelText = StringVar()
        labelText.set("Calculation Method")
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

        button1 = Button(app, text="Calculate", width=10, command = self.changeLabel)
        button1.pack()

        menubar = Menu(app)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Select Spreadsheet", command=self.filechoose)

        filemenu.add_separator()

        filemenu.add_command(label="Exit", command=app.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        aboutmenu = Menu(menubar, tearoff=0)

        aboutmenu.add_command(label="The App", command=self.newWindow)
        menubar.add_cascade(label="About", menu=aboutmenu)

        self.labelText2 = StringVar()
        self.labelText2.set("")
        label2 = Label(app, textvariable=self.labelText2, height=2, fg='blue')
        label2.pack()

        app.config(menu=menubar)

        app.mainloop()

    def changeLabel(self):
        try:
            if self.option == 0:
                self.notSelected()
            elif self.option ==1:
                self.calc.loadSheet(self.filepath)
                self.calc.fellenius()
                self.done()
            elif self.option ==2:
                self.calc.loadSheet(self.filepath)
                self.calc.bishop(1)
                self.done()
            elif self.option ==3:
                self.calc.loadSheet(self.filepath,self.option)
                self.calc.jambu(1)
                self.done()
        except IOError:
            self.missingFile()
        except:
            self.wrongSheet()
        return

    def sel(self):
        self.option = self.var.get()
    def done(self):
        tkMessageBox.showinfo("Status", "F.S = "+self.calc.FS)
    def missingFile(self):
        tkMessageBox.showinfo("Status", "No file selected!")
    def unknowError(self):
        tkMessageBox.showinfo("Status", "Unknow Error!")
    def wrongSheet(self):
        erroMessage = sys.exc_info()[:2]
        tkMessageBox.showinfo("Status", "Method and Spreadsheet doesn't match!\n"+"Error: \n"+str(erroMessage[0])+"\n"+str(erroMessage[1]))

    def filechoose(self):
        self.filepath = askopenfilename()
        if self.filepath != '':
            self.labelText2.set(ntpath.basename(self.filepath)+" selected")
    def newWindow(self):
        top = Toplevel()
        top.title("The App")
        top.geometry('380x80')
        labelText = StringVar()
        labelText.set("App developed in Python by Estevao Fonseca")
        label1 = Label(top, textvariable=labelText, height=2)
        label1.pack(side='top', padx=10, pady=15)
        top.mainloop()
gui = Gui()
gui.run()

