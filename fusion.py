import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter import font
from tkinter.messagebox import showinfo
import PyPDF2
import os

import configparser

###############################################################################################
#                              ZONE DE MODIFICATION                                           #
###############################################################################################


###############################################################################################



number_copy = ''
class SeaofBTCapp(tk.Tk):

    def __init__(self, *args, **kwargs):
     
        tk.Tk.__init__(self, *args, **kwargs)

        #INITIALISATION GENERAL
        #configuration de la fenetre
        self.geometry("500x500")
        self.bind('<Key>', self.get_barcode)
        self.title("PDF Fusion")
        #self.attributes("-fullscreen", True)       
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand = True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, PageOne):

            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)
        
    def get_barcode(self, event):

        global number_copy

        if event.char in '0123456789':
            number_copy += event.char

        elif event.keysym == 'Return':
            number_copy_cast = int(number_copy)
            item = self.frames[StartPage].tree.selection()[0]
            if(number_copy_cast > 0):
                self.frames[StartPage].tree.item(item, values = (number_copy_cast, self.frames[StartPage].tree.item(item)['values'][1]))
            else:
                self.frames[StartPage].tree.delete(item)
            number_copy = ''

        elif event.keysym == 'BackSpace':
            item = self.frames[StartPage].tree.selection()[0]
            self.frames[StartPage].tree.delete(item)
            number_copy = ''
        
    def show_frame(self, cont):
        global id_page

        id_page=cont
        frame = self.frames[cont]
        frame.tkraise()
        
class StartPage(tk.Frame):
    #constructeur de la page d'acceuil
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)

        global nombre_de_ce2
        global nombre_de_cm1
        global nombre_de_cm2

        #configuration ligne/colonne StartPage
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)
        self.rowconfigure(0, weight=1)  #reserve pour le titre
        self.rowconfigure(1, weight=2)
        self.rowconfigure(2, weight=4)
        self.rowconfigure(3, weight=2)
        self.rowconfigure(4, weight=2)
        self.rowconfigure(5, weight=2)

        ttk.Label(self, text = "Fusion", font = ("Comic Sans MS", 30)).grid(row= 0, column = 0,columnspan = 3)
        ttk.Button(self, text="Ajouter PDF", command=lambda: self.click(controller)).grid(row= 1, column = 0,columnspan = 3)

        #TREEVIEW
        self.tree = ttk.Treeview(self, selectmode = "extended", height = 0)
        self.tree["columns"]=("one","two")
        
        self.tree.column("#0")
        self.tree.column("one")
        self.tree.column("two")
        
        self.tree.heading("#0", text="fichier")
        self.tree.heading("one", text="copie")
        self.tree.heading("two", text="chemin")

        self.tree["displaycolumns"]=("one")

        self.tree.grid(row=2, column=0,columnspan = 3, sticky='news')

        ttk.Button(self, text="CE2", command=lambda: self.auto(nombre_de_ce2)).grid(row = 3, column = 0)
        ttk.Button(self, text="CM1", command=lambda: self.auto(nombre_de_cm1)).grid(row = 3, column = 1)
        ttk.Button(self, text="CM2", command=lambda: self.auto(nombre_de_cm2)).grid(row = 3, column = 2)

        ttk.Button(self, text="CE2F", command=lambda: self.auto(nombre_de_ce2_f)).grid(row = 4, column = 0)
        ttk.Button(self, text="CM1+CM2", command=lambda: self.auto(nombre_de_cm2+nombre_de_cm1)).grid(row = 4, column = 1)
        ttk.Button(self, text="Tous", command=lambda: self.auto(nombre_de_cm2+nombre_de_cm1+nombre_de_ce2)).grid(row = 4, column = 2)

        self.chkValue = tk.BooleanVar() 
        self.chkValue.set(False)
        ttk.Checkbutton(self, text='Insertion page blanche', var=self.chkValue).grid(column=2, row=5)


        ttk.Button(self, text="Reset", command=lambda: self.reset()).grid(row = 5, column = 0)
        ttk.Button(self, text="Fusion", command=lambda: self.validation()).grid(row = 5, column = 1)

    def auto(self,data):
        item = self.tree.selection()[0]
        if(data > 0):
            self.tree.item(item, values = (data, self.tree.item(item)['values'][1]))

    def click(self, controller):
        pdf_to_merge = filedialog.askopenfilename()

        name = os.path.basename(pdf_to_merge)

        if pdf_to_merge != '':
            self.tree.insert("" , "end",    text=name, values=(1,pdf_to_merge))

    def reset(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def delete(self):
        item = self.tree.selection()[0]
        self.tree.delete(item)
                      
    def validation(self):
        userfilename = filedialog.asksaveasfilename()
        pdfWriter = PyPDF2.PdfFileWriter()

        if userfilename != '':
            for Parent in self.tree.get_children():
                pdfFileObj = open(self.tree.item(Parent)['values'][1], 'rb')
                pdfFileBlank = open("ressource/blank.pdf", 'rb')

                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pdfReaderBlank = PyPDF2.PdfFileReader(pdfFileBlank)
                
                nb_copy_pdf = self.tree.item(Parent)['values'][0]

                

                for i in range(nb_copy_pdf):
                    for pageNum in range(pdfReader.numPages):
                        pageObj = pdfReader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                        if self.chkValue.get() == True:
                            #Ajout de la page blanche
                            pageObj = pdfReaderBlank.getPage(0)
                            pdfWriter.addPage(pageObj)
                    if pdfReader.numPages%2 == 1:
                        #si nombre de page impair on rajoute une page blanche
                        pageObj = pdfReaderBlank.getPage(0)
                        pdfWriter.addPage(pageObj)

                        if self.chkValue.get() == True:
                            #Ajout de la page blanche si case coché et impair
                            pageObj = pdfReaderBlank.getPage(0)
                            pdfWriter.addPage(pageObj)

                pdfOutput = open(userfilename + '.pdf', 'wb')            
                pdfWriter.write(pdfOutput)            
                    
                #fermeture des fichiers
                pdfOutput.close()
                pdfFileBlank.close()
                pdfFileObj.close()

                #reset du treeview                    
                self.tree.delete(Parent)

class PageOne(tk.Frame):
    #Constructeur de la page 1
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


#ouverture du fichier de configuration
config = configparser.ConfigParser()
config.read('config.ini')

#recuperation des données du fichier config.ini
nombre_de_ce2 = int(config.get('parametre', 'nombre_de_ce2'))
nombre_de_ce2_f = int(config.get('parametre', 'nombre_de_ce2_f'))
nombre_de_cm1 = int(config.get('parametre', 'nombre_de_cm1'))
nombre_de_cm2 = int(config.get('parametre', 'nombre_de_cm2'))

app = SeaofBTCapp()
app.mainloop()