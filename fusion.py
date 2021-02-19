import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter import font
from tkinter.messagebox import showinfo
from tkinter.messagebox import askquestion
from tkinter.messagebox import showerror
from tkinter import simpledialog
import PyPDF2
import os
import configparser
from docx2pdf import convert
from pathlib import Path


number_copy = ''
class SeaofBTCapp(tk.Tk):

    def __init__(self, *args, **kwargs):
     
        tk.Tk.__init__(self, *args, **kwargs)

        #INITIALISATION GENERAL
        #configuration de la fenetre
        self.geometry("700x700")
        #self.resizable(width=0, height=0)
        self.title("PDF + Word Fusion")
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
        self.columnconfigure( 0, weight=1)
        self.columnconfigure( 1, weight=1)
        self.columnconfigure( 2, weight=1)
        self.columnconfigure( 3, weight=1)

        self.rowconfigure( 0, weight = 1)
        self.rowconfigure( 1, weight = 1)
        self.rowconfigure( 2, weight = 6)
        self.rowconfigure( 3, weight = 1)
        self.rowconfigure( 4, weight = 1)
        self.rowconfigure( 5, weight = 1)

        ttk.Label(self, text = "Fusion", font = ("Comic Sans MS", 30)).grid(row= 0, column = 0,columnspan = 4)
        ttk.Button(self, text="Ajouter fichier", command=lambda: self.click(controller)).grid(row= 1, column = 0,columnspan = 4,sticky='')
        
        ################
        # Labelframe 0 #
        ################

        labelframe0 = LabelFrame(self, text="Liste des fichiers")
        labelframe0.grid(row=2,column = 0,columnspan = 3, sticky='news')

        labelframe0.columnconfigure(1, weight=1)
        labelframe0.rowconfigure(1, weight=1)
        
        self.tree = ttk.Treeview(labelframe0, selectmode = "extended", height = 0)
        self.tree["columns"]=("one","two","three", "four","five")
        
        self.tree.column("#0")
        self.tree.column("one",     width = 25, minwidth=5)
        self.tree.column("two",     width = 25, minwidth=5)
        self.tree.column("three",   width = 25, minwidth=5)
        self.tree.column("four",    width = 25, minwidth=5)
        self.tree.column("five",    width = 25, minwidth=5)
        
        self.tree.heading("#0", text="fichier")
        self.tree.heading("one", text="copie")
        self.tree.heading("two", text="chemin")
        self.tree.heading("three", text="Nb page")
        self.tree.heading("four", text="debut")
        self.tree.heading("five", text="fin")

        self.tree["displaycolumns"]=("one", "three", "four","five")

        self.tree.grid(row = 1, column = 1, sticky='news')

        ################
        # Labelframe 1 #
        ################

        labelframe1 = LabelFrame(self, text="Edition")
        labelframe1.grid(row=2,column = 3, sticky='news')
        
        for i in  range(3):
            labelframe1.columnconfigure(i, weight=1)

        for i in  range(4):
            labelframe1.rowconfigure(i, weight=1)

        ttk.Label(labelframe1, text = "Copie").grid(column=0,row=0)
        self.edition_nombre_de_copie = ttk.Entry(labelframe1)
        self.edition_nombre_de_copie.grid(column=1,row=0, sticky = 'ew')

        ttk.Label(labelframe1, text = "Debut").grid(column=0,row=1)
        self.edition_nombre_debut = ttk.Entry(labelframe1)
        self.edition_nombre_debut.grid(column=1,row=1, sticky = 'ew')

        ttk.Label(labelframe1, text = "Fin").grid(column=0,row=2)
        self.edition_nombre_fin = ttk.Entry(labelframe1)
        self.edition_nombre_fin.grid(column=1,row=2, sticky = 'ew')

        ttk.Button(labelframe1, text="<<--", command=lambda: self.editer()).grid(row = 3, column = 1, sticky = 'ew')
        
        ################
        # Labelframe 2 #
        ################

        labelframe2 = LabelFrame(self, text="Preset")
        labelframe2.grid(row=3,column = 0, columnspan = 4, sticky='news')

        for i in  range(6):
            labelframe2.columnconfigure(i, weight=1)
        
        labelframe2.rowconfigure(0, weight=1)
        
        ttk.Button(labelframe2, text="CE2", command=lambda: self.auto(nombre_de_ce2)).grid(column=0, row = 0, sticky = 'ew')
        ttk.Button(labelframe2, text="CM1", command=lambda: self.auto(nombre_de_cm1)).grid(column=1, row = 0, sticky = 'ew')
        ttk.Button(labelframe2, text="CM2", command=lambda: self.auto(nombre_de_cm2)).grid(column=2, row = 0, sticky = 'ew')
        ttk.Button(labelframe2, text="CE2F", command=lambda: self.auto(nombre_de_ce2_f)).grid(column=3, row = 0, sticky = 'ew')
        ttk.Button(labelframe2, text="CM1+CM2", command=lambda: self.auto(nombre_de_cm2+nombre_de_cm1)).grid(column=4, row = 0, sticky = 'ew')
        ttk.Button(labelframe2, text="Tous", command=lambda: self.auto(nombre_de_cm2+nombre_de_cm1+nombre_de_ce2)).grid(column=5, row = 0, sticky = 'ew')

        ################
        # Labelframe 3 #
        ################

        labelframe3 = LabelFrame(self, text="Parametre d'impression")
        labelframe3.grid(row=4,column = 0, columnspan = 4, sticky='news')

        labelframe3.columnconfigure(0, weight=1)
        labelframe3.columnconfigure(1, weight=1)

        labelframe3.rowconfigure(0, weight=1)

        self.chkValue = tk.BooleanVar() 
        self.chkValue.set(False)
        ttk.Checkbutton(labelframe3, text='Insertion page blanche', var=self.chkValue).grid(column=0, row=0)

        self.chkValue2 = tk.BooleanVar() 
        self.chkValue2.set(False)
        ttk.Checkbutton(labelframe3, text='Insertion séparateur', var=self.chkValue2).grid(column=1, row=0)

        ################
        # Labelframe 4 #
        ################

        labelframe4 = LabelFrame(self, text="Finalisation")
        labelframe4.grid(row=5,column = 0, columnspan = 4, sticky='news')

        for i in  range(2):
            labelframe4.columnconfigure(i, weight=1)
        labelframe4.rowconfigure(1, weight=1)

        ttk.Button(labelframe4, text="Reset", command=lambda: self.reset()).grid(column = 0, row = 1, sticky = '')
        ttk.Button(labelframe4, text="Fusion", command=lambda: self.validation()).grid(column = 1, row = 1, sticky = '')

    def editer(self):
        #on recupere la ligne selectionner
        item = self.tree.selection()[0]
        
        #Page du debut
        if self.edition_nombre_de_copie.get() == '':
            nb_copie = int(self.tree.item(item)['values'][0])
        else:
            nb_copie = int(self.edition_nombre_de_copie.get())

        #Page du debut
        if self.edition_nombre_debut.get() == '':
            nb_debut = int(self.tree.item(item)['values'][3])
        else:
            nb_debut = int(self.edition_nombre_debut.get())
        
        #Page de fin
        if self.edition_nombre_fin.get() == '':
            nb_fin = int(self.tree.item(item)['values'][4])
        else:
            nb_fin = int(self.edition_nombre_fin.get())

        nb_page = int(self.tree.item(item)['values'][2])

        if  (nb_debut <= nb_fin) and (nb_debut > 0) and (nb_fin > 0) and (nb_fin <= nb_page) and (nb_debut <= nb_page):
            self.tree.item(item, values = (nb_copie, self.tree.item(item)['values'][1],nb_page ,nb_debut, nb_fin))
            
            #raz des entry
            self.edition_nombre_de_copie.delete(0,END)
            self.edition_nombre_debut.delete(0,END)
            self.edition_nombre_fin.delete(0,END)
        else:
            showerror(title="Erreur de saisie", message= "saisie invalide.")


        self.edition_nombre_de_copie.delete(0,END)
        self.edition_nombre_debut.delete(0,END)
        self.edition_nombre_fin.delete(0,END)

    def auto(self,data):
        item = self.tree.selection()[0]
        if(data > 0):
            self.tree.item(item, values = (data, self.tree.item(item)['values'][1], self.tree.item(item)['values'][2], self.tree.item(item)['values'][3], self.tree.item(item)['values'][4]))

    def click(self, controller):
        pdf_to_merge = filedialog.askopenfilename(filetypes=(("WORD/PDF files","*.docx"),("WORD/PDF files","*.pdf")), multiple=True)
        #pdf_to_merge = filedialog.askopenfilename(multiple=True)
        
        

        print(pdf_to_merge)
        for path in pdf_to_merge:
            p = Path(path)        
        
            if path != '':
                if p.suffix == '.pdf':
                    self.tree.insert("" , "end",    text=p.name, values=(1,path,self.obtenir_nombre_de_page_pdf(path), 1, self.obtenir_nombre_de_page_pdf(path) ))
                    showinfo("Import", "Importation PDF terminée."  )

                elif p.suffix == '.docx':
                    temp_path = 'conversion\\' + p.stem +'.pdf'
                    p_copy = Path(temp_path)
                    convert(path, temp_path)
                    self.tree.insert("" , "end",    text=p_copy.name, values=(1,temp_path,self.obtenir_nombre_de_page_pdf(temp_path), 1,self.obtenir_nombre_de_page_pdf(temp_path)))	
                    showinfo("Conversion Word en PDF", "Conversion de "+ p.name + " en " + p_copy.name + " terminée."  )

                else:
                    showinfo("Problème fichier", "Une erreur est survenue avec le fichier."  )
        

    def obtenir_nombre_de_page_pdf(self, adresse):
        #ouverture du fichier
        pdfFileObj = open(adresse, 'rb')

        #generation de l'objet pdf
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

        #sauvegarde du nombre de page
        num_page = int(pdfReader.numPages)

        #fermeture du fichier
        pdfFileObj.close()

        return num_page


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
                pdfFileSeparator = open("ressource/intercalaire.pdf", 'rb')

                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pdfReaderBlank = PyPDF2.PdfFileReader(pdfFileBlank)
                pdfReaderSeparator = PyPDF2.PdfFileReader(pdfFileSeparator)
                
                nb_copy_pdf = self.tree.item(Parent)['values'][0]

                nb_debut = self.tree.item(Parent)['values'][3]
                nb_fin = self.tree.item(Parent)['values'][4]

                for i in range(nb_copy_pdf):
                    for pageNum in range(nb_debut-1, nb_fin):
                        pageObj = pdfReader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                        if self.chkValue.get() == True:
                            #Ajout de la page blanche
                            pageObj = pdfReaderBlank.getPage(0)
                            pdfWriter.addPage(pageObj)
                    
                    if ((nb_fin - (nb_debut-1)) %2 == 1) and (self.chkValue.get() != True):
                        #si nombre de page impair on rajoute une page blanche
                        pageObj = pdfReaderBlank.getPage(0)
                        pdfWriter.addPage(pageObj)

                        if self.chkValue.get() == True:
                            #Ajout de la page blanche si case coché et impair
                            pageObj = pdfReaderBlank.getPage(0)
                            pdfWriter.addPage(pageObj)
                
                if self.chkValue2.get() == True:
                    #Ajout de l'intercalaire'
                    print("intercalaire ok")
                    pageObj = pdfReaderSeparator.getPage(0)
                    pdfWriter.addPage(pageObj)
                    pageObj = pdfReaderSeparator.getPage(1)
                    pdfWriter.addPage(pageObj)
                    

                pdfOutput = open(userfilename + '.pdf', 'wb')            
                pdfWriter.write(pdfOutput)            
                    
                #fermeture des fichiers
                pdfOutput.close()
                pdfFileBlank.close()
                pdfFileObj.close()

                #reset du treeview                    
                self.tree.delete(Parent)
        showinfo("Fusion", "Fusion terminée."  )
            

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