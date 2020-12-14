import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter import font
from tkinter.messagebox import showinfo

import PyPDF2

number_copy = ''

class SeaofBTCapp(tk.Tk):

    def __init__(self, *args, **kwargs):
     
        tk.Tk.__init__(self, *args, **kwargs)

        #INITIALISATION GENERAL
        #configuration de la fenetre
        self.geometry("500x500")
        self.bind('<Key>', self.get_barcode)
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
                print(number_copy_cast)
                self.frames[StartPage].tree.item(item, values = number_copy_cast)

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

        #configuration ligne/colonne StartPage
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=3)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)

        self.pdfWriter = PyPDF2.PdfFileWriter()

        ttk.Button(self, text="Ajouter PDF", command=lambda: self.click(controller)).grid(row= 0, column = 0)
        ttk.Button(self, text="Supprimer PDF", command=lambda: self.delete()).grid(row= 0, column = 1)

         #TREEVIEW
        self.tree = ttk.Treeview(self, selectmode = "extended", height = 0)
        self.tree["columns"]=("one")
        
        self.tree.column("#0")
        self.tree.column("one")
        
        self.tree.heading("#0", text="fichier")
        self.tree.heading("one", text="copie")

        self.tree.grid(row=1, column=0,columnspan = 2, sticky='news')

        self.nb_copy = ttk.Entry(self)
        self.nb_copy.grid(column =0, row = 2)

        ttk.Button(self, text="ok", command=lambda: self.change_nb_copy()).grid(row= 2, column = 1)

        ttk.Button(self, text="Annuler", command=lambda: self.reset()).grid(row= 3, column = 0)
        ttk.Button(self, text="Valider", command=lambda: self.validation()).grid(row= 3, column = 1)

        

    def click(self, controller): 
        pdf_to_merge = filedialog.askopenfilename()
        self.tree.insert("" , "end",    text=pdf_to_merge, values=(1))

    def change_nb_copy(self):
        item = self.tree.selection()[0]
        nb_copy_ = int(self.nb_copy.get())
        if(nb_copy_ > 0):
            print("ok")
            self.tree.item(item, values = (nb_copy_))

    def reset(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def delete(self):
        item = self.tree.selection()[0]
        self.tree.delete(item)
        self.pdfWriter = PyPDF2.PdfFileWriter()
                      
    def validation(self):
        userfilename = filedialog.asksaveasfilename()

        for Parent in self.tree.get_children():
            
            pdfFileObj = open(self.tree.item(Parent)['text'], 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

            nb_copy_pdf = self.tree.item(Parent)['values'][0]

            print(nb_copy_pdf)

            for i in range(nb_copy_pdf):
                for pageNum in range(pdfReader.numPages):
                    pageObj = pdfReader.getPage(pageNum)
                    self.pdfWriter.addPage(pageObj)            

            pdfOutput = open(userfilename + '.pdf', 'wb')            
            self.pdfWriter.write(pdfOutput)            
            pdfOutput.close()


class PageOne(tk.Frame):
    #Constructeur de la page 1
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


app = SeaofBTCapp()
app.mainloop()