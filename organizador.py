import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror
import pandas as pd
import os, shutil


class Window:
    def __init__(self, master):
        self.master = master
        master.title("Organizador de archivos")
        master.geometry('320x360')
        master.eval('tk::PlaceWindow . center')

        # Label e input file de directorio
        self.frame = tk.Frame(master)
        self.frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.labelDirectory = tk.Label(
            self.frame, text="Directorio de origen")
        self.labelDirectory.place(x=0, y=0, anchor='nw')
        self.labelDirectory.pack(fill=tk.X)

        self.entryDirectory = tk.Entry(self.frame)
        self.entryDirectory.pack(fill=tk.X)

        self.buttonDirectory = tk.Button(
            self.frame, text="Seleccionar", command=self.getDir)
        self.buttonDirectory.pack(fill=tk.X)

        # Label e input file de directorio
        self.frame = tk.Frame(master)
        self.frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.labelOutputDirectory = tk.Label(
            self.frame, text="Directorio de salida")
        self.labelOutputDirectory.place(x=0, y=0, anchor='nw')
        self.labelOutputDirectory.pack(fill=tk.X)

        self.entryOutputDirectory = tk.Entry(self.frame)
        self.entryOutputDirectory.pack(fill=tk.X)

        self.buttonOutputDirectory = tk.Button(
            self.frame, text="Seleccionar", command=self.getOutputDir)
        self.buttonOutputDirectory.pack(fill=tk.X)

        # Label e input file de archivo
        self.frameFile = tk.Frame(master)
        self.frameFile.pack(fill=tk.X, padx=10, pady=10)

        self.labelFile = tk.Label(
            self.frameFile, text="Archivo excel")
        self.labelFile.place(x=0, y=0, anchor='nw')
        self.labelFile.pack(fill=tk.X)

        self.entryFile = tk.Entry(self.frameFile)
        self.entryFile.pack(fill=tk.X)

        self.buttonFile = tk.Button(
            self.frameFile, text="Seleccionar", command=self.getFile)
        self.buttonFile.pack(fill=tk.X)

        # Label e input file de archivo
        self.frameFile = tk.Frame(master)
        self.frameFile.pack(fill=tk.X, padx=10, pady=10)

        self.buttonInit = tk.Button(
            self.frameFile, text="Iniciar Organización", command=self.initOrganization)
        self.buttonInit.pack(fill=tk.X)

        self.progressBar = ttk.Progressbar(length=100)

    def getFile(self):
        filename = filedialog.askopenfilename(
            initialdir="/", title="Select a File", filetypes=(("Text files", "*.xlsx*"), ("all files", "*.*")))
        if(filename != ""):
            self.entryFile.insert(0, filename)

    def getDir(self):
        directory = filedialog.askdirectory()
        if directory != "":
            self.entryDirectory.insert(0, directory)

    def getOutputDir(self):
        directory = filedialog.askdirectory()
        if directory != "":
            self.entryOutputDirectory.insert(0, directory)

    def initOrganization(self):
        originDirectory = ""
        filePath = ""
        outputDirectory = ""
        self.stateWindow(tk.DISABLED)
        self.showProgressBar(True)

        if(len(self.entryDirectory.get()) <= 0):
            showerror("Error", "No ha seleccionado directorio")
            self.stateWindow(tk.NORMAL)
            self.showProgressBar(False)
            return
        if(len(self.entryOutputDirectory.get()) <= 0):
            showerror("Error", "No ha seleccionado directorio de salida")
            self.stateWindow(tk.NORMAL)
            self.showProgressBar(False)
            return
        if(len(self.entryFile.get()) <= 0):
            showerror("Error", "No ha seleccionado archivo")
            self.showProgressBar(False)
            self.stateWindow(tk.NORMAL)
            return

        filePath = self.entryFile.get()
        originDirectory = self.entryDirectory.get()
        outputDirectory = self.entryOutputDirectory.get()

        excelContent = pd.read_excel(filePath)

        for i in excelContent.index:

            localidad = excelContent['Localidad'][i]

            charactersToRemove = '<>:?*/'

            for character in charactersToRemove:
                localidad = localidad.replace(character, "")

            outputFileDirectory = outputDirectory + '/' + localidad
            fileName = excelContent['RFC'][i]+".pdf"
            originFilePath = originDirectory+'/'+fileName
            outputFilePath = outputFileDirectory+'/'+fileName

            if(os.path.exists(originFilePath) == False):
                continue

            if(os.path.exists(outputFileDirectory) == False):
                os.mkdir(outputFileDirectory)

            shutil.copy(originFilePath, outputFilePath)

        self.showProgressBar(False)
        self.stateWindow(tk.NORMAL)

    def stateWindow(self, state):
        self.entryDirectory['state'] = state
        self.entryFile['state'] = state
        self.entryOutputDirectory['state'] = state
        self.buttonDirectory['state'] = state
        self.buttonFile['state'] = state
        self.buttonOutputDirectory['state'] = state
        self.buttonInit['state'] = state

    def showProgressBar(self, isVisible):
        if(isVisible):
            self.progressBar.pack(fill=tk.X, padx=10, pady=10)
            self.progressBar.start()
        else:
            self.progressBar.stop()
            self.progressBar.pack_forget()


root = tk.Tk()
window = Window(root)
root.mainloop()
