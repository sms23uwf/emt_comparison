# -*- coding: utf-8 -*-
"""
    
    main_menu.py
   --------------
   
   This module is the main menu for an application.

"""

__author__ = "Steven M. Satterfield"


from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename

baseFilePath = ""
comparisonFilePath = ""

def NewFile():
    print("New File!")
    


def OpenFile():
    name = askopenfilename()
    print(name)



def About():
    print("This is a simple example of a menu.")
  
    
  
def CallCompareEMTFiles():
    print("not yet implemented")
    print("baseFileName: {}".format(baseFilePath))
    print("comparisonFileName: {}".format(comparisonFilePath))
    
def GetBaseFilePath():
    baseFilePath = askopenfilename()
    


def GetComparisonFilePath():
    comparisonFilePath = askopenfilename()
    
    
    
def GetDiffFiles():
    entryform = Tk()
    entryform.title("Compare two EMT files for differences")
    fBaseLabel = ttk.Label(entryform, text="Select Base File:")
    fBaseEntry = ttk.Entry(entryform)
    bBaseEntry = ttk.Button(entryform, text="...", command=GetBaseFilePath)
    fBaseEntry.text = baseFilePath
    
    fCompLabel = ttk.Label(entryform, text="Select Comparison File:")
    fCompEntry = ttk.Entry(entryform)
    bCompEntry = ttk.Button(entryform, text="...", command=GetComparisonFilePath)
    fCompEntry.text = comparisonFilePath
    
    btnGo = ttk.Button(entryform, text="Go", command=CallCompareEMTFiles)
    btnCancel = ttk.Button(entryform, text="Cancel", command=entryform.quit)
    
    fBaseLabel.grid(row=0, column=1, padx=15, pady=15)
    fBaseEntry.grid(row=0, column=2, padx=15, pady=15)
    bBaseEntry.grid(row=0, column=3, padx=0, pady=15)
    fCompLabel.grid(row=1, column=1, padx=15, pady=15)
    fCompEntry.grid(row=1, column=2, padx=15, pady=15)
    bCompEntry.grid(row=1, column=3, padx=0, pady=15)
    
    btnGo.grid(row=5, column=1, padx=15, pady=15)
    btnCancel.grid(row=5, column=2, padx=15, pady=15)
    
   
root = Tk()

menu = Menu(root)
root.config(menu=menu)
filemenu = Menu(menu)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="New", command=NewFile)
filemenu.add_command(label="Open...", command=OpenFile)
filemenu.add_command(label="Compare EMT Files", command=GetDiffFiles)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
    
helpmenu = Menu(menu)
menu.add_cascade(label="Help", menu=helpmenu)
helpmenu.add_command(label="About...", command=About)

mainloop()
