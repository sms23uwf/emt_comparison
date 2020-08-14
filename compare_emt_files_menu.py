# -*- coding: utf-8 -*-
"""
    
    compare_emt_files.py
   ----------------------
   
   This module takes two file paths and performs
   a comparison of the two files, looking for 
   specific differences.

"""

__author__ = "Steven M. Satterfield"

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from compare_emt_files import CompareEMTFiles

class CompareEMTFilesMenu(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.baseFilePath = ''
        self.comparisonFilePath = ''
        
        #compareEMTFiles = CompareEMTFiles()
        
        layout = QFormLayout()
       
        label = QLabel("Compare EMT Files")
       
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Expanding)
        separator.setLineWidth(3)
        
        bFileButton = QPushButton("Select Base File")
        cFileButton = QPushButton("Select Comparison File")
        
        exButton = QPushButton("Compare EMT Files")
        cxButton = QPushButton("Cancel")
        
        layout.addWidget(label)
        layout.addWidget(separator)
        
        layout.addWidget(bFileButton)
        layout.addWidget(cFileButton)
        
        layout.addWidget(separator)
        layout.addWidget(exButton)
        layout.addWidget(cxButton)
       
        self.setLayout(layout)
        
        bFileButton.clicked.connect(self.getBaseFilePath)
        cFileButton.clicked.connect(self.getComparisonFilePath)
        exButton.clicked.connect(self.compareEMTFiles)
        cxButton.clicked.connect(self.close)

    def getBaseFilePath(self):
        fdlg = QFileDialog()
        fdlg.setFileMode(QFileDialog.AnyFile)
        self.baseFilePath = fdlg.getOpenFileName(None, "Select Base EMT File", "", "Excel Files (*.emt)")


    def getComparisonFilePath(self):
        fdlg = QFileDialog()
        fdlg.setFileMode(QFileDialog.AnyFile)
        self.comparisonFilePath = fdlg.getOpenFileName(None, "Select Base EMT File", "", "Excel Files (*.emt)")
        
        
    def compareEMTFiles(self):
        if len(self.baseFilePath) > 1 and len(self.comparisonFilePath) > 1:
            CompareEMTFiles(self.baseFilePath, self.comparisonFilePath)

        
app = QApplication(sys.argv)
dialog = CompareEMTFilesMenu()
dialog.show()
app.exec_()
       