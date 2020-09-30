# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 17:43:06 2020

@author: Steve
"""
import print_utility
import constant

class Attribute():
    def __init__(self):
        self._name = ''
        self._value = ''
        self._cvalue = ''
        self._bfile = False
        self._cfile = False
        self._hasDifferences = False
        
        
    def set_name(self, name):
        self._name = name
        
        
    def get_name(self):
        return self._name
    
    
    def set_value(self, value):
        self._value = value
        
        
    def get_value(self):
        return self._value
    
    
    def set_cvalue(self, value):
        self._cvalue = value
        
        
    def get_cvalue(self):
        return self._cvalue


    def set_bfile(self, _file):
        self._bfile = _file

        
    def get_bfile(self):
        return self._bfile            

    
    def set_cfile(self, _file):
        self._cfile = _file

        
    def get_cfile(self):
        return self._cfile            

        
    def set_hasDifferences(self, hasDifferences):
        self._hasDifferences = hasDifferences

        
    def get_hasDifferences(self):
        return self._hasDifferences
    
    
    def print_attribute(self, ws, wsRow, baseLabelCol, baseValueCol, comparisonLabelCol, comparisonValueCol):
        print_utility.writeLabelCell(ws, wsRow, baseLabelCol, constant.XL_MISSING_TEXT if self.get_bfile() == False else self.get_name())
        print_utility.writeValueCell(ws, wsRow, baseValueCol, '' if self.get_bfile() == False else self.get_value())
        print_utility.writeLabelCell(ws, wsRow, comparisonLabelCol, constant.XL_MISSING_TEXT if self.get_cfile() == False else self.get_name())
        print_utility.writeValueCell(ws, wsRow, comparisonValueCol, '' if self.get_cfile() == False else self.get_cvalue())     
        
    def to_dict(self):
        return {
            'name': self._name,
            'value': self._value,
            'cvalue': self._cvalue
        }
    