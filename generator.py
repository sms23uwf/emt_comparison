# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:49:09 2020

@author: Steve
"""

class Generator():
    def __init__(self):
        self._generator_number = ''
        self._attributes = []
        self._pri_sequences = []
        self._freq_sequences = []
        self._bfile = ''
        self._cfile = ''

        
    def set_generator_number(self, generator_number):
        self._generator_number = generator_number

        
    def get_generator_number(self):
        return self._generator_number

    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)

        
    def get_attributes(self):
        return self._attributes

    
    def add_pri_sequence(self, _pri_sequence):
        self._pri_sequences.append(_pri_sequence)

        
    def get_pri_sequences(self):
        return self._pri_sequences

    
    def add_freq_sequence(self, _freq_sequence):
        self._freq_sequences.append(_freq_sequence)

        
    def get_freq_sequences(self):
        return self._freq_sequences

    
    def set_bfile(self, _file):
        self._bfile = _file

        
    def get_bfile(self):
        return self._bfile            

    
    def set_cfile(self, _file):
        self._cfile = _file

        
    def get_cfile(self):
        return self._cfile            
    
    
    def findAttribute(self, attributeName):
        for attribute in self.get_attributes():
            if attribute.get_name() == attributeName:
                return attribute
            
        return []
    
    
    def sync_attributes(self, comparisonObj):
        for cAttribute in enumerate(comparisonObj.get_attributes()):
            bAttribute = self.findAttribute(cAttribute.get_name())
            
            if bAttribute:
               bAttribute.set_cfile(comparisonObj.get_cfile())
               bAttribute.set_cvalue(cAttribute.get_cvalue())
            else:
                self.add_attribute(cAttribute)


    