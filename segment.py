# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 17:26:03 2020

@author: Steve
"""

class Segment():
    def __init__(self):
        self._segment_number = ''
        self._attributes = []
        self._bfile = ''
        self._cfile = ''
        self._hasDifferences = False
        
        
    def set_segment_number(self, segment_number):
        self._segment_number = segment_number
        
        
    def get_segment_number(self):
        return self._segment_number
    
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
        
    def get_attributes(self):
        return self._attributes
    
    
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
    

    def claimBaseAttributes(self):
        for attribute in self.get_attributes():
            attribute.set_bfile(True)
    
    
    def findAttribute(self, attributeName):
        for attribute in self.get_attributes():
            if attribute.get_name() == attributeName:
                return attribute
            
        return []


    def sync_attributes(self, comparisonObj):
        localDifferences = False
        
        self.claimBaseAttributes()
        
        for cAttribute in comparisonObj.get_attributes():
            cAttribute.set_cfile(True)
            bAttribute = self.findAttribute(cAttribute.get_name())
            
            if bAttribute:
               bAttribute.set_cfile(True)
               bAttribute.set_cvalue(cAttribute.get_cvalue())
               if bAttribute.get_value() != bAttribute.get_cvalue():
                   bAttribute.set_hasDifferences(True)
                   localDifferences = True
            else:
                self.set_hasDifferences(True)
                self.add_attribute(cAttribute)

        return localDifferences
    