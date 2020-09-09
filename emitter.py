# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 17:39:36 2020

@author: Steve
"""

class Emitter():
    def __init__(self):
        self._elnot = ''
        self._attributes = []
        self._modes = []
        self._bfile = ''
        self._cfile = ''
        self._hasDifferences = False
        
    def set_elnot(self, elnot):
        self._elnot = elnot
        
    def get_elnot(self):
        return self._elnot
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
    def get_attributes(self):
        return self._attributes
    
    def add_mode(self, _mode):
        self._modes.append(_mode)
        
    def get_modes(self):
        return self._modes        
    
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
    
    def findAttribute(self, attributeName):
        for attribute in self.get_attributes():
            if attribute.get_name() == attributeName:
                return attribute
            
        return []
    
    
    def findEmitterMode(self, emitterModeName):
            for emitterMode in self.get_modes():
                if emitterMode.get_name() == emitterModeName:
                    return emitterMode
              
            return []

    
    def sync_attributes(self, comparisonObj):
        for cAttribute in enumerate(comparisonObj.get_attributes()):
            bAttribute = self.findAttribute(cAttribute.get_name())
            
            if bAttribute:
               bAttribute.set_cfile(comparisonObj.get_cfile())
               bAttribute.set_cvalue(cAttribute.get_cvalue())
            else:
                self.add_attribute(cAttribute)

    
    def sync_modes(self, comparisonObj):
        for cMode in enumerate(comparisonObj.get_modes()):
            baseMode = self.findEmitterMode(cMode.get_name())
            if baseMode:
                baseMode.set_cfile(comparisonObj.get_cfile())
                baseMode.sync_attributes(cMode)
                baseMode.sync_generators(cMode)
