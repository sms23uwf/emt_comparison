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