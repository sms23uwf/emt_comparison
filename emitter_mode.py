# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:39:17 2020

@author: Steve
"""

class EmitterMode():
    def __init__(self):
        self._mode_name = ''
        self._attributes = []
        self._generators = []
        
    def set_mode_name(self, mode_name):
        self._mode_name = mode_name
        
    def get_name(self):
        return self._mode_name
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
    def get_attributes(self):
        return self._attributes
    
    def add_generator(self, _generator):
        self._generators.append(_generator)
        
    def get_generators(self):
        return self._generators