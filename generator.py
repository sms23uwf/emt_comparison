# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:49:09 2020

@author: Steve
"""

class Generator():
    def __init__(self):
        self._generator_number = ''
        self._attributes = []
        
    def set_generator_number(self, generator_number):
        self._generator_number = generator_number
        
    def get_generator_number(self):
        return self._generator_number
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
    def get_attributes(self):
        return self._attributes
    