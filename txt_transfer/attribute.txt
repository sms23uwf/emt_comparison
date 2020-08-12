# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 17:43:06 2020

@author: Steve
"""

class Attribute():
    def __init__(self):
        self._name = ''
        self._value = ''
        
    def set_name(self, name):
        self._name = name
        
    def get_name(self):
        return self._name
    
    def set_value(self, value):
        self._value = value
        
    def get_value(self):
        return self._value
    
    