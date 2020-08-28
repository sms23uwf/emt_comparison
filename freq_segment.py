# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 18:35:15 2020

@author: Steve
"""

class Freq_Segment():
    def __init__(self):
        self._segment_number = ''
        self._attributes = []
        self._bfile = ''
        self._cfile = ''
        
        
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
    