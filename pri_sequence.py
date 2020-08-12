# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 17:20:09 2020

@author: Steve
"""

class Pri_Sequence():
    def __init__(self):
        self._ordinal_pos = ''
        self._number_of_segments = ''
        self._attributes = []
        self._segments = []
        
    def set_ordinal_pos(self, ordinal_pos):
        self._ordinal_pos = ordinal_pos
        
    
    def get_ordinal_pos(self):
        return self._ordinal_pos
    
    
    def set_number_of_segments(self, number_of_segments):
        self._number_of_segments = number_of_segments
        
    
    def get_number_of_segments(self):
        return self._number_of_segments
    
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
    
    def get_attributes(self):
        return self._attributes

    
    def add_segment(self, _segment):
        self._segments.append(_segment)
        

    def get_segments(self):
        return self._segments    