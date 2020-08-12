# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 17:26:03 2020

@author: Steve
"""

class Pri_Segment():
    def __init__(self):
        self._segment_number = ''
        self._attributes = []
        
    def set_segment_number(self, segment_number):
        self._segment_number = segment_number
        
    def get_segment_number(self):
        return self._segment_number
    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)
        
    def get_attributes(self):
        return self._attributes
    