# -*- coding: utf-8 -*-
"""
Created on Wed Aug 26 09:43:24 2020

@author: Steve
"""

class ColorMarkRowBoundary():
    def __init__(self):
        self._topboundary = 0
        self._bottomboundary = 0
        
    def set_topboundary(self, topboundary):
        self._topboundary = topboundary
        
    def get_topboundary(self):
        return self._toprboundary    
    
    def set_bottomboundary(self, bottomboundary):
        self._bottomboundary = bottomboundary
        
    def get_bottomboundary(self):
        return self._bottomrboundary        