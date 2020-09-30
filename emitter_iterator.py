# -*- coding: utf-8 -*-
"""
    
    emitter_iterator.py
   ----------------------
   
   This module iteration capability for the Emitter class.

"""

__author__ = "Steven M. Satterfield"


class EmitterIterator:
    def __init__(self, emitter):
        self._emitter = emitter
        self._index = 0
        
    
    def __next__(self):
        if self._index < (len(self._emitter._attributes) + len(self._emitter._modes)):
            if self._index < len(self._emitter._attributes):
                result = (self._emitter._attributes[self._index], 'Attribute')
            else:
                result = (self._emitter._modes[self._index - len(self._emitter._attributes)], 'Mode')
                
            self._index += 1
            return result
        
        #end of iteration
        raise StopIteration
        