# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 17:20:09 2020

@author: Steve
"""

class Sequence():
    def __init__(self):
        self._ordinal_pos = ''
        self._number_of_segments = ''
        self._attributes = []
        self._segments = []
        self._bfile = ''
        self._cfile = ''
        self._hasDifferences = False
        
        
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


    def findSegmentBySegmentNumber(self, segmentNumber):
        for segment in self.get_segments():
            if segment.get_segment_number() == segmentNumber:
                return segment
            
        return []


    def sync_attributes(self, comparisonObj):
        localDifferences = False
        
        for cAttribute in enumerate(comparisonObj.get_attributes()):
            bAttribute = self.findAttribute(cAttribute.get_name())
            
            if bAttribute:
               bAttribute.set_cfile(comparisonObj.get_cfile())
               bAttribute.set_cvalue(cAttribute.get_cvalue())
               if bAttribute.get_value() != bAttribute.get_cvalue():
                   bAttribute.set_hasDifferences(True)
                   localDifferences = True
            else:
                self.set_hasDifferences(True)
                self.add_attribute(cAttribute)

        return localDifferences


    def sync_segments(self, comparisonObj):
        localDifferences = False
        
        for cSegment in enumerate(comparisonObj.get_segments()):
            bSegment = self.findSegmentBySegmentNumber(cSegment.get_segment_number())
            
            if bSegment:
                bSegment.set_cfile(comparisonObj.get_cfile())
                localDifferences = bSegment.sync_attributes(cSegment)
            else:
                cSegment.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_segment(cSegment)
                
        return localDifferences
