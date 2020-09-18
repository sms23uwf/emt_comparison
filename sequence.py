# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 17:20:09 2020

@author: Steve
"""

import print_utility
import constant

class Sequence():
    def __init__(self):
        self._ordinal_pos = ''
        self._number_of_segments = ''
        self._attributes = []
        self._segments = []
        self._bfile = False
        self._cfile = False
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
    
    
    def claimBaseAttributes(self):
        for attribute in self.get_attributes():
            attribute.set_bfile(True)


    def claimBaseSegments(self):
        for segment in self.get_segments():
            segment.set_bfile(True)

    
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
        
        self.claimBaseAttributes()
        
        for cAttribute in comparisonObj.get_attributes():
            cAttribute.set_cfile(True)
            bAttribute = self.findAttribute(cAttribute.get_name())
            
            if bAttribute:
               bAttribute.set_cfile(True)
               bAttribute.set_cvalue(cAttribute.get_cvalue())
               if bAttribute.get_value() != bAttribute.get_cvalue():
                   self.set_hasDifferences(True)
                   bAttribute.set_hasDifferences(True)
                   localDifferences = True
            else:
                localDifferences = True
                self.set_hasDifferences(True)
                self.add_attribute(cAttribute)

        return localDifferences


    def sync_segments(self, comparisonObj):
        localDifferences = False
        
        self.claimBaseSegments()
        
        for cSegment in comparisonObj.get_segments():
            cSegment.set_cfile(True)
            bSegment = self.findSegmentBySegmentNumber(cSegment.get_segment_number())
            
            if bSegment:
                bSegment.set_cfile(True)
                localDifferences = bSegment.sync_attributes(cSegment)
                if localDifferences == True:
                    self.set_hasDifferences(True)
                    bSegment.set_hasDifferences(True)
            else:
                localDifferences = True
                cSegment.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_segment(cSegment)
                
        return localDifferences


    def print_pri_sequence(self, ws, elnot, modeName, generatorNumber):
        self.print_sequence(ws, print_utility.wsPRISequencesRow, elnot, modeName, generatorNumber, "PRI_SEQUENCE:")
        print_utility.wsPRISequencesRow += 1
        
        if self.get_bfile() == True and self.get_bfile() == True:
            self.print_attribute_differences_pri(ws)
        
        
    def print_freq_sequence(self, ws, elnot, modeName, generatorNumber):
        self.print_sequence(ws, print_utility.wsFREQSequencesRow, elnot, modeName, generatorNumber, "FREQ_SEQUENCE:")
        print_utility.wsFREQSequencesRow += 1

        if self.get_bfile() == True and self.get_bfile() == True:
            self.print_attribute_differences_freq(ws)
        

    def print_sequence(self, ws, wsRow, elnot, modeName, generatorNumber, label):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_VAL, elnot)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_VAL, elnot)

        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT if self.get_bfile() == False else label)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, '' if self.get_bfile() == False else self.get_ordinal_pos())
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT if self.get_cfile() == False else label)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, '' if self.get_cfile() == False else self.get_ordinal_pos())


    def print_attribute_differences_pri(self, ws):
        for baseAttribute in self.get_attributes():
            if baseAttribute.get_hasDifferences() == True:
                baseAttribute.print_attribute(ws, print_utility.wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL)
                print_utility.wsPRISequencesRow += 1


    def print_attribute_differences_freq(self, ws):
        for baseAttribute in self.get_attributes():
            if baseAttribute.get_hasDifferences() == True:
                baseAttribute.print_attribute(ws, print_utility.wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL)
                print_utility.wsFREQSequencesRow += 1
