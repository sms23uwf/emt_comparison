# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:49:09 2020

@author: Steve
"""

import print_utility
import constant
import list_utility
from itertools import filterfalse


class Generator():
    def __init__(self):
        self._generator_number = ''
        self._attributes = []
        self._pri_sequences = []
        self._freq_sequences = []
        self._bfile = False
        self._cfile = False
        self._hasDifferences = False

        
    def set_generator_number(self, generator_number):
        self._generator_number = generator_number

        
    def get_generator_number(self):
        return self._generator_number

    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)

        
    def set_attributes(self, attributes):
        self._attributes = []
        for attribute in attributes:
            self.add_attribute(attribute)


    def get_attributes(self):
        return self._attributes

    
    def add_pri_sequence(self, _pri_sequence):
        self._pri_sequences.append(_pri_sequence)

        
    def set_pri_sequences(self, pri_sequences):
        self._pri_sequences = []
        for seq in pri_sequences:
            self.add_pri_sequence(seq)


    def get_pri_sequences(self):
        return self._pri_sequences

    
    def add_freq_sequence(self, _freq_sequence):
        self._freq_sequences.append(_freq_sequence)


    def set_freq_sequences(self, freq_sequences):
        self._freq_sequences = []
        for seq in freq_sequences:
            self.add_freq_sequence(seq)

        
    def get_freq_sequences(self):
        return self._freq_sequences

    
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


    def claimBasePRISequences(self):
        for sequence in self.get_pri_sequences():
            sequence.set_bfile(True)


    def claimBaseFREQSequences(self):
        for sequence in self.get_freq_sequences():
            sequence.set_bfile(True)


    def findAttribute(self, attributeName):
        for attribute in self.get_attributes():
            if attribute.get_name() == attributeName:
                return attribute
            
        return []
    
    
    def findPRISequenceByOrdinalPos(self, ordinalPos):
        for sequence in self.get_pri_sequences():
            if sequence.get_ordinal_pos() == ordinalPos:
                return sequence
         
        return []
    
    
    def findFREQSequenceByOrdinalPos(self, ordinalPos):
        for sequence in self.get_freq_sequences():
            if sequence.get_ordinal_pos() == ordinalPos:
                return sequence 
            
        return []

    
    def attributes_to_dict(self):
        attribute_holder = []
        for attribute in self.get_attributes():
            attribute_holder.append(attribute)
            
        self._attributes = []
        
        for attribute in attribute_holder:
            self.add_attribute(attribute.__dict__)


    def clean_attributes(self):
        self.set_attributes(list(filterfalse(list_utility.filtertrue, self.get_attributes())))
        
        # for attribute in self.get_attributes():
        #     if attribute.get_hasDifferences() == False:
        #         self.get_attributes().remove(attribute)


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


    def sequences_to_dict(self):
        pHolder = []
        for obj in self.get_pri_sequences():
            pHolder.append(obj)
            
        self._pri_sequences = []
        
        for obj in pHolder:
            obj.attributes_to_dict()
            obj.segments_to_dict()
            self.add_pri_sequence(obj.__dict__)

        fHolder = []
        for obj in self.get_freq_sequences():
            fHolder.append(obj)
            
        self._freq_sequences = []
        
        for obj in fHolder:
            obj.attributes_to_dict()
            obj.segments_to_dict()
            self.add_freq_sequence(obj.__dict__)


    def clean_sequences(self):
        self.set_pri_sequences(list(filterfalse(list_utility.filtertrue, self.get_pri_sequences())))
        self.set_freq_sequences(list(filterfalse(list_utility.filtertrue, self.get_freq_sequences())))

        # for pri_sequence in self.get_pri_sequences():
        #     if pri_sequence.get_hasDifferences() == False:
        #         self.get_pri_sequences().remove(pri_sequence)
        #     else:
        #         pri_sequence.clean_attributes()
        #         pri_sequence.clean_segments()
                
                        
        # for freq_sequence in self.get_freq_sequences():
        #     if freq_sequence.get_hasDifferences() == False:
        #         self.get_freq_sequences().remove(freq_sequence)
        #     else:
        #         freq_sequence.clean_attributes()
        #         freq_sequence.clean_segments()
                
                                
        
    def sync_priSequences(self, comparisonObj):
        
        localAttrDifferences = False
        localSegmentDifferences = False
        localDifferences = False
        
        self.claimBasePRISequences()
        
        for cSequence in comparisonObj.get_pri_sequences():
            cSequence.set_cfile(True)
            bSequence = self.findPRISequenceByOrdinalPos(cSequence.get_ordinal_pos())
            
            if bSequence:
                bSequence.set_cfile(True)
                localSegmentDifferences = bSequence.sync_segments(cSequence)
                localAttrDifferences = bSequence.sync_attributes(cSequence)
                if localAttrDifferences == True or localSegmentDifferences == True:
                    self.set_hasDifferences(True)
                    bSequence.set_hasDifferences(True)
                    localDifferences = True
            else:
                localDifferences = True
                cSequence.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_pri_sequence(cSequence)
                
        return localDifferences
    

    def sync_freqSequences(self, comparisonObj):
        
        localAttrDifferences = False
        localSegmentDifferences = False
        localDifferences = False

        self.claimBaseFREQSequences()
        
        for cSequence in comparisonObj.get_freq_sequences():
            cSequence.set_cfile(True)
            bSequence = self.findFREQSequenceByOrdinalPos(cSequence.get_ordinal_pos())
            
            if bSequence:
                bSequence.set_cfile(True)
                localSegmentDifferences = bSequence.sync_segments(cSequence)
                localAttrDifferences = bSequence.sync_attributes(cSequence)
                if localAttrDifferences == True or localSegmentDifferences == True:
                    self.set_hasDifferences(True)
                    bSequence.set_hasDifferences(True)
                    localDifferences = True
            else:
                localDifferences = True
                cSequence.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_pri_sequence(cSequence)
                
        return localDifferences
               

    def print_generator(self, ws, elnot, modeName):
        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_VAL, elnot)
        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_VAL, elnot)

        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT if self.get_bfile() == False else "GENERATOR:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_VAL, '' if self.get_bfile() == False else self.get_generator_number())
        print_utility.writeLabelCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT if self.get_cfile() == False else "GENERATOR:")
        print_utility.writeValueCell(ws, print_utility.wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_VAL, '' if self.get_cfile() == False else self.get_generator_number())
        
        print_utility.wsGeneratorsRow += 1

        if self.get_bfile() == True and self.get_bfile() == True:
            self.print_attribute_differences(ws)
        
        
        
    def print_attribute_differences(self, ws):
        for baseAttribute in self.get_attributes():
            if baseAttribute.get_hasDifferences() == True:
                baseAttribute.print_attribute(ws, print_utility.wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL)
                print_utility.wsGeneratorsRow += 1
        
        
    def to_dict(self):
        return {
            'generator_number': self._generator_number,
            'attributes': self._attributes,
            'pri_sequences': self._pri_sequences,
            'freq_sequences': self._freq_sequences
        }