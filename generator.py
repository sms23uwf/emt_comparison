# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:49:09 2020

@author: Steve
"""

import constant


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
         