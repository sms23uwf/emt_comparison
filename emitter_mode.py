# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 12:39:17 2020

@author: Steve
"""

class EmitterMode():
    def __init__(self):
        self._mode_name = ''
        self._attributes = []
        self._generators = []
        self._bfile = ''
        self._cfile = ''
        self._hasDifferences = False

        
    def set_mode_name(self, mode_name):
        self._mode_name = mode_name

        
    def get_name(self):
        return self._mode_name

    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)

        
    def get_attributes(self):
        return self._attributes

    
    def add_generator(self, _generator):
        self._generators.append(_generator)

        
    def get_generators(self):
        return self._generators

    
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


    def findGenerator(self, generatorNumber):
            for generator in self.get_generators():
                if generator.get_generator_number() == generatorNumber:
                    return generator
              
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


    def sync_generators(self, comparisonObj):
        
        localAttrDifferences = False
        localPRISeqDifferences = False
        localFREQSeqDifferences = False
        localDifferences = False
        
        for cGenerator in enumerate(comparisonObj.get_generators()):
            baseGenerator = self.findGenerator(cGenerator.get_generator_number())
                    
            if baseGenerator:
                baseGenerator.set_cfile(comparisonObj.get_cfile())
                localPRISeqDifferences = baseGenerator.sync_priSequences(cGenerator)
                localFREQSeqDifferences = baseGenerator.syn_freqSequences(cGenerator)
                localAttrDifferences = baseGenerator.sync_attributes(cGenerator)
                
            else:
                cGenerator.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_generator(cGenerator)
                
        if localAttrDifferences == True or localPRISeqDifferences == True or localFREQSeqDifferences == True:
            localDifferences = True
            
        return localDifferences
    
