# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 17:39:36 2020

@author: Steve
"""

import print_utility
import constant
import list_utility
from emitter_iterator import EmitterIterator

from itertools import filterfalse


class Emitter():
    
    def __init__(self):
        self._elnot = ''
        self._attributes = []
        self._modes = []
        self._bfile = False
        self._cfile = False
        self._hasDifferences = False
        
        
    def set_elnot(self, elnot):
        self._elnot = elnot
        
        
    def get_elnot(self):
        return self._elnot

    
    def add_attribute(self, _attribute):
        self._attributes.append(_attribute)


    def set_attributes(self, attributes):
        self._attributes = []
        for attribute in attributes:
            self.add_attribute(attribute)

        
    def get_attributes(self):
        return self._attributes

    
    def add_mode(self, _mode):
        self._modes.append(_mode)

        
    def set_modes(self, modes):
        self._modes = []
        for mode in modes:
            self.add_mode(mode)


    def get_modes(self):
        return self._modes        

    
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
    
    
    def claimBaseAttributes(self):
        for attribute in self.get_attributes():
            attribute.set_bfile(True)
            

    def claimBaseModes(self):
        for emitterMode in self.get_modes():
            emitterMode.set_bfile(True)

            
    def findEmitterMode(self, emitterModeName):
            for emitterMode in self.get_modes():
                if emitterMode.get_name() == emitterModeName:
                    return emitterMode
              
            return []

    
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
    
    
    def clean_modes(self):
        
        self.set_modes(list(filterfalse(list_utility.filtertrue, self.get_modes())))
        
        
        # for mode in self.get_modes():
        #     if mode.get_hasDifferences() == False:
        #         self.get_modes().remove(mode)
        #     else:
        #         mode.clean_attributes()
        #         mode.clean_generators()
        
      
    def attributes_to_dict(self):
        attribute_holder = []
        for attribute in self.get_attributes():
            attribute_holder.append(attribute)
            
        self._attributes = []
        
        for attribute in attribute_holder:
            self.add_attribute(attribute.__dict__)
            
            
    def modes_to_dict(self):
        mode_holder = []
        for mode in self.get_modes():
            mode_holder.append(mode)
            
        self._modes = []
        
        for mode in mode_holder:
            mode.attributes_to_dict()
            mode.generators_to_dict()
            self.add_mode(mode.__dict__)
            
            
    def sync_modes(self, comparisonObj):
        localAttrDifferences = False
        localGeneratorDifferences = False
        localDifferences = False
        
        self.claimBaseModes()
        
        for cMode in comparisonObj.get_modes():
            cMode.set_cfile(True)
            baseMode = self.findEmitterMode(cMode.get_name())
            if baseMode:
                baseMode.set_cfile(True)
                localAttrDifferences = baseMode.sync_attributes(cMode)
                localGeneratorDifferences = baseMode.sync_generators(cMode)
                if localAttrDifferences == True or localGeneratorDifferences == True:
                    localDifferences = True
            else:
                localDifferences = True
                cMode.set_hasDifferences(True)
                self.set_hasDifferences(True)
                self.add_mode(cMode)
            
        return localDifferences
    
    
    def print_emitter(self, ws, baseLabelCol, baseValueCol, comparisonLabelCol, comparisonValueCol):
        print_utility.writeLabelCell(ws, print_utility.wsEmittersRow, baseLabelCol, constant.XL_MISSING_TEXT if self.get_bfile() == False else "ELNOT:")
        print_utility.writeValueCell(ws, print_utility.wsEmittersRow, baseValueCol, '' if self.get_bfile() == False else self.get_elnot())
        print_utility.writeLabelCell(ws, print_utility.wsEmittersRow, comparisonLabelCol, constant.XL_MISSING_TEXT if self.get_cfile() == False else "ELNOT:")
        print_utility.writeValueCell(ws, print_utility.wsEmittersRow, comparisonValueCol, '' if self.get_cfile() == False else self.get_elnot())
        print_utility.wsEmittersRow += 1
        
        if self.get_bfile() == True and self.get_bfile() == True:
            self.print_attribute_differences(ws)

        
    
    def print_attribute_differences(self, ws):
        for baseAttribute in self.get_attributes():
            if baseAttribute.get_hasDifferences() == True:
                baseAttribute.print_attribute(ws, print_utility.wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL)
                print_utility.wsEmittersRow += 1
    
    
    def to_dict(self):
        return {
        'elnot': self._elnot,
        'attributes': self._attributes,
        'modes': self._modes
       }
    
    
    def __iter__(self):
        return EmitterIterator(self)