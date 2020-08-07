# -*- coding: utf-8 -*-
"""
    
    compare_emt_files.py
   ----------------------
   
   This module takes two file paths and performs
   a comparison of the two files, looking for 
   specific differences.

"""

__author__ = "Steven M. Satterfield"

import constant
from emitter import Emitter
from attribute import Attribute
from emitter_mode import EmitterMode
from generator import Generator



writeFile = 'diff4.txt'
baseFileName = 'SPECTRE_OUTPUT_07032020.emt'
comparisonFileName = 'SPECTRE_OUTPUT_08032020.emt'
  
base_emitters = []
comparison_emitters = []

currentEntity = constant.EMITTER

def lineAsKeyValue(line):
    line_kv = line.split(constant.VALUE_SEPARATOR)
    return line_kv

    
def parseBaseFile():
    emitter = Emitter()
    emitter_mode = EmitterMode()
    generator = Generator()
    
    passNumber = 0
    modePass = 0
    generatorPass = 0
    
    with open(baseFileName) as f1:
        for cnt, line in enumerate(f1):
            if line.strip() == constant.EMITTER:
                currentEntity = constant.EMITTER
                if passNumber > 0:
                    base_emitters.append(emitter)
                    
                emitter = Emitter()
                passNumber += 1
                modePass = 0
                generatorPass = 0
                
            elif line.strip() == constant.EMITTER_MODE:
                currentEntity = constant.EMITTER_MODE
                if modePass > 0:
                    emitter.add_mode(emitter_mode)
                    
                emitter_mode = EmitterMode()
                modePass += 1
                generatorPass = 0
                
            elif line.strip() == constant.GENERATOR:
                currentEntity = constant.GENERATOR
                if generatorPass > 0:
                    emitter_mode.add_generator(generator)
                    
                generator = Generator()
                generatorPass += 1
                
            else:
                if line.strip().__contains__(constant.VALUE_SEPARATOR):
                    line_kv = lineAsKeyValue(line.strip())
                    line_key = line_kv[0].strip()
                    line_value = line_kv[1].strip()
                    
                    if currentEntity == constant.EMITTER:
 
                        if line_key.strip() == constant.EMITTER_ELNOT:
                            emitter.set_elnot(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            emitter.add_attribute(attrib)
                    
                    elif currentEntity == constant.EMITTER_MODE:
                   
                        if line_key.strip() == constant.MODE_NAME:
                            emitter_mode.set_mode_name(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            emitter_mode.add_attribute(attrib)
                            
                    elif currentEntity == constant.GENERATOR:

                        if line_key.strip() == constant.GENERATOR_NUMBER:
                            generator.set_generator_number(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            generator.add_attribute(attrib)
        else:
            base_emitters.append(emitter)
        
    
def parseComparisonFile():
    emitter = Emitter()
    emitter_mode = EmitterMode()
    generator = Generator()
    
    passNumber = 0
    modePass = 0
    generatorPass = 0
    
    with open(comparisonFileName) as f2:
        for cnt, line in enumerate(f2):
            if line.strip() == constant.EMITTER:
                currentEntity = constant.EMITTER
                if passNumber > 0:
                    comparison_emitters.append(emitter)
                    
                emitter = Emitter()
                passNumber += 1
                modePass = 0
                generatorPass = 0
                
            elif line.strip() == constant.EMITTER_MODE:
                currentEntity = constant.EMITTER_MODE
                if modePass > 0:
                    emitter.add_mode(emitter_mode)
                    
                emitter_mode = EmitterMode()
                modePass += 1
                generatorPass = 0
                
            elif line.strip() == constant.GENERATOR:
                currentEntity = constant.GENERATOR
                if generatorPass > 0:
                    emitter_mode.add_generator(generator)
                    
                generator = Generator()
                generatorPass += 1
                
            else:
                if line.strip().__contains__(constant.VALUE_SEPARATOR):
                    line_kv = lineAsKeyValue(line.strip())
                    line_key = line_kv[0].strip()
                    line_value = line_kv[1].strip()
               
                    if currentEntity == constant.EMITTER:
 
                        if line_key.strip() == constant.EMITTER_ELNOT:
                            emitter.set_elnot(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            emitter.add_attribute(attrib)
                    
                    elif currentEntity == constant.EMITTER_MODE:
                   
                        if line_key.strip() == constant.MODE_NAME:
                            emitter_mode.set_mode_name(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            emitter_mode.add_attribute(attrib)
                            
                    elif currentEntity == constant.GENERATOR:

                        if line_key.strip() == constant.GENERATOR_NUMBER:
                            generator.set_generator_number(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            generator.add_attribute(attrib)
        else:
            comparison_emitters.append(emitter)
            
                    
def findElnot(elnotValue, inArray):
    print("elnotValue: {}, inArray: {}".format(elnotValue, inArray))
    if inArray == constant.COMPARISON_ARRAY:
        for emitter in comparison_emitters:
            if emitter.get_elnot() == elnotValue:
                print("just about to return emitter for elnot: {}".format(elnotValue))
                return emitter
        
        print("just about to return empty array for search for elnot {}".format(elnotValue))
        return []
    else:
        for emitter in base_emitters:
            if emitter.get_elnot() == elnotValue:
                print("just about to return emitter for elnot: {}".format(elnotValue))
                return emitter
            
        print("just about to return empty array for search for elnot {}".format(elnotValue))
        return []


def findEmitterMode(emitterModeName, comparisonEmitter):
        for comparisonEmitterMode in comparisonEmitter.get_modes():
            if comparisonEmitterMode.get_name() == emitterModeName:
                return comparisonEmitterMode
            
        return []
    

def findAttribute(baseAttributeName, comparisonEmitter):
    for comparisonAttribute in comparisonEmitter.get_attributes():
        if comparisonAttribute.get_name() == baseAttributeName:
            return comparisonAttribute
        
    return []

        
    
def compareEMTFiles():
    wf = open(writeFile, "w+")
    
    print("parsing Base File ...")
    parseBaseFile()
    
    print("parsing Comparison File ...")
    parseComparisonFile()
    
    for baseEmitter in base_emitters:
        bElnot = baseEmitter.get_elnot()
        comparisonEmitter = findElnot(bElnot, constant.COMPARISON_ARRAY)
        
        if not comparisonEmitter:
            wf.write("Base file contains emitter with elnot: {} that is not found in Comparison file.\n".format(bElnot))
        else:
            for baseAttribute in baseEmitter.get_attributes():
                #print("baseAttribute: {} = {}".format(baseAttribute.get_name(), baseAttribute.get_value()))
                comparisonAttribute = findAttribute(baseAttribute.get_name(), comparisonEmitter)
                
                if not comparisonAttribute:
                    wf.write("Base file contains emitter {} with attribute {} that is missing from Comparison file.\n".format(bElnot, baseAttribute.get_name()))
                
            for baseEmitterMode in baseEmitter.get_modes():
                comparisonEmitterMode = findEmitterMode(baseEmitterMode.get_name(), baseEmitter)
                
                if not comparisonEmitterMode:
                    wf.write("Base file emitter {} contains emitterMode {} that is missing from Comparison file.\n".format(bElnot, baseEmitterMode.get_name()))
                
     
    for comparisonEmitter in comparison_emitters:
        cElnot = comparisonEmitter.get_elnot()
        baseEmitter = findElnot(cElnot, constant.BASE_ARRAY)

        if not baseEmitter:
            wf.write("Comparison file contains emitter with elnot: {} that is not found in Base file.\n".format(cElnot))
        else:
            for comparisonAttribute in comparisonEmitter.get_attributes():
                #print("comparisonAttribute: {} = {}".format(comparisonAttribute.get_name(), comparisonAttribute.get_value()))
                baseAttribute = findAttribute(comparisonAttribute.get_name(), baseEmitter)
                
                if not baseAttribute:
                    wf.write("Comparison file contains emitter {} with attribute {} that is missing from Base file.\n".format(cElnot, comparisonAttribute.get_name()))
                    
            for comparisonEmitterMode in comparisonEmitter.get_modes():
                baseEmitterMode = findEmitterMode(comparisonEmitterMode.get_name(), comparisonEmitter)
                
                if not baseEmitterMode:
                    wf.write("Comparison file emitter {} contains emitterMode {} that is missing from Base file.\n".format(cElnot, comparisonEmitterMode.get_name()))
                    
        
    wf.close()
                   
if __name__ == '__main__':
    compareEMTFiles()
    
                    
                
                
                       
                   

       