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



writeFile = 'diff.txt'
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
    if inArray == constant.COMPARISON_ARRAY:
        for emitter in comparison_emitters:
            if emitter.get_elnot() == elnotValue:
                return emitter
        
        return []
    else:
        for emitter in base_emitters:
            if emitter.get_elnot() == elnotValue:
                return emitter
            
        return []


def findEmitterMode(emitterModeName, comparisonEmitter):
        for comparisonEmitterMode in comparisonEmitter.get_modes():
            if comparisonEmitterMode.get_name() == emitterModeName:
                return comparisonEmitterMode
          
        return []
    

def findGenerator(generatorNumber, comparisonEmitterMode):
        for comparisonGenerator in comparisonEmitterMode.get_generators():
            if comparisonGenerator.get_generator_number() == generatorNumber:
                print("just about to return generator: {} as match for generator: {}".format(comparisonGenerator.get_generator_number(), generatorNumber))
                return comparisonGenerator
          
        return []


def findAttribute(baseAttributeName, comparisonEmitter):
    for comparisonAttribute in comparisonEmitter.get_attributes():
        if comparisonAttribute.get_name() == baseAttributeName:
            return comparisonAttribute
        
    return []

        
    
def compareEMTFiles(wf, base_emitter_collection, comparison_emitter_collection, comparisonArray, bFileName, cFileName):
     
    for baseEmitter in base_emitter_collection:
        bElnot = baseEmitter.get_elnot()
        comparisonEmitter = findElnot(bElnot, comparisonArray)
        
        if not comparisonEmitter:
            wf.write("{} contains emitter with elnot: {} that is not found in {}.\n".format(bFileName, bElnot, cFileName))
        else:
            for baseAttribute in baseEmitter.get_attributes():
                #print("baseAttribute: {} = {}".format(baseAttribute.get_name(), baseAttribute.get_value()))
                comparisonAttribute = findAttribute(baseAttribute.get_name(), comparisonEmitter)
                
                if not comparisonAttribute:
                    wf.write("{} contains emitter {} with attribute {} that is missing from {}.\n".format(bFileName, bElnot, baseAttribute.get_name(), cFileName))
                else:
                    if comparisonAttribute.get_value() != baseAttribute.get_value():
                        wf.write("{} emitter {} contains attribute: {} with value: {} which is different from {} attribute value of {}.\n".format(bFileName, bElnot, baseAttribute.get_name(), baseAttribute.get_value(), cFileName, comparisonAttribute.get_value()))
                        
            for baseEmitterMode in baseEmitter.get_modes():
                comparisonEmitterMode = findEmitterMode(baseEmitterMode.get_name(), comparisonEmitter)
                
                if not comparisonEmitterMode:
                    wf.write("{} emitter {} contains emitterMode {} that is missing from {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), cFileName))
                else:
                    for baseModeAttribute in baseEmitterMode.get_attributes():
                        comparisonModeAttribute = findAttribute(baseModeAttribute.get_name(), comparisonEmitterMode)
                        
                        if not comparisonModeAttribute:
                            wf.write("{} contains emitter mode {} with attribute {} that is missing from the emitter mode in {}.\n".format(bFileName, baseEmitterMode.get_name(), baseModeAttribute.get_name(), cFileName))
                        else:
                            if comparisonModeAttribute.get_value() != baseModeAttribute.get_value():
                                wf.write("{} emitter mode {} contains attribute {} with value: {} which is different from {} attribute value of {}.\n".format(bFileName, baseEmitterMode.get_name(), baseModeAttribute.get_name(), baseModeAttribute.get_value(), cFileName, comparisonModeAttribute.get_value()))
                                
                    for baseGenerator in baseEmitterMode.get_generators():
                        comparisonGenerator = findGenerator(baseGenerator.get_generator_number(), comparisonEmitterMode)
                    
                    if not comparisonGenerator:
                        wf.write("{} mode {} contains generator: {} that is missing from the mode in {}".format(bFileName, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), cFileName))
                        
                                
                            
 
def compareTheFiles():

    print("parsing Base File ...")
    parseBaseFile()
    
    print("parsing Comparison File ...")
    parseComparisonFile()

    wf = open(writeFile, "w+")

    compareEMTFiles(wf, base_emitters, comparison_emitters, constant.COMPARISON_ARRAY, baseFileName, comparisonFileName)

    compareEMTFiles(wf, comparison_emitters, base_emitters, constant.BASE_ARRAY, comparisonFileName, baseFileName)

    wf.close()

           
if __name__ == '__main__':
    compareTheFiles()
    
                    
                
                
                       
                   

       