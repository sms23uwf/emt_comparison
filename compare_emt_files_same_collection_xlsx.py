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
from pri_sequence import Pri_Sequence
from pri_segment import Pri_Segment
from freq_sequence import Freq_Sequence
from freq_segment import Freq_Segment
from color_mark_row_boundary import ColorMarkRowBoundary

import xlwings as xw
from xlwings import constants
from xlwings.utils import rgb_to_int

class CompareEMTFiles():
    def __init__(self, bFilePath, cFilePath):
        self.baseFileName = bFilePath[0]
        self.comparisonFileName = cFilePath[0]
        self.compareTheFiles()
        self.bfDisplay = self.baseFileName
        self.cfDisplay = self.comparisonFileName


    writeFile = 'diff.txt'
      
    base_emitters = []
    comparison_emitters = []
    all_emitters = []
    
    currentEntity = constant.EMITTER
    
    def lineAsKeyValue(self, line):
        line_kv = line.split(constant.VALUE_SEPARATOR)
        return line_kv
    
        
    def parseFile(self, fName, emitter_collection, isBase):
        emitter = Emitter()
        emitter_mode = EmitterMode()
        generator = Generator()
        pri_sequence = Pri_Sequence()
        freq_sequence = Freq_Sequence()
        pri_segment = Pri_Segment()
        freq_segment = Freq_Segment()
            
        passNumber = 0
        modePass = 0
        generatorPass = 0
        priSeqPass = 0
        freqSeqPass = 0
        priSegmentPass = 0
        freqSegmentPass = 0
        
        with open(fName) as f1:
            for cnt, line in enumerate(f1):
                if line.strip() == constant.EMITTER:
                    currentEntity = constant.EMITTER
                    if passNumber > 0:
                        if modePass > 0:
                            emitter.add_mode(emitter_mode)
                            
                        emitter_collection.append(emitter)
                        modePass = 0
                        
                    emitter = Emitter()
                    passNumber += 1
                    
                elif line.strip() == constant.EMITTER_MODE:
                    currentEntity = constant.EMITTER_MODE
                    if modePass > 0:
                        if generatorPass > 0:
                            emitter_mode.add_generator(generator)
                            
                        emitter.add_mode(emitter_mode)
                        generatorPass = 0
                        
                    emitter_mode = EmitterMode()
                    modePass += 1
                    
                elif line.strip() == constant.GENERATOR:
                    currentEntity = constant.GENERATOR
                    if generatorPass > 0:
                        if priSeqPass > 0:
                            generator.add_pri_sequence(pri_sequence)
    
                        if freqSeqPass > 0:
                            generator.add_freq_sequence(freq_sequence)
    
                        emitter_mode.add_generator(generator)
                        priSeqPass = 0
                        freqSeqPass = 0
                        
                    generator = Generator()
                    generatorPass += 1
                    
                elif line.strip() == constant.PRI_SEQUENCE:
                    currentEntity = constant.PRI_SEQUENCE
                    if priSeqPass > 0:
                        if priSegmentPass > 0:
                            pri_sequence.add_segment(pri_segment)
                        
                        generator.add_pri_sequence(pri_sequence)
                        priSegmentPass = 0
                        
                    pri_sequence = Pri_Sequence()
                    pri_sequence.set_ordinal_pos(priSeqPass)
                    priSeqPass += 1
                    
                elif line.strip() == constant.PRI_SEGMENT:
                    currentEntity = constant.PRI_SEGMENT
                    if priSegmentPass > 0:
                        pri_sequence.add_segment(pri_segment)
                        
                    pri_segment = Pri_Segment()
                    priSegmentPass += 1
    
                elif line.strip() == constant.FREQ_SEQUENCE:
                    currentEntity = constant.FREQ_SEQUENCE
                    if freqSeqPass > 0:
                        if freqSegmentPass > 0:
                            freq_sequence.add_segment(freq_segment)
    
                        generator.add_freq_sequence(freq_sequence)
                        freqSegmentPass = 0
                        
                    freq_sequence = Freq_Sequence()
                    freq_sequence.set_ordinal_pos(freqSeqPass)
                    freqSeqPass += 1
                    
                elif line.strip() == constant.FREQ_SEGMENT:
                    currentEntity = constant.FREQ_SEGMENT
                    if freqSegmentPass > 0:
                        freq_sequence.add_segment(freq_segment)
                        
                    freq_segment = Freq_Segment()
                    freqSegmentPass += 1
                    
                else:
                    if line.strip().__contains__(constant.VALUE_SEPARATOR):
                        line_kv = self.lineAsKeyValue(line.strip())
                        line_key = line_kv[0].strip()
                        line_value = line_kv[1].strip()
                        
                        if currentEntity == constant.EMITTER:
     
                            if line_key.strip() == constant.EMITTER_ELNOT:
                                emitter.set_elnot(line_value)
                                if isBase == True:
                                    emitter.set_bfile(self.bfDisplay)
                                else:
                                    emitter.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                emitter.add_attribute(attrib)
                        
                        elif currentEntity == constant.EMITTER_MODE:
                       
                            if line_key.strip() == constant.MODE_NAME:
                                emitter_mode.set_mode_name(line_value)
                                if isBase == True:
                                    emitter_mode.set_bfile(self.bfDisplay)
                                else:
                                    emitter_mode.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                emitter_mode.add_attribute(attrib)
                                
                        elif currentEntity == constant.GENERATOR:
    
                            if line_key.strip() == constant.GENERATOR_NUMBER:
                                generator.set_generator_number(line_value)
                                if isBase == True:
                                    generator.set_bfile(self.bfDisplay)
                                else:
                                    generator.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                generator.add_attribute(attrib)
                                
                        elif currentEntity == constant.PRI_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                pri_sequence.set_number_of_segments(line_value)
                                if isBase == True:
                                    pri_sequence.set_bfile(self.bfDisplay)
                                else:
                                    pri_sequence.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                pri_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.PRI_SEGMENT:
                            
                            if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                                pri_segment.set_segment_number(line_value)
                                if isBase == True:
                                    pri_segment.set_bfile(self.bfDisplay)
                                else:
                                    pri_segment.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                pri_segment.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                freq_sequence.set_number_of_segments(line_value)
                                if isBase == True:
                                    freq_sequence.set_bfile(self.bfDisplay)
                                else:
                                    freq_sequence.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                freq_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEGMENT:
                            
                            if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                                freq_segment.set_segment_number(line_value)
                                if isBase == True:
                                    freq_segment.set_bfile(self.bfDisplay)
                                else:
                                    freq_segment.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                freq_segment.add_attribute(attrib)
                                
            else:
                emitter_collection.append(emitter)
    
    def parseBFile(self, fName, emitter_collection):
        emitter = Emitter()
        emitter_mode = EmitterMode()
        generator = Generator()
        pri_sequence = Pri_Sequence()
        freq_sequence = Freq_Sequence()
        pri_segment = Pri_Segment()
        freq_segment = Freq_Segment()
            
        passNumber = 0
        modePass = 0
        generatorPass = 0
        priSeqPass = 0
        freqSeqPass = 0
        priSegmentPass = 0
        freqSegmentPass = 0
        
        with open(fName) as f1:
            for cnt, line in enumerate(f1):
                if line.strip() == constant.EMITTER:
                    currentEntity = constant.EMITTER
                    if passNumber > 0:
                        if modePass > 0:
                            emitter.add_mode(emitter_mode)
                            
                        emitter_collection.append(emitter)
                        modePass = 0
                        
                    emitter = Emitter()
                    passNumber += 1
                    
                elif line.strip() == constant.EMITTER_MODE:
                    currentEntity = constant.EMITTER_MODE
                    if modePass > 0:
                        if generatorPass > 0:
                            emitter_mode.add_generator(generator)
                            
                        emitter.add_mode(emitter_mode)
                        generatorPass = 0
                        
                    emitter_mode = EmitterMode()
                    modePass += 1
                    
                elif line.strip() == constant.GENERATOR:
                    currentEntity = constant.GENERATOR
                    if generatorPass > 0:
                        if priSeqPass > 0:
                            generator.add_pri_sequence(pri_sequence)
    
                        if freqSeqPass > 0:
                            generator.add_freq_sequence(freq_sequence)
    
                        emitter_mode.add_generator(generator)
                        priSeqPass = 0
                        freqSeqPass = 0
                        
                    generator = Generator()
                    generatorPass += 1
                    
                elif line.strip() == constant.PRI_SEQUENCE:
                    currentEntity = constant.PRI_SEQUENCE
                    if priSeqPass > 0:
                        if priSegmentPass > 0:
                            pri_sequence.add_segment(pri_segment)
                        
                        generator.add_pri_sequence(pri_sequence)
                        priSegmentPass = 0
                        
                    pri_sequence = Pri_Sequence()
                    pri_sequence.set_ordinal_pos(priSeqPass)
                    priSeqPass += 1
                    
                elif line.strip() == constant.PRI_SEGMENT:
                    currentEntity = constant.PRI_SEGMENT
                    if priSegmentPass > 0:
                        pri_sequence.add_segment(pri_segment)
                        
                    pri_segment = Pri_Segment()
                    priSegmentPass += 1
    
                elif line.strip() == constant.FREQ_SEQUENCE:
                    currentEntity = constant.FREQ_SEQUENCE
                    if freqSeqPass > 0:
                        if freqSegmentPass > 0:
                            freq_sequence.add_segment(freq_segment)
    
                        generator.add_freq_sequence(freq_sequence)
                        freqSegmentPass = 0
                        
                    freq_sequence = Freq_Sequence()
                    freq_sequence.set_ordinal_pos(freqSeqPass)
                    freqSeqPass += 1
                    
                elif line.strip() == constant.FREQ_SEGMENT:
                    currentEntity = constant.FREQ_SEGMENT
                    if freqSegmentPass > 0:
                        freq_sequence.add_segment(freq_segment)
                        
                    freq_segment = Freq_Segment()
                    freqSegmentPass += 1
                    
                else:
                    if line.strip().__contains__(constant.VALUE_SEPARATOR):
                        line_kv = self.lineAsKeyValue(line.strip())
                        line_key = line_kv[0].strip()
                        line_value = line_kv[1].strip()
                        
                        if currentEntity == constant.EMITTER:
     
                            if line_key.strip() == constant.EMITTER_ELNOT:
                                emitter.set_elnot(line_value)
                                emitter.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                emitter.add_attribute(attrib)
                        
                        elif currentEntity == constant.EMITTER_MODE:
                       
                            if line_key.strip() == constant.MODE_NAME:
                                emitter_mode.set_mode_name(line_value)
                                emitter_mode.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                emitter_mode.add_attribute(attrib)
                                
                        elif currentEntity == constant.GENERATOR:
    
                            if line_key.strip() == constant.GENERATOR_NUMBER:
                                generator.set_generator_number(line_value)
                                generator.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                generator.add_attribute(attrib)
                                
                        elif currentEntity == constant.PRI_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                pri_sequence.set_number_of_segments(line_value)
                                pri_sequence.set_bfile(self.bfDisplay)
                             else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                pri_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.PRI_SEGMENT:
                            
                            if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                                pri_segment.set_segment_number(line_value)
                                pri_segment.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                pri_segment.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                freq_sequence.set_number_of_segments(line_value)
                                freq_sequence.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                freq_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEGMENT:
                            
                            if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                                freq_segment.set_segment_number(line_value)
                                freq_segment.set_bfile(self.bfDisplay)
                             else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value)
                                freq_segment.add_attribute(attrib)
                                
            else:
                emitter_collection.append(emitter)


    def parseCFile(self, fName, emitter_collection):
        emitter = Emitter()
        emitter_mode = EmitterMode()
        generator = Generator()
        pri_sequence = Pri_Sequence()
        freq_sequence = Freq_Sequence()
        pri_segment = Pri_Segment()
        freq_segment = Freq_Segment()
            
        passNumber = 0
        modePass = 0
        generatorPass = 0
        priSeqPass = 0
        freqSeqPass = 0
        priSegmentPass = 0
        freqSegmentPass = 0
        
        with open(fName) as f1:
            for cnt, line in enumerate(f1):
                if line.strip() == constant.EMITTER:
                    currentEntity = constant.EMITTER
                    if passNumber > 0:
                        if modePass > 0:
                            emitter.add_mode(emitter_mode)
                           
                        baseEmitter = self.findElnot(emitter.get_elnot(), emitter_collection) 
                        if baseEmitter:
                            baseEmitter.set_cfile(self.cfDisplay)

                            baseEmitter.sync_attributes(emitter)
                            baseEmitter.sync_modes(emitter)
                            
                            
                            #HERE
                            
                            for cMode in enumerate(emitter.get_modes()):
                                baseMode = self.findEmitterMode(cMode.get_name(), baseEmitter)
                                if baseMode:
                                    baseMode.set_cfile(self.cfDisplay)
                                    
                                    for cAttribute in enumerate(cMode.get_attributes()):
                                        bAttribute = self.findAttribute(cAttribute.get_name(), baseMode)
                                         
                                        if bAttribute:
                                            bAttribute.set_cfile(self.cfDisplay)
                                            bAttribute.set_cvalue(cAttribute.get_cvalue())
                                        else:
                                            baseMode.add_attribute(cAttribute)
                                    
                                    for cGenerator in enumerate(cMode.get_generators()):
                                        baseGenerator = self.findGenerator(cGenerator.get_generator_number(), baseMode)
                                                
                                        if baseGenerator:
                                            baseGenerator.set_cfile(self.cfDisplay)

                                            for cAttribute in enumerate(cGenerator.get_attributes()):
                                                bAttribute = self.findAttribute(cAttribute.get_name(), baseGenerator)
                                                
                                                if bAttribute:
                                                    bAttribute.set_cfile(self.cfDisplay)
                                                    bAttribute.set_cvalue(cAttribute.get_cvalue())
                                                else:
                                                    baseGenerator.add_attribute(cAttribute)
                                                    
                                            for cPRISequence in enumerate(cGenerator.get_pri_sequences()):
                                                bPRISequence = self.findPRISequenceByOrdinalPos(cPRISequence.get_ordinal_pos(), baseGenerator)
                                                
                                                if bPRISequence:
                                                    bPRISequence.set_cfile(self.cfDisplay)
                                                    
                                                    for cSegment in enumerate(cPRISequence.get_segments()):
                                                        bSegment = self.findSegmentBySegmentNumber(cSegment.get_segment_number(), bPRISequence)
                                                        
                                                        if bSegment:
                                                            bSegment.set_cfile(self.cfDisplay)
                                                            bSegment.set_cvalue(cSegment.get_cvalue())
                                                            
                                                            for cAttribute in enumerate(cSegment.get_attributes()):
                                                                bAttribute = self.findAttribute(cAttribute.get_name(), bSegment)
                                                                
                                                                if bAttribute:
                                                                    bAttribute.set_cfile(self.cfDisplay)
                                                                    bAttribute.set_cvalue(cAttribute.get_cvalue())
                                                                else:
                                                                    bSegment.add_attribute(cAttribute)
                                                        else:
                                                            bPRISequence.add_segment(cSegment)
                                                            
                                            for cFREQSequence in enumerate(cGenerator.get_freq_sequences()):
                                                bFREQSequence = self.findFREQSequenceByOrdinalPos(cFREQSequence.get_ordinal_pos(), baseGenerator)
                                                
                                                if bFREQSequence:
                                                    bFREQSequence.set_cfile(self.cfDisplay)
                                                    
                                                    for cSegment in enumerate(cFREQSequence.get_segments()):
                                                        bSegment = self.findSegmentBySegmentNumber(cSegment.get_segment_number(), bFREQSequence)
                                                        
                                                        if bSegment:
                                                            bSegment.set_cfile(self.cfDisplay)
                                                            bSegment.set_cvalue(cSegment.get_cvalue())
                                                            
                                                            for cAttribute in enumerate(cSegment.get_attributes()):
                                                                bAttribute = self.findAttribute(cAttribute.get_name(), bSegment)
                                                                
                                                                if bAttribute:
                                                                    bAttribute.set_cfile(self.cfDisplay)
                                                                    bAttribute.set_cvalue(cAttribute.get_cvalue())
                                                                else:
                                                                    bSegment.add_attribute(cAttribute)
                                                        else:
                                                            bFREQSequence.add_segment(cSegment)
                                else:
                                    baseEmitter.add_mode(cMode)
                        else:
                            emitter_collection.append(emitter)
                        modePass = 0
                        
                    emitter = Emitter()
                    passNumber += 1
                    
                elif line.strip() == constant.EMITTER_MODE:
                    currentEntity = constant.EMITTER_MODE
                    if modePass > 0:
                        if generatorPass > 0:
                            emitter_mode.add_generator(generator)
                        
                        emitter.add_mode(emitter_mode)
                        generatorPass = 0
                        
                    emitter_mode = EmitterMode()
                    modePass += 1
                    
                elif line.strip() == constant.GENERATOR:
                    currentEntity = constant.GENERATOR
                    if generatorPass > 0:
                        if priSeqPass > 0:
                            generator.add_pri_sequence(pri_sequence)
    
                        if freqSeqPass > 0:
                            generator.add_freq_sequence(freq_sequence)
    
                        emitter_mode.add_generator(generator)
                        priSeqPass = 0
                        freqSeqPass = 0
                        
                    generator = Generator()
                    generatorPass += 1
                    
                elif line.strip() == constant.PRI_SEQUENCE:
                    currentEntity = constant.PRI_SEQUENCE
                    if priSeqPass > 0:
                        if priSegmentPass > 0:
                            pri_sequence.add_segment(pri_segment)
                        
                        generator.add_pri_sequence(pri_sequence)
                        priSegmentPass = 0
                        
                    pri_sequence = Pri_Sequence()
                    pri_sequence.set_ordinal_pos(priSeqPass)
                    priSeqPass += 1
                    
                elif line.strip() == constant.PRI_SEGMENT:
                    currentEntity = constant.PRI_SEGMENT
                    if priSegmentPass > 0:
                        pri_sequence.add_segment(pri_segment)
                        
                    pri_segment = Pri_Segment()
                    priSegmentPass += 1
    
                elif line.strip() == constant.FREQ_SEQUENCE:
                    currentEntity = constant.FREQ_SEQUENCE
                    if freqSeqPass > 0:
                        if freqSegmentPass > 0:
                            freq_sequence.add_segment(freq_segment)
    
                        generator.add_freq_sequence(freq_sequence)
                        freqSegmentPass = 0
                        
                    freq_sequence = Freq_Sequence()
                    freq_sequence.set_ordinal_pos(freqSeqPass)
                    freqSeqPass += 1
                    
                elif line.strip() == constant.FREQ_SEGMENT:
                    currentEntity = constant.FREQ_SEGMENT
                    if freqSegmentPass > 0:
                        freq_sequence.add_segment(freq_segment)
                        
                    freq_segment = Freq_Segment()
                    freqSegmentPass += 1
                    
                else:
                    if line.strip().__contains__(constant.VALUE_SEPARATOR):
                        line_kv = self.lineAsKeyValue(line.strip())
                        line_key = line_kv[0].strip()
                        line_value = line_ kv[1].strip()
                        
                        if currentEntity == constant.EMITTER:
     
                            if line_key.strip() == constant.EMITTER_ELNOT:
                                emitter.set_elnot(line_value)
                                emitter.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                emitter.add_attribute(attrib)
                        
                        elif currentEntity == constant.EMITTER_MODE:
                       
                            if line_key.strip() == constant.MODE_NAME:
                                emitter_mode.set_mode_name(line_value)
                                emitter_mode.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                emitter_mode.add_attribute(attrib)
                                
                        elif currentEntity == constant.GENERATOR:
    
                            if line_key.strip() == constant.GENERATOR_NUMBER:
                                generator.set_generator_number(line_value)
                                generator.set_bfile(self.bfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                generator.add_attribute(attrib)
                                
                        elif currentEntity == constant.PRI_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                pri_sequence.set_number_of_segments(line_value)
                                pri_sequence.set_cfile(self.cfDisplay)
                             else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                pri_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.PRI_SEGMENT:
                            
                            if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                                pri_segment.set_segment_number(line_value)
                                pri_segment.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                pri_segment.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                freq_sequence.set_number_of_segments(line_value)
                                freq_sequence.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                freq_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEGMENT:
                            
                            if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                                freq_segment.set_segment_number(line_value)
                                freq_segment.set_cfile(self.cfDisplay)
                             else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_cvalue(line_value)
                                freq_segment.add_attribute(attrib)
                                
            else:
                emitter_collection.append(emitter)

    
    def parseBaseFile(self):
        print("baseFileName: {}".format(self.baseFileName))
        self.parseFile(self.baseFileName, self.base_emitters, True)
            
        
    def parseComparisonFile(self):
        self.parseFile(self.comparisonFileName, self.base_emitters, False)
        
              
    def findElnot(self, elnotValue, inArray):
        if inArray == constant.COMPARISON_ARRAY:
            for emitter in self.comparison_emitters:
                if emitter.get_elnot() == elnotValue:
                    return emitter
            
            return []
        else:
            for emitter in self.base_emitters:
                if emitter.get_elnot() == elnotValue:
                    return emitter
                
            return []
    
    
    def findEmitterMode(self, emitterModeName, comparisonEmitter):
            for comparisonEmitterMode in comparisonEmitter.get_modes():
                if comparisonEmitterMode.get_name() == emitterModeName:
                    return comparisonEmitterMode
              
            return []
        
    
    def findGenerator(self, generatorNumber, comparisonEmitterMode):
            for comparisonGenerator in comparisonEmitterMode.get_generators():
                if comparisonGenerator.get_generator_number() == generatorNumber:
                    return comparisonGenerator
              
            return []
    
    
    def findPRISequenceByOrdinalPos(self, ordinalPos, comparisonGenerator):
        for cPRISequence in comparisonGenerator.get_pri_sequences():
            if cPRISequence.get_ordinal_pos() == ordinalPos:
                return cPRISequence
         
        return []
    
    
    def findFREQSequenceByOrdinalPos(self, ordinalPos, comparisonGenerator):
        for cFREQSequence in comparisonGenerator.get_freq_sequences():
            if cFREQSequence.get_ordinal_pos() == ordinalPos:
                return cFREQSequence 
            
        return []
        
    
    def findSegmentBySegmentNumber(self, segmentNumber, comparisonSequence):
        for cSegment in comparisonSequence.get_segments():
            if cSegment.get_segment_number() == segmentNumber:
                return cSegment
            
        return []
    
        
    def findAttribute(self, baseAttributeName, comparisonObject):
        for comparisonAttribute in comparisonObject.get_attributes():
            if comparisonAttribute.get_name() == baseAttributeName:
                return comparisonAttribute
            
        return []
    
            
    def writeCell(self, ws, wsRow, cellValue):
        ws.range(wsRow, 1).value = cellValue
        ws.range(wsRow, 1).api.VerticalAlignment = constants.VAlign.xlVAlignTop
        #ws.range(wsRow, 1).WrapText = True

        
    def writeSpecificCell(self, ws, wsRow, wsCol, cellValue):
        ws.range(wsRow, wsCol).value = cellValue
        ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignCenter
        ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        ws.range(wsRow, wsCol).WrapText = True

    def writeValueCell(self, ws, wsRow, wsCol, cellValue):
        ws.range(wsRow, wsCol).value = cellValue
        try:
            ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignLeft
            ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
            ws.range(wsRow, wsCol).WrapText = True
        except Exception:
            print("exception was thrown")
            

    def writeLabelCell(self, ws, wsRow, wsCol, cellValue):
        ws.range(wsRow, wsCol).value = cellValue
        try:
            ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignRight
            ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
            ws.range(wsRow, wsCol).WrapText = True
            if cellValue == constant.XL_MISSING_TEXT:
                ws.range(wsRow, wsCol).api.Font.Color = rgb_to_int((255,0,0))

        except Exception:
            print("exception was thrown")

    
    
    
    def getShortFileNames(self):
        bfIdx = self.baseFileName.rfind('/')
        cfIdx = self.comparisonFileName.rfind('/')
        
        if bfIdx > -1:
            bfArray = self.baseFileName.rsplit('/', 1)
            self.bfDisplay = bfArray[1]
        else:
            self.bfDisplay = self.baseFileName
            
        if cfIdx > -1:
            cfArray = self.comparisonFileName.rsplit('/', 1)
            self.cfDisplay = cfArray[1]
        else:
            self.cfDisplay = self.comparisonFileName
            
        
    
    def writeTitleCells(self, ws):
        
        bfTitle = "Base File: {}".format(self.bfDisplay)
        cfTitle = "Comparison File: {}".format(self.cfDisplay)
        
        ws.range((1,1), (1,12)).merge(across=True)
        self.writeSpecificCell(ws, 1, 1, bfTitle)
        
        ws.range((1,13), (1,25)).merge(across=True)
        self.writeSpecificCell(ws, 1, 13, cfTitle)



    def setColorMarkBoundary(self, wsBkgColorMarks, wsRows):
        if wsRows > 2:
            wsBkgColorMarks[-1].set_bottomboundary(wsRows)
            wsRows += 1
            wsColorMarks = ColorMarkRowBoundary()
            wsColorMarks.set_topboundary(wsRows)
            wsBkgColorMarks.append(wsColorMarks)
        else:
            wsColorMarks = ColorMarkRowBoundary()
            wsColorMarks.set_topboundary(wsRows)
            wsBkgColorMarks.append(wsColorMarks)
            
            
    def writeMissingComparisonEmitter(self, ws, wsRow, bElnot):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
        
    
    def writeMissingComparisonEmitterAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullEmitterAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, cAttributeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, cAttributeValue)
        
        
    def writeMissingComparisonMode(self, ws, wsRow, bEmitterModeName):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_VAL, bEmitterModeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
        
         
    def writeFullEmitterMode(self, ws, wsRow, bEmitterModeName, cEmitterModeName):
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_VAL, bEmitterModeName)
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_VAL, cEmitterModeName)

        
    def writeMissingComparisonModeAttribute(self, ws, wsRow, bModeAttributeName, bModeAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, bModeAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, bModeAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullModeAttribute(self, ws, wsRow, bModeAttributeName, bModeAttributeValue, cModeAttributeName, cModeAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, bModeAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, bModeAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, cModeAttributeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, cModeAttributeValue)
        

    def writeMissingComparisonGenerator(self, ws, wsRow, bGeneratorNumber):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_VAL, bGeneratorNumber)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)


    def writeMissingComparisonGeneratorAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullGeneratorAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, cAttributeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, cAttributeValue)

        
    def writeFullPRISequenceCounts(self, ws, wsRow, bgPRISequenceCount, cpPRISequenceCount):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, bgPRISequenceCount)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, cpPRISequenceCount)
        

    def writeFullFREQSequenceCounts(self,  ws, wsRow, bgFREQSequenceCount, cpFREQSequenceCount):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, bgFREQSequenceCount)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, cpFREQSequenceCount)

      
    def writeMissingComparisonPRISequence(self, ws, wsRow, sequencePos):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        

    def writeMissingComparisonSequenceAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeFullSequenceAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cAttributeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cAttributeValue)
                                                   
    
    def writeMissingComparisonSegment(self, ws, wsRow, bSegmentNumber):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bSegmentNumber)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeMissingComparisonSegmentAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
                 
    def writeFullSegmentAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bAttributeName)
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cAttributeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cAttributeValue)
        

    def writeMissingComparisonFREQSequence(self, ws, wsRow, sequencePos):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseEmitter(self, ws, wsRow, cElnot):
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
        

    def writeMissingBaseEmitterAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, cAttributeName)
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, cAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseMode(self, ws, wsRow, cModeName):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_VAL, cModeName)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
        
            
    def writeMissingBaseModeAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, cAttributeName)
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, cAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
                   
    def writeMissingBaseGenerator(self, ws, wsRow, cGeneratorNumber):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_VAL, cGeneratorNumber)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)


    def writeMissingBaseGeneratorAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, cAttributeName)
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, cAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)


    def writeMissingBasePRISequence(self, ws, wsRow, sequencePos):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseSequenceAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cAttributeName)
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBasePRISegment(self, ws, wsRow, cPRISegmentNumber):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cPRISegmentNumber)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
                                                
        
    def writeMissingBaseSegmentAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cAttributeName)
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cAttributeValue)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeMissingBaseFREQSequence(self, ws, wsRow, sequencePos):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)


    def writeMissingBaseFREQSegment(self, ws, wsRow, cFREQSegmentNumber):
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        self.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cFREQSegmentNumber)
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeEmitter(self, wsEmitters, wsEmittersRow, bElnot, cElnot, emitterWritten):
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
        emitterWritten = True
        

    def writeMode(self, wsModes, wsModesRow, bElnot, cElnot, baseEmitterModeName, comparisonEmitterModeName, modeWritten):
        self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)

        self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterModeName)
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterModeName)
        modeWritten = True
        
    
    def writeGenerator(self, wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten):
        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_VAL, comparisonEmitter.get_elnot())

        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterMode.get_name())
        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterMode.get_name())

        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_VAL, baseGenerator.get_generator_number())
        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_VAL, comparisonGenerator.get_generator_number())
        generatorWritten = True
        
        
    def writePRISequence(self, wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten):
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_VAL, comparisonEmitter.get_elnot())

        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterMode.get_name())
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterMode.get_name())

        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, baseGenerator.get_generator_number())
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, comparisonGenerator.get_generator_number())
        
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, basePRISequence.get_ordinal_pos())
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, comparisonPRISequence.get_ordinal_pos())
        priSequenceWritten = True
        
    
    def writePRISegment(self, wsPRISequences, wsPRISequencesRow, bPRISegment, cPRISegment, priSegmentWritten):
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bPRISegment.get_segment_number())
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cPRISegment.get_segment_number())
        priSegmentWritten = True
        
    
    def writeFREQSequence(self, wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten):
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_VAL, comparisonEmitter.get_elnot())

        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterMode.get_name())
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterMode.get_name())

        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, baseGenerator.get_generator_number())
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, comparisonGenerator.get_generator_number())
        
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, baseFREQSequence.get_ordinal_pos())
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, comparisonFREQSequence.get_ordinal_pos())
        freqSequenceWritten = True
        

    def writeFREQSegment(self, wsFREQSequences, wsFREQSequencesRow, bFREQSegment, cFREQSegment, freqSegmentWritten):
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bFREQSegment.get_segment_number())
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cFREQSegment.get_segment_number())
        freqSegmentWritten = True
        
    
    # def buildAllEmittersArray(self):
        
    #     base_emitter_collection = self.base_emitters
    #     comparison_emitter_collection = self.comparison_emitters
    #     all_emitters_collection = self.all_emitters
    #     comparisonArray = constant.COMPARISON_ARRAY
    #     baseArray = constant.BASE_ARRAY
        
    #     for baseEmitter in base_emitter_collection:
    #         bElnot = baseEmitter.get_elnot()
    #         comparisonEmitter = self.findElnot(bElnot, comparisonArray)
            
    #         if not comparisonEmitter:
    #            baseEmitter.set_cfile(constant.XL_MISSING_TEXT)
    #         else:
    #             baseEmitter.set_cfile(self.cfDisplay)
    #             for baseAttribute in baseEmitter.get_attributes():
    #                 comparisonAttribute = self.findAttribute(baseAttribute.get_name(), comparisonEmitter)
                    
    #                 if not comparisonAttribute:
    #                     baseAttribute.set_cvalue(constant.XL_MISSING_TEXT)
    #                 else:
    #                     baseAttribute.set_cvalue(comparisonAttribute.get_value())
                            
    #             for baseEmitterMode in baseEmitter.get_modes():
    #                 comparisonEmitterMode = self.findEmitterMode(baseEmitterMode.get_name(), comparisonEmitter)
                    
                    
    #                 if not comparisonEmitterMode:
    #                     baseEmitterMode.set_cfile(constant.XL_MISSING_TEXT)
    #                 else:
    #                     baseEmitterMode.set_cfile()
    #                     if comparisonEmitterMode.get_name() != baseEmitterMode.get_name():
                        
                        
    #                     for baseModeAttribute in baseEmitterMode.get_attributes():
    #                         comparisonModeAttribute = self.findAttribute(baseModeAttribute.get_name(), comparisonEmitterMode)
                            
    #                         if not comparisonModeAttribute:
                                
    #                         else:
    #                             if comparisonModeAttribute.get_value() != baseModeAttribute.get_value():
                                    
                                    
    #                     for baseGenerator in baseEmitterMode.get_generators():
    #                         comparisonGenerator = self.findGenerator(baseGenerator.get_generator_number(), comparisonEmitterMode)
                        
                            
    #                         if not comparisonGenerator:
                                
    #                         else:
                                
    #                             for baseGeneratorAttribute in baseGenerator.get_attributes():
    #                                 comparisonGeneratorAttribute = self.findAttribute(baseGeneratorAttribute.get_name(), comparisonGenerator)
                                    
    #                                 if not comparisonGeneratorAttribute:
                                        
                                        
                                        
    #                                 else:
                                        
    #                                     if comparisonGeneratorAttribute.get_value() != baseGeneratorAttribute.get_value():

                            
    #                             bgPRISequences = baseGenerator.get_pri_sequences()
    #                             bgFREQSequences = baseGenerator.get_freq_sequences()
    #                             cpPRISequences = comparisonGenerator.get_pri_sequences()
    #                             cpFREQSequences = comparisonGenerator.get_freq_sequences()
                                
    #                             if len(bgPRISequences) != len(cpPRISequences):

        
    #                             if len(bgFREQSequences) != len(cpFREQSequences):
                                    
                                
    #                             for basePRISequence in baseGenerator.get_pri_sequences():
    #                                 comparisonPRISequence = self.findPRISequenceByOrdinalPos(basePRISequence.get_ordinal_pos(), comparisonGenerator)
                                    
    #                                 if not comparisonPRISequence:
                                        
    #                                 else:
                                        
    #                                     for bPRISeqAttribute in basePRISequence.get_attributes(): 
    #                                         cPRISeqAttribute = self.findAttribute(bPRISeqAttribute.get_name(), comparisonPRISequence)
                                            
    #                                         if not cPRISeqAttribute:
                                                

    #                                         else:

    #                                             if cPRISeqAttribute.get_value () != bPRISeqAttribute.get_value():
                                                    
        
    #                                     for bPRISegment in basePRISequence.get_segments():
    #                                         cPRISegment = self.findSegmentBySegmentNumber(bPRISegment.get_segment_number(), comparisonPRISequence)
                                            
    #                                         if not cPRISegment:
                                                
    #                                         else:
                                                
    #                                             for bSegmentAttribute in bPRISegment.get_attributes():
    #                                                 cSegmentAttribute = self.findAttribute(bSegmentAttribute.get_name(), cPRISegment)
                                                    
    #                                                 if not cSegmentAttribute:
                                                        

    #                                                 else:

    #                                                     if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
                    
    #                             for baseFREQSequence in baseGenerator.get_freq_sequences():
    #                                 comparisonFREQSequence = self.findFREQSequenceByOrdinalPos(baseFREQSequence.get_ordinal_pos(), comparisonGenerator)
    #                                 freqSequenceWritten = False
                                    
    #                                 if not comparisonFREQSequence:
                                        

    #                                 else:
                                        
    #                                     for bFREQSeqAttribute in baseFREQSequence.get_attributes():
    #                                         cFREQSeqAttribute = self.findAttribute(bFREQSeqAttribute.get_name(), comparisonFREQSequence)
                                            
    #                                         if not cFREQSeqAttribute:
                                                

    #                                         else:
    #                                             if cFREQSeqAttribute.get_value() != bFREQSeqAttribute.get_value():
                                                    
        
    #                                     for bFREQSegment in baseFREQSequence.get_segments():
    #                                         cFREQSegment = self.findSegmentBySegmentNumber(bFREQSegment.get_segment_number(), comparisonFREQSequence)
    #                                         freqSegmentWritten = False
                                            
    #                                         if not cFREQSegment:

                                                
    #                                         else:
                                                
    #                                             for bSegmentAttribute in bFREQSegment.get_attributes():
    #                                                 cSegmentAttribute = self.findAttribute(bSegmentAttribute.get_name(), cFREQSegment)
                                                    
    #                                                 if not cSegmentAttribute:
                                                        

    #                                                 else:

    #                                                     if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
        
        
        

    def compareEMTFiles(self, wb, base_emitter_collection, comparison_emitter_collection, comparisonArray, baseArray):
         
        wsCount = wb.sheets.count
        
        if wsCount > 0:
            wb.sheets[0].name = "Emitters"
        else:
            wb.sheets.add('Emitters')
            
        if wsCount > 1:
            wb.sheets[1].name = "Modes"
        else:
            wb.sheets.add('Modes', after='Emitters')
            
        if wsCount > 2:
            wb.sheets[2].name = "Generators"
        else:
            wb.sheets.add('Generators', after='Modes')
        
        wsEmitters = wb.sheets[0]
        wsModes = wb.sheets[1]
        wsGenerators = wb.sheets[2]
        
        wsPRISequences = wb.sheets.add('PRISequences', after='Generators')
        wsFREQSequences = wb.sheets.add('FREQSequences', after='PRISequences')
        
        
        self.writeTitleCells(wsEmitters)
        self.writeTitleCells(wsModes)
        self.writeTitleCells(wsGenerators)
        self.writeTitleCells(wsPRISequences)
        self.writeTitleCells(wsFREQSequences)
       
        wsEmittersRow = 2
        wsModesRow = 2
        wsGeneratorsRow = 2
        wsPRISequencesRow = 2
        wsFREQSequencesRow = 2
        cellValue = ""
        
        wsEmittersBkgColorMarks = []
        wsModesBkgColorMarks = []
        wsGeneratorsBkgColorMarks = []
        wsPRISequencesBkgColorMarks = []
        wsFREQSequencesBkgColorMarks = []
        
        
        for baseEmitter in base_emitter_collection:
            bElnot = baseEmitter.get_elnot()
            comparisonEmitter = self.findElnot(bElnot, comparisonArray)
            
            emitterWritten = False
            
            self.setColorMarkBoundary(wsEmittersBkgColorMarks, wsEmittersRow)
            self.setColorMarkBoundary(wsModesBkgColorMarks, wsModesRow)
            self.setColorMarkBoundary(wsGeneratorsBkgColorMarks, wsGeneratorsRow)
            self.setColorMarkBoundary(wsPRISequencesBkgColorMarks, wsPRISequencesRow)
            self.setColorMarkBoundary(wsFREQSequencesBkgColorMarks, wsFREQSequencesRow)
            
            if not comparisonEmitter:
                self.writeMissingComparisonEmitter(wsEmitters, wsEmittersRow, bElnot)
                wsEmittersRow += 1
               
            else:

                for baseAttribute in baseEmitter.get_attributes():
                    comparisonAttribute = self.findAttribute(baseAttribute.get_name(), comparisonEmitter)
                    
                    if not comparisonAttribute:
                        if emitterWritten == False:
                            self.writeEmitter(wsEmitters, wsEmittersRow, bElnot, comparisonEmitter.get_elnot(), emitterWritten)
                            wsEmittersRow += 1
                            
                        self.writeMissingComparisonEmitterAttribute(wsEmitters, wsEmittersRow, baseAttribute.get_name(), baseAttribute.get_value())
                        wsEmittersRow += 1
                        
                    else:
                        if comparisonAttribute.get_value() != baseAttribute.get_value():

                            if emitterWritten == False:
                                self.writeEmitter(wsEmitters, wsEmittersRow, bElnot, comparisonEmitter.get_elnot(), emitterWritten)
                                wsEmittersRow += 1

                            self.writeFullEmitterAttribute(wsEmitters, wsEmittersRow, baseAttribute.get_name(), baseAttribute.get_value(), comparisonAttribute.get_name(), comparisonAttribute.get_value())
                            wsEmittersRow += 1

                            
                for baseEmitterMode in baseEmitter.get_modes():
                    comparisonEmitterMode = self.findEmitterMode(baseEmitterMode.get_name(), comparisonEmitter)
                    
                    modeWritten = False
                    
                    if not comparisonEmitterMode:
                        self.writeMissingComparisonMode(wsEmitters, wsEmittersRow, baseEmitterMode.get_name())
                        wsEmittersRow += 1
                    else:
                        if comparisonEmitterMode.get_name() != baseEmitterMode.get_name():
                            self.writeFullEmitterMode(wsEmitters, wsEmittersRow, baseEmitterMode.get_name(), comparisonEmitterMode.get_name())
                            wsEmittersRow += 1
                        
                        
                        for baseModeAttribute in baseEmitterMode.get_attributes():
                            comparisonModeAttribute = self.findAttribute(baseModeAttribute.get_name(), comparisonEmitterMode)
                            
                            if not comparisonModeAttribute:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                
                                self.writeMissingComparisonModeAttribute(wsModes, wsModesRow, baseModeAttribute.get_name(), baseModeAttribute.get_value())
                                
                                
                                wsModesRow += 1
                            else:
                                if comparisonModeAttribute.get_value() != baseModeAttribute.get_value():
                                    
                                    if modeWritten == False:
                                        self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                        wsModesRow += 1
                                    
                                    self.writeFullModeAttribute(wsModes, wsModesRow, baseModeAttribute.get_name(), baseModeAttribute.get_value(), comparisonModeAttribute.get_name(), comparisonModeAttribute.get_value())
                                    wsModesRow += 1
                                    
                        for baseGenerator in baseEmitterMode.get_generators():
                            comparisonGenerator = self.findGenerator(baseGenerator.get_generator_number(), comparisonEmitterMode)
                        
                            generatorWritten = False
                            
                            if not comparisonGenerator:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1

                                self.writeMissingComparisonGenerator(wsModes, wsModesRow, baseGenerator.get_generator_number())
                                wsModesRow += 1
                            else:
                                
                                for baseGeneratorAttribute in baseGenerator.get_attributes():
                                    comparisonGeneratorAttribute = self.findAttribute(baseGeneratorAttribute.get_name(), comparisonGenerator)
                                    
                                    if not comparisonGeneratorAttribute:
                                        
                                        
                                        if generatorsWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                                
                                        self.writeMissingComparisonGeneratorAttribute(wsGenerators, wsGeneratorsRow, baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_value())
                                        wsGeneratorsRow += 1
                                        
                                    else:
                                        
                                        if comparisonGeneratorAttribute.get_value() != baseGeneratorAttribute.get_value():

                                            if generatorWritten == False:
                                                self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                                wsGeneratorsRow += 1

                                            self.writeFullGeneratorAttribute(wsGenerators, wsGeneratorsRow, baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_value(), comparisonGeneratorAttribute.get_name(), comparisonGeneratorAttribute.get_value())
                                            wsGeneratorsRow += 1
                            
                                bgPRISequences = baseGenerator.get_pri_sequences()
                                bgFREQSequences = baseGenerator.get_freq_sequences()
                                cpPRISequences = comparisonGenerator.get_pri_sequences()
                                cpFREQSequences = comparisonGenerator.get_freq_sequences()
                                
                                if len(bgPRISequences) != len(cpPRISequences):

                                    if generatorWritten == False:
                                        self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                        wsGeneratorsRow += 1

                                    self.writeFullPRISequenceCounts(wsGenerators, wsGeneratorsRow, len(bgPRISequences), len(cpPRISequences))
                                    wsGeneratorsRow += 1
        
                                if len(bgFREQSequences) != len(cpFREQSequences):
                                    
                                    if generatorWritten == False:
                                        self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                        wsGeneratorsRow += 1
                                    
                                    self.writeFullFREQSequenceCounts(wsGenerators, wsGeneratorsRow, len(bgFREQSequences), len(cpFREQSequences))
                                    wsGeneratorsRow += 1
                                
                                for basePRISequence in baseGenerator.get_pri_sequences():
                                    comparisonPRISequence = self.findPRISequenceByOrdinalPos(basePRISequence.get_ordinal_pos(), comparisonGenerator)
                                    priSequenceWritten = False
                                    
                                    if not comparisonPRISequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeMissingComparisonPRISequence(wsGenerators, wsGeneratorsRow, basePRISequence.get_ordinal_pos())
                                        wsGeneratorsRow += 1
                                        
                                    else:
                                        
                                        for bPRISeqAttribute in basePRISequence.get_attributes(): 
                                            cPRISeqAttribute = self.findAttribute(bPRISeqAttribute.get_name(), comparisonPRISequence)
                                            
                                            if not cPRISeqAttribute:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                    
                                                self.writeMissingComparisonSequenceAttribute(wsPRISequences, wsPRISequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value())
                                                wsPRISequencesRow += 1

                                            else:

                                                if cPRISeqAttribute.get_value () != bPRISeqAttribute.get_value():
                                                    
                                                    if priSequenceWritten == False:
                                                        self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                        wsPRISequencesRow += 1
                                                    
                                                    self.writeFullSequenceAttribute(wsPRISequences, wsPRISequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value(), cPRISeqAttribute.get_name(), cPRISeqAttribute.get_value())
                                                    wsPRISequencesRow += 1
        
                                        for bPRISegment in basePRISequence.get_segments():
                                            cPRISegment = self.findSegmentBySegmentNumber(bPRISegment.get_segment_number(), comparisonPRISequence)
                                            priSegmentWritten = False
                                            
                                            if not cPRISegment:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                
                                                self.writeMissingComparisonSegment(wsPRISequences, wsPRISequencesRow, bPRISegment.get_segment_number())
                                                wsPRISequencesRow += 1
                                                
                                            else:
                                                
                                                for bSegmentAttribute in bPRISegment.get_attributes():
                                                    cSegmentAttribute = self.findAttribute(bSegmentAttribute.get_name(), cPRISegment)
                                                    
                                                    if not cSegmentAttribute:
                                                        
                                                        if priSequenceWritten == False:
                                                            self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                            wsPRISequencesRow += 1
                                                        
                                                        if priSegmentWritten == False: 
                                                            self.writePRISegment(wsPRISequences, wsPRISequencesRow, bPRISegment, cPRISegment, priSegmentWritten)
                                                            wsPRISequencesRow += 1
                                                           
                                                        self.writeMissingComparisonSegmentAttribute(wsPRISequences, wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())
                                                        wsPRISequencesRow += 1

                                                    else:

                                                        if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
                                                            if priSequenceWritten == False:
                                                                self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                                wsPRISequencesRow += 1
                                                            
                                                            if priSegmentWritten == False:
                                                                self.writePRISegment(wsPRISequences, wsPRISequencesRow, bPRISegment, cPRISegment, priSegmentWritten)
                                                                wsPRISequencesRow += 1
                                                                
                                                            self.writeFullSegmentAttribute(wsPRISequences, wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), cSegmentAttribute.get_name(), cSegmentAttribute.get_value())
                                                            wsPRISequencesRow += 1
                    
                                for baseFREQSequence in baseGenerator.get_freq_sequences():
                                    comparisonFREQSequence = self.findFREQSequenceByOrdinalPos(baseFREQSequence.get_ordinal_pos(), comparisonGenerator)
                                    freqSequenceWritten = False
                                    
                                    if not comparisonFREQSequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeMissingComparisonFREQSequence(wsGenerators, wsGeneratorsRow, baseFREQSequence.get_ordinal_pos())
                                        wsGeneratorsRow += 1

                                    else:
                                        
                                        for bFREQSeqAttribute in baseFREQSequence.get_attributes():
                                            cFREQSeqAttribute = self.findAttribute(bFREQSeqAttribute.get_name(), comparisonFREQSequence)
                                            
                                            if not cFREQSeqAttribute:
                                                
                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1
                                                 
                                                self.writeMissingComparisonSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, bFREQSeqAttribute.get_name(), bFREQSeqAttribute.get_value())    
                                                wsFREQSequencesRow += 1

                                            else:
                                                if cFREQSeqAttribute.get_value() != bFREQSeqAttribute.get_value():
                                                    
                                                    if freqSequenceWritten == False:
                                                        self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                        wsFREQSequencesRow += 1
                                                    
                                                    self.writeFullSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, bFREQSeqAttribute.get_name(), bFREQSeqAttribute.get_value(), cFREQSeqAttribute.get_name(), cFREQSeqAttribute.get_value())
                                                    wsFREQSequencesRow += 1
        
                                        for bFREQSegment in baseFREQSequence.get_segments():
                                            cFREQSegment = self.findSegmentBySegmentNumber(bFREQSegment.get_segment_number(), comparisonFREQSequence)
                                            freqSegmentWritten = False
                                            
                                            if not cFREQSegment:

                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1

                                                self.writeMissingComparisonSegment(wsFREQSequences, wsFREQSequencesRow, bFREQSegment.get_segment_number())
                                                wsFREQSequencesRow += 1
                                                
                                            else:
                                                
                                                for bSegmentAttribute in bFREQSegment.get_attributes():
                                                    cSegmentAttribute = self.findAttribute(bSegmentAttribute.get_name(), cFREQSegment)
                                                    
                                                    if not cSegmentAttribute:
                                                        
                                                        if freqSequenceWritten == False:
                                                            self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                            wsFREQSequencesRow += 1
                                                        
                                                        if freqSegmentWritten == False:
                                                            self.writeFREQSegment(wsFREQSequences, wsFREQSequencesRow, bFREQSegment, cFREQSegment, freqSegmentWritten)
                                                            wsFREQSequencesRow += 1
                                                          
                                                        self.writeMissingComparisonSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())    
                                                        wsFREQSequencesRow += 1

                                                    else:

                                                        if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
                                                            if freqSequenceWritten == False:
                                                                self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                                wsFREQSequencesRow += 1
                                                            
                                                            if freqSegmentWritten == False:
                                                                self.writeFREQSegment(wsFREQSequences, wsFREQSequencesRow, bFREQSegment, cFREQSegment, freqSegmentWritten)
                                                                wsFREQSequencesRow += 1
                                                            
                                                            self.writeFullSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), cSegmentAttribute.get_name(), cSegmentAttribute.get_value())
                                                            wsFREQSequencesRow += 1
                                                            
        for comparisonEmitter in comparison_emitter_collection:
            cElnot = comparisonEmitter.get_elnot()
            baseEmitter = self.findElnot(cElnot, baseArray)
            
            emitterWritten = False
            
            self.setColorMarkBoundary(wsEmittersBkgColorMarks, wsEmittersRow)
            self.setColorMarkBoundary(wsModesBkgColorMarks, wsModesRow)
            self.setColorMarkBoundary(wsGeneratorsBkgColorMarks, wsGeneratorsRow)
            self.setColorMarkBoundary(wsPRISequencesBkgColorMarks, wsPRISequencesRow)
            self.setColorMarkBoundary(wsFREQSequencesBkgColorMarks, wsFREQSequencesRow)
            
            if not baseEmitter:
                self.writeMissingBaseEmitter(wsEmitters, wsEmittersRow, cElnot)
                wsEmittersRow += 1
               
            else:

                for comparisonAttribute in comparisonEmitter.get_attributes():
                    baseAttribute = self.findAttribute(comparisonAttribute.get_name(), baseEmitter)
                    
                    if not baseAttribute:
                        if emitterWritten == False:
                            self.writeEmitter(wsEmitters, wsEmittersRow, cElnot, baseEmitter.get_elnot(), emitterWritten)
                            wsEmittersRow += 1
                            
                        self.writeMissingBaseEmitterAttribute(wsEmitters, wsEmittersRow, comparisonAttribute.get_name(), comparisonAttribute.get_value())    
                        wsEmittersRow += 1
                        
                            
                for comparisonEmitterMode in comparisonEmitter.get_modes():
                    baseEmitterMode = self.findEmitterMode(comparisonEmitterMode.get_name(), baseEmitter)
                    
                    modeWritten = False
                    
                    if not baseEmitterMode:
                        
                        self.writeMissingBaseMode(wsEmitters, wsEmittersRow, comparisonEmitterMode.get_name())
                        wsEmittersRow += 1
                    else:
                        
                        for comparisonModeAttribute in comparisonEmitterMode.get_attributes():
                            baseModeAttribute = self.findAttribute(comparisonModeAttribute.get_name(), baseEmitterMode)
                            
                            if not baseModeAttribute:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, cElnot, baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                
                                self.writeMissingBaseModeAttribute(wsModes, wsModesRow, comparisonModeAttribute.get_name(), comparisonModeAttribute.get_value())
                                wsModesRow += 1
                                    
                        for comparisonGenerator in comparisonEmitterMode.get_generators():
                            baseGenerator = self.findGenerator(comparisonGenerator.get_generator_number(), baseEmitterMode)
                        
                            generatorWritten = False
                            
                            if not baseGenerator:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                
                                self.writeMissingBaseGenerator(wsModes, wsModesRow, comparisonGenerator.get_generator_number())
                                wsModesRow += 1
                            else:
                                
                                for comparisonGeneratorAttribute in comparisonGenerator.get_attributes():
                                    baseGeneratorAttribute = self.findAttribute(comparisonGeneratorAttribute.get_name(), baseGenerator)
                                    
                                    if not baseGeneratorAttribute:
                                        
                                        if generatorsWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                                
                                        self.writeMissingBaseGeneratorAttribute(wsGenerators, wsGeneratorsRow, comparisonGeneratorAttribute.get_name(), comparisonGeneratorAttribute.get_value())
                                        wsGeneratorsRow += 1
                                        
                                bgPRISequences = baseGenerator.get_pri_sequences()
                                bgFREQSequences = baseGenerator.get_freq_sequences()
                                cpPRISequences = comparisonGenerator.get_pri_sequences()
                                cpFREQSequences = comparisonGenerator.get_freq_sequences()
                                
                                for comparisonPRISequence in comparisonGenerator.get_pri_sequences():
                                    basePRISequence = self.findPRISequenceByOrdinalPos(comparisonPRISequence.get_ordinal_pos(), baseGenerator)
                                    priSequenceWritten = False
                                    
                                    if not basePRISequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeMissingBasePRISequence(wsGenerators, wsGeneratorsRow, comparisonPRISequence.get_ordinal_pos())
                                        wsGeneratorsRow += 1
                                    else:
                                        
                                        for cPRISeqAttribute in comparisonPRISequence.get_attributes(): 
                                            bPRISeqAttribute = self.findAttribute(cPRISeqAttribute.get_name(), basePRISequence)
                                            
                                            if not bPRISeqAttribute:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                
                                                self.writeMissingBaseSequenceAttribute(wsPRISequences, wsPRISequencesRow, cPRISeqAttribute.get_name(), cPRISeqAttribute.get_value())
                                                wsPRISequencesRow += 1

                                        for cPRISegment in comparisonPRISequence.get_segments():
                                            bPRISegment = self.findSegmentBySegmentNumber(cPRISegment.get_segment_number(), basePRISequence)
                                            priSegmentWritten = False
                                            
                                            if not bPRISegment:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                
                                                self.writeMissingBasePRISegment(wsPRISequences, wsPRISequencesRow, cPRISegment.get_segment_number())
                                                wsPRISequencesRow += 1
                                            else:
                                                
                                                for cSegmentAttribute in cPRISegment.get_attributes():
                                                    bSegmentAttribute = self.findAttribute(cSegmentAttribute.get_name(), bPRISegment)
                                                    
                                                    if not bSegmentAttribute:
                                                        
                                                        if priSequenceWritten == False:
                                                            self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                            wsPRISequencesRow += 1
                                                        
                                                        if priSegmentWritten == False:
                                                            self.writePRISegment(wsPRISequences, wsPRISequencesRow, bPRISegment, cPRISegment, priSegmentWritten)
                                                            wsPRISequencesRow += 1
                                                          
                                                        self.writeMissingBaseSegmentAttribute(wsPRISequences, wsPRISequencesRow, cSegmentAttribute.get_name(), cSegmentAttribute.get_value())    
                                                        wsPRISequencesRow += 1

                                for comparisonFREQSequence in comparisonGenerator.get_freq_sequences():
                                    baseFREQSequence = self.findFREQSequenceByOrdinalPos(comparisonFREQSequence.get_ordinal_pos(), baseGenerator)
                                    freqSequenceWritten = False
                                    
                                    if not baseFREQSequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeMissingBaseFREQSequence(wsGenerators, wsGeneratorsRow, comparisonFREQSequence.get_ordinal_pos())
                                        wsGeneratorsRow += 1

                                    else:
                                        
                                        for cFREQSeqAttribute in comparisonFREQSequence.get_attributes():
                                            bFREQSeqAttribute = self.findAttribute(cFREQSeqAttribute.get_name(), baseFREQSequence)
                                            
                                            if not bFREQSeqAttribute:
                                                
                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1
                                                 
                                                self.writeMissingBaseSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, cFREQSeqAttribute.get_name(), cFREQSeqAttribute.get_value())    
                                                wsFREQSequencesRow += 1

                                        for cFREQSegment in comparisonFREQSequence.get_segments():
                                            bFREQSegment = self.findSegmentBySegmentNumber(cFREQSegment.get_segment_number(), baseFREQSequence)
                                            freqSegmentWritten = False
                                            
                                            if not bFREQSegment:

                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1

                                                self.writeMissingBaseFREQSegment(wsFREQSequences, wsFREQSequencesRow, cFREQSegment.get_segment_number())
                                                wsFREQSequencesRow += 1
                                                
                                            else:
                                                
                                                for cSegmentAttribute in cFREQSegment.get_attributes():
                                                    bSegmentAttribute = self.findAttribute(cSegmentAttribute.get_name(), bFREQSegment)
                                                    
                                                    if not bSegmentAttribute:
                                                        
                                                        if freqSequenceWritten == False:
                                                            self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                            wsFREQSequencesRow += 1
                                                        
                                                        if freqSegmentWritten == False:
                                                            self.writeFREQSegment(wsFREQSequences, wsFREQSequencesRow, bFREQSegment, cFREQSegment, freqSegmentWritten)
                                                            wsFREQSequencesRow += 1
                                                            
                                                        self.writeMissingBaseSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, cSegmentAttribute.get_name(), cSegmentAttribute.get_value())    
                                                        wsFREQSequencesRow += 1

            
            wsEmitters.autofit('c')
            wsEmitters.autofit('r')

            wsModes.autofit('c')
            wsModes.autofit('r')

            wsGenerators.autofit('c')
            wsGenerators.autofit('r')

            wsPRISequences.autofit('c')
            wsPRISequences.autofit('r')

            wsFREQSequences.autofit('c')
            wsFREQSequences.autofit('r')

            
    
    def compareTheFiles(self):
    
        print("starting file comparison.")
        
        self.getShortFileNames()
        
        print("parsing Base File ...")
        self.parseBaseFile()
        
        print("parsing Comparison File ...")
        self.parseComparisonFile()
    
        xwApp = xw.App(visible=False)
        
        wb = xw.Book()
     
       
        self.compareEMTFiles(wb, self.base_emitters, self.comparison_emitters, constant.COMPARISON_ARRAY, constant.BASE_ARRAY)
    
        
        #self.compareEMTFiles(wf, self.comparison_emitters, self.base_emitters, constant.BASE_ARRAY, self.comparisonFileName, self.baseFileName)
    
       
        wb.save(r'EMT_Differences')
        wb.close()
        xwApp.quit()
    
        print("finished!")
               
    #if __name__ == '__main__':
     #   compareTheFiles()
    
                    
                
                
                       
                   

       