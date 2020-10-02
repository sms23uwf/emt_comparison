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
from sequence import Sequence
from segment import Segment
from color_mark_row_boundary import ColorMarkRowBoundary

import print_utility

import pandas as pd


class CompareEMTFiles():
    def __init__(self, bFilePath, cFilePath):
        self.baseFileName = bFilePath[0]
        self.comparisonFileName = cFilePath[0]
        self.compareTheFiles()
        self.bfDisplay = self.baseFileName
        self.cfDisplay = self.comparisonFileName

      
    base_emitters = []
    output_emitters = []
    
    currentEntity = constant.EMITTER
    
    def lineAsKeyValue(self, line):
        line_kv = line.split(constant.VALUE_SEPARATOR)
        return line_kv
    
        
    def parseFile(self, fName, isBase):
        emitter = Emitter()
        emitter_mode = EmitterMode()
        generator = Generator()
        pri_sequence = Sequence()
        freq_sequence = Sequence()
        pri_segment = Segment()
        freq_segment = Segment()
            
        passNumber = 0
        modePass = 0
        generatorPass = 0
        priSeqPass = 0
        freqSeqPass = 0
        priSegmentPass = 0
        freqSegmentPass = 0
        
        emitter_collection = self.base_emitters
        
        with open(fName) as f1:
            for cnt, line in enumerate(f1):
                if line.strip() == constant.EMITTER:
                    currentEntity = constant.EMITTER
                    if passNumber > 0:
                        if modePass > 0:
                            emitter.add_mode(emitter_mode)
                            modePass = 0
                            
                        if isBase == False:
                            baseEmitter = self.findBaseElnot(emitter.get_elnot()) 
                            if baseEmitter:
                                localAttrDifferences = False
                                localModeDifferences = False
                                localDifferences = False
                                
                                baseEmitter.set_bfile(True)
                                baseEmitter.set_cfile(True)
                                emitter.set_bfile(True)
                                emitter.set_cfile(True)
                                
                                localAttrDifferences = baseEmitter.sync_attributes(emitter)
                                localModeDifferences = baseEmitter.sync_modes(emitter)
                                
                                if localAttrDifferences == True or localModeDifferences == True:
                                    localDifferences = True
                                    
                                baseEmitter.set_hasDifferences(localDifferences)
                            else:
                                emitter.set_hasDifferences(True)
                                emitter.set_cfile(True)
                                self.base_emitters.append(emitter)

                        else:
                            self.base_emitters.append(emitter)

                        
                    emitter = Emitter()
                    emitter.set_bfile(True) if isBase == True else emitter.set_cfile(True)

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
                        
                    pri_sequence = Sequence()
                    pri_sequence.set_ordinal_pos(priSeqPass)
                    priSeqPass += 1
                    
                elif line.strip() == constant.PRI_SEGMENT:
                    currentEntity = constant.PRI_SEGMENT
                    if priSegmentPass > 0:
                        pri_sequence.add_segment(pri_segment)
                        
                    pri_segment = Segment()
                    priSegmentPass += 1
    
                elif line.strip() == constant.FREQ_SEQUENCE:
                    currentEntity = constant.FREQ_SEQUENCE
                    if freqSeqPass > 0:
                        if freqSegmentPass > 0:
                            freq_sequence.add_segment(freq_segment)
    
                        generator.add_freq_sequence(freq_sequence)
                        freqSegmentPass = 0
                        
                    freq_sequence = Sequence()
                    freq_sequence.set_ordinal_pos(freqSeqPass)
                    freqSeqPass += 1
                    
                elif line.strip() == constant.FREQ_SEGMENT:
                    currentEntity = constant.FREQ_SEGMENT
                    if freqSegmentPass > 0:
                        freq_sequence.add_segment(freq_segment)
                        
                    freq_segment = Segment()
                    freqSegmentPass += 1
                    
                else:
                    if line.strip().__contains__(constant.VALUE_SEPARATOR):
                        line_kv = self.lineAsKeyValue(line.strip())
                        line_key = line_kv[0].strip()
                        line_value = line_kv[1].strip()
                        
                        if currentEntity == constant.EMITTER:
     
                            if line_key.strip() == constant.EMITTER_ELNOT:
                                emitter.set_elnot(line_value)
                                emitter.set_bfile(True) if isBase == True else emitter.set_cfile(True)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                                emitter.add_attribute(attrib)
                        
                        elif currentEntity == constant.EMITTER_MODE:
                       
                            if line_key.strip() == constant.MODE_NAME:
                                emitter_mode.set_mode_name(line_value)
                                emitter_mode.set_bfile(True) if isBase == True else emitter_mode.set_cfile(True)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                                emitter_mode.add_attribute(attrib)
                                
                        elif currentEntity == constant.GENERATOR:
    
                            if line_key.strip() == constant.GENERATOR_NUMBER:
                                generator.set_generator_number(line_value)
                                generator.set_bfile(True) if isBase == True else generator.set_cfile(True)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                                generator.add_attribute(attrib)
                                
                        elif currentEntity == constant.PRI_SEQUENCE:
                            pri_sequence.set_bfile(True) if isBase == True else pri_sequence.set_cfile(True)
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                            pri_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.PRI_SEGMENT:
                            
                            if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                                pri_segment.set_segment_number(line_value)
                                pri_segment.set_bfile(True) if isBase == True else pri_segment.set_cfile(True)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                                pri_segment.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEQUENCE:
                            freq_sequence.set_bfile(True) if isBase == True else freq_sequence.set_cfile(True)
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                            freq_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEGMENT:
                            
                            if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                                freq_segment.set_segment_number(line_value)
                                freq_segment.set_bfile(True) if isBase == True else freq_segment.set_cfile(True)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                attrib.set_value(line_value) if isBase == True else attrib.set_cvalue(line_value)
                                freq_segment.add_attribute(attrib)
                                
            else:
                if isBase == False:
                    baseEmitter = self.findBaseElnot(emitter.get_elnot()) 
                    if baseEmitter:
                        localAttrDifferences = False
                        localModeDifferences = False
                        localDifferences = False
                        
                        baseEmitter.set_bfile(True)
                        baseEmitter.set_cfile(True)
                        emitter.set_bfile(True)
                        emitter.set_cfile(True)

                        localAttrDifferences = baseEmitter.sync_attributes(emitter)
                        localModeDifferences = baseEmitter.sync_modes(emitter)
                        
                        if localAttrDifferences == True or localModeDifferences == True:
                            localDifferences = True
                            
                        baseEmitter.set_hasDifferences(localDifferences)
                        
                    else:
                        emitter.set_hasDifferences(True)
                        self.base_emitters.append(emitter)

    def parseBaseFile(self):
        self.parseFile(self.baseFileName, True)
            
        
    def parseComparisonFile(self):
        self.parseFile(self.comparisonFileName, False)
        
          
    def findBaseElnot(self, elnotValue):
        for emitter in self.base_emitters:
            if emitter.get_elnot() == elnotValue:
                return emitter
            
        return []
    
    
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
            
    
    def compareTheFiles(self):
    
        print("starting file comparison.")
        
        self.getShortFileNames()
        
        print("parsing Base File ...")
        self.parseBaseFile()
        
        print("parsing Comparison File ...")
        self.parseComparisonFile()
    
        dfsWB = print_utility.dictsToPandasDFs(self.base_emitters, self.bfDisplay, self.cfDisplay)
        writer = pd.ExcelWriter('EMT_Differences.xlsx', engine='xlsxwriter')
        
        dfEmitters = dfsWB[0]
        dfModes = dfsWB[1]
        dfGenerators = dfsWB[2]
        dfPRISequences = dfsWB[3]
        dfFREQSequences = dfsWB[4]
        
     
        print("outputting excel display of differences ...")
        dfEmitters.to_excel(writer, sheet_name='Emitters')
        dfModes.to_excel(writer, sheet_name='Modes')
        dfGenerators.to_excel(writer, sheet_name='Generators')
        dfPRISequences.to_excel(writer, sheet_name='PRI Sequences')
        dfFREQSequences.to_excel(writer, sheet_name='FREQ Sequences')
        
        writer.save()
    
        print("finished!")
               
    #if __name__ == '__main__':
     #   compareTheFiles()
       