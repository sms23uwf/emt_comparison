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

import xlwings as xw
from xlwings import constants
from xlwings.utils import rgb_to_int
import print_utility


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
    
        
    def findAttribute(self, baseAttributeName, comparisonEmitter):
        for comparisonAttribute in comparisonEmitter.get_attributes():
            if comparisonAttribute.get_name() == baseAttributeName:
                return comparisonAttribute
            
        return []
    
            
    # def writeCell(self, ws, wsRow, cellValue):
    #     ws.range(wsRow, 1).value = cellValue
    #     ws.range(wsRow, 1).api.VerticalAlignment = constants.VAlign.xlVAlignTop
    #     #ws.range(wsRow, 1).WrapText = True

        
    # def writeSpecificCell(self, ws, wsRow, wsCol, cellValue):
    #     ws.range(wsRow, wsCol).value = cellValue
    #     ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignCenter
    #     ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    #     ws.range(wsRow, wsCol).WrapText = True

    # def writeValueCell(self, ws, wsRow, wsCol, cellValue):
    #     ws.range(wsRow, wsCol).value = cellValue
    #     try:
    #         ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignLeft
    #         ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
    #         ws.range(wsRow, wsCol).WrapText = True
    #     except Exception:
    #         print("exception was thrown")
            

    # def writeLabelCell(self, ws, wsRow, wsCol, cellValue):
    #     ws.range(wsRow, wsCol).value = cellValue
    #     try:
    #         ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignRight
    #         ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
    #         ws.range(wsRow, wsCol).WrapText = True
    #         if cellValue == constant.XL_MISSING_TEXT:
    #             ws.range(wsRow, wsCol).api.Font.Color = rgb_to_int((255,0,0))

    #     except Exception:
    #         print("exception was thrown")

    
    
    
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
        print_utility.writeSpecificCell(ws, 1, 1, bfTitle)
        
        ws.range((1,13), (1,25)).merge(across=True)
        print_utility.writeSpecificCell(ws, 1, 13, cfTitle)



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
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
        
    
    def writeMissingComparisonEmitterAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullEmitterAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, cAttributeValue)
        
        
    def writeMissingComparisonMode(self, ws, wsRow, bEmitterModeName):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_VAL, bEmitterModeName)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
        
         
    def writeFullEmitterMode(self, ws, wsRow, bEmitterModeName, cEmitterModeName):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_VAL, bEmitterModeName)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_VAL, cEmitterModeName)

        
    def writeMissingComparisonModeAttribute(self, ws, wsRow, bModeAttributeName, bModeAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, bModeAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, bModeAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullModeAttribute(self, ws, wsRow, bModeAttributeName, bModeAttributeValue, cModeAttributeName, cModeAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, bModeAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, bModeAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, cModeAttributeName)
        print_utility.writeValuelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, cModeAttributeValue)
        

    def writeMissingComparisonGenerator(self, ws, wsRow, bGeneratorNumber):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_VAL, bGeneratorNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)


    def writeMissingComparisonGeneratorAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)

        
    def writeFullGeneratorAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, cAttributeValue)

        
    def writeFullPRISequenceCounts(self, ws, wsRow, bgPRISequenceCount, cpPRISequenceCount):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, bgPRISequenceCount)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, cpPRISequenceCount)
        

    def writeFullFREQSequenceCounts(self,  ws, wsRow, bgFREQSequenceCount, cpFREQSequenceCount):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, bgFREQSequenceCount)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, cpFREQSequenceCount)

      
    def writeMissingComparisonPRISequence(self, ws, wsRow, sequencePos):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        

    def writeMissingComparisonSequenceAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeFullSequenceAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cAttributeValue)
                                                   
    
    def writeMissingComparisonSegment(self, ws, wsRow, bSegmentNumber):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bSegmentNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeMissingComparisonSegmentAttribute(self, ws, wsRow, bAttributeName, bAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
                 
    def writeFullSegmentAttribute(self, ws, wsRow, bAttributeName, bAttributeValue, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cAttributeValue)
        

    def writeMissingComparisonFREQSequence(self, ws, wsRow, sequencePos):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseEmitter(self, ws, wsRow, cElnot):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
        

    def writeMissingBaseEmitterAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, cAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseMode(self, ws, wsRow, cModeName):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_VAL, cModeName)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
        
            
    def writeMissingBaseModeAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, cAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
                   
    def writeMissingBaseGenerator(self, ws, wsRow, cGeneratorNumber):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_VAL, cGeneratorNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)

    def writeMissingBaseGeneratorAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, cAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)


    def writeMissingBasePRISequence(self, ws, wsRow, sequencePos):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBaseSequenceAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        

    def writeMissingBasePRISegment(self, ws, wsRow, cPRISegmentNumber):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cPRISegmentNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
                                                
        
    def writeMissingBaseSegmentAttribute(self, ws, wsRow, cAttributeName, cAttributeValue):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cAttributeName)
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cAttributeValue)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeMissingBaseFREQSequence(self, ws, wsRow, sequencePos):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, sequencePos)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)


    def writeMissingBaseFREQSegment(self, ws, wsRow, cFREQSegmentNumber):
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cFREQSegmentNumber)
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
        
        
    def writeEmitter(self, ws, wsRow, bElnot, cElnot, emitterWritten=False):
        print_utility.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(ws, wsRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
        emitterWritten = True
        

    def writeMode(self, wsModes, wsModesRow, bElnot, cElnot, baseEmitterModeName, comparisonEmitterModeName, modeWritten=False):
        print_utility.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)

        print_utility.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterModeName)
        print_utility.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterModeName)
        modeWritten = True
        
    
    def writeGenerator(self, wsGenerators, wsGeneratorsRow, bElnot, modeName, generatorNumber, generatorWritten=False):
        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_ELNOT_VAL, bElnot)

        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        print_utility.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        generatorWritten = True
        
        
    def writePRISequence(self, wsPRISequences, wsPRISequencesRow, bElnot, modeName, generatorNumber, ordinalPos, priSequenceWritten=False):
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_VAL, bElnot)

        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        priSequenceWritten = True
        
    
    def writePRISegment(self, wsPRISequences, wsPRISequencesRow, segmentNumber, priSegmentWritten=False):
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        print_utility.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        print_utility.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        priSegmentWritten = True
        
    
    def writeFREQSequence(self, wsFREQSequences, wsFREQSequencesRow, bElnot, modeName, generatorNumber, ordinalPos, freqSequenceWritten=False):
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_VAL, bElnot)

        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        freqSequenceWritten = True
        

    def writeFREQSegment(self, wsFREQSequences, wsFREQSequencesRow, segmentNumber, freqSegmentWritten=False):
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        print_utility.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        print_utility.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        freqSegmentWritten = True
        
            
    def displayDifferences(self, wb):
         
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
       
        #wsEmittersRow = 2
        #wsModesRow = 2
        #wsGeneratorsRow = 2
        # wsPRISequencesRow = 2
        # wsFREQSequencesRow = 2
        cellValue = ""
        
        wsEmittersBkgColorMarks = []
        wsModesBkgColorMarks = []
        wsGeneratorsBkgColorMarks = []
        wsPRISequencesBkgColorMarks = []
        wsFREQSequencesBkgColorMarks = []
        
        
        
        for baseEmitter in self.base_emitters:
            bElnot = baseEmitter.get_elnot()
            
            self.setColorMarkBoundary(wsEmittersBkgColorMarks, print_utility.wsEmittersRow)
            self.setColorMarkBoundary(wsModesBkgColorMarks, print_utility.wsModesRow)
            self.setColorMarkBoundary(wsGeneratorsBkgColorMarks, print_utility.wsGeneratorsRow)
            self.setColorMarkBoundary(wsPRISequencesBkgColorMarks, print_utility.wsPRISequencesRow)
            self.setColorMarkBoundary(wsFREQSequencesBkgColorMarks, print_utility.wsFREQSequencesRow)
            
            if baseEmitter.get_hasDifferences() == True:
                baseEmitter.print_emitter(wsEmitters, constant.BASE_XL_COL_ELNOT_LBL, constant.BASE_XL_COL_ELNOT_VAL, constant.COMP_XL_COL_ELNOT_LBL, constant.COMP_XL_COL_ELNOT_VAL)

                
                    
                for baseEmitterMode in baseEmitter.get_modes():
                    if baseEmitterMode.get_hasDifferences() == True:
                        baseEmitterMode.print_mode(wsModes, bElnot, bElnot)

                        for baseGenerator in baseEmitterMode.get_generators():
                            if baseGenerator.get_hasDifferences() == True:
                                baseGenerator.print_generator(wsGenerators, bElnot, baseEmitterMode.get_name())
                                
                                for basePRISequence in baseGenerator.get_pri_sequences():
                                    if basePRISequence.get_hasDifferences() == True:
                                        basePRISequence.print_pri_sequence(wsPRISequences, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number())
                                                    
                                        for bPRISegment in basePRISequence.get_segments():
                                            if bPRISegment.get_hasDifferences() == True:
                                                if bPRISegment.get_cfile() == False:
                                                    self.writeMissingComparisonSegment(wsPRISequences, print_utility.wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    print_utility.wsPRISequencesRow += 1
                                                elif bPRISegment.get_bfile() == False:
                                                    self.writeMissingBaseSegment(wsPRISequences, print_utility.wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    print_utility.wsPRISequencesRow += 1
                                                else:
                                                    self.writePRISegment(wsPRISequences, print_utility.wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    print_utility.wsPRISequencesRow += 1
                                                    
                                                    for bSegmentAttribute in bPRISegment.get_attributes():
                                                        if bSegmentAttribute.get_cfile() == False:
                                                            self.writeMissingComparisonSegmentAttribute(wsPRISequences, print_utility.wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())                                                            
                                                        elif bSegmentAttribute.get_bfile() == False:
                                                            self.writeMissingBaseSegmentAttribute(wsPRISequences, print_utility.wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())                                                            
                                                        else:
                                                            self.writeFullSegmentAttribute(wsPRISequences, print_utility.wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())
                                                        print_utility.wsPRISequencesRow += 1
                                                            
                                for baseSequence in baseGenerator.get_freq_sequences():
                                    if baseSequence.get_hasDifferences() == True:
                                        baseSequence.print_freq_sequence(wsFREQSequences, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number())
                                        
                                        for bSegment in baseSequence.get_segments():
                                            if bSegment.get_hasDifferences() == True:
                                                if bSegment.get_cfile() == False:
                                                    self.writeMissingComparisonSegment(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegment.get_segment_number())
                                                    print_utility.wsFREQSequencesRow += 1
                                                elif bSegment.get_bfile() == False:
                                                    self.writeMissingBaseSegment(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegment.get_segment_number())
                                                    print_utility.wsFREQSequencesRow += 1
                                                else:
                                                    self.writePRISegment(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegment.get_segment_number())
                                                    print_utility.wsFREQSequencesRow += 1
                                                    
                                                    for bSegmentAttribute in bSegment.get_attributes():
                                                        if bSegmentAttribute.get_cfile() == False:
                                                            self.writeMissingComparisonSegmentAttribute(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())                                                            
                                                        elif bSegmentAttribute.get_bfile() == False:
                                                            self.writeMissingBaseSegmentAttribute(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())                                                            
                                                        else:
                                                            self.writeFullSegmentAttribute(wsFREQSequences, print_utility.wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())
                                                        print_utility.wsFREQSequencesRow += 1
                                                            

            
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
     
        self.displayDifferences(wb)
    
       
        wb.save(r'EMT_Differences')
        wb.close()
        xwApp.quit()
    
        print("finished!")
               
    #if __name__ == '__main__':
     #   compareTheFiles()
    
                    
                
                
                       
                   

       