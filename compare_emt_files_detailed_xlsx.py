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
        
        print("isBase: {}".format(isBase))
        print("isBase == True: {}".format(isBase == True))
        print("isBase == False: {}".format(isBase == False))
        print("bfDisplay: {}".format(self.bfDisplay))
        print("cfDisplay: {}".format(self.cfDisplay))

        with open(fName) as f1:
            for cnt, line in enumerate(f1):
                if line.strip() == constant.EMITTER:
                    currentEntity = constant.EMITTER
                    if passNumber > 0:
                        if modePass > 0:
                            emitter.add_mode(emitter_mode)
                            
                        if isBase == False:
                            print("checkpoint 1 with elnot: {}".format(emitter.get_elnot()))
                            baseEmitter = self.findBaseElnot(emitter.get_elnot(), emitter_collection) 
                            if baseEmitter:
                                print("checkpoint 2")
                                localAttrDifferences = False
                                localModeDifferences = False
                                localDifferences = False
                                
                                baseEmitter.set_bfile(self.baseFileName)
                                print("baseEmitter.get_bfile(): {}".format(baseEmitter.get_bfile()))
                                
                                baseEmitter.set_cfile(self.cfDisplay)
                                emitter.set_bfile(self.baseFileName)
                                emitter.set_cfile(self.cfDisplay)
    
                                localAttrDifferences = baseEmitter.sync_attributes(emitter)
                                localModeDifferences = baseEmitter.sync_modes(emitter)
                                
                                if localAttrDifferences == True or localModeDifferences == True:
                                    localDifferences = True
                                    
                                baseEmitter.set_hasDifferences(localDifferences)
                                
                            else:
                                emitter.set_hasDifferences(True)
                                emitter_collection.append(emitter)

                        modePass = 0
                        
                    emitter = Emitter()
                    if isBase == True:
                        emitter.set_bfile(self.bfDisplay)
                    else:
                        emitter.set_cfile(self.cfDisplay)
                        
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
                                    emitter.set_bfile(self.baseFileName)
                                else:
                                    emitter.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                emitter.add_attribute(attrib)
                        
                        elif currentEntity == constant.EMITTER_MODE:
                       
                            if line_key.strip() == constant.MODE_NAME:
                                emitter_mode.set_mode_name(line_value)
                                if isBase == True:
                                    emitter_mode.set_bfile(self.baseFileName)
                                else:
                                    emitter_mode.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                emitter_mode.add_attribute(attrib)
                                
                        elif currentEntity == constant.GENERATOR:
    
                            if line_key.strip() == constant.GENERATOR_NUMBER:
                                generator.set_generator_number(line_value)
                                if isBase == True:
                                    generator.set_bfile(self.baseFileName)
                                else:
                                    generator.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                generator.add_attribute(attrib)
                                
                        elif currentEntity == constant.PRI_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                pri_sequence.set_number_of_segments(line_value)
                                if isBase == True:
                                    pri_sequence.set_bfile(self.baseFileName)
                                else:
                                    pri_sequence.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                pri_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.PRI_SEGMENT:
                            
                            if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                                pri_segment.set_segment_number(line_value)
                                if isBase == True:
                                    pri_segment.set_bfile(self.baseFileName)
                                else:
                                    pri_segment.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                pri_segment.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEQUENCE:
                            
                            if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                                freq_sequence.set_number_of_segments(line_value)
                                if isBase == True:
                                    freq_sequence.set_bfile(self.baseFileName)
                                else:
                                    freq_sequence.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                freq_sequence.add_attribute(attrib)
    
                        elif currentEntity == constant.FREQ_SEGMENT:
                            
                            if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                                freq_segment.set_segment_number(line_value)
                                if isBase == True:
                                    freq_segment.set_bfile(self.baseFileName)
                                else:
                                    freq_segment.set_cfile(self.cfDisplay)
                            else:
                                attrib = Attribute()
                                attrib.set_name(line_key)
                                if isBase == True:
                                    attrib.set_value(line_value)
                                else:
                                    attrib.set_cvalue(line_value)
                                freq_segment.add_attribute(attrib)
                                
            else:
                if isBase == False:
                    baseEmitter = self.findElnot(emitter.get_elnot(), emitter_collection) 
                    if baseEmitter:
                        localAttrDifferences = False
                        localModeDifferences = False
                        localDifferences = False
                        
                        baseEmitter.set_bfile(self.baseFileName)
                        baseEmitter.set_cfile(self.cfDisplay)
                        emitter.set_bfile(self.baseFileName)
                        emitter.set_cfile(self.cfDisplay)

                        localAttrDifferences = baseEmitter.sync_attributes(emitter)
                        localModeDifferences = baseEmitter.sync_modes(emitter)
                        
                        if localAttrDifferences == True or localModeDifferences == True:
                            localDifferences = True
                            
                        baseEmitter.set_hasDifferences(localDifferences)
                        
                    else:
                        emitter.set_hasDifferences(True)
                        emitter_collection.append(emitter)

    def parseBaseFile(self):
        print("baseFileName: {}".format(self.baseFileName))
        self.parseFile(self.baseFileName, self.base_emitters, True)
            
        
    def parseComparisonFile(self):
        self.parseFile(self.comparisonFileName, self.base_emitters, False)
        
          
    def findBaseElnot(self, elnotValue, emitter_collection):
        for emitter in emitter_collection:
            print("found elnot: {}".format(emitter.get_elnot()))
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
        self.writeLabelCell(ws, wsRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(ws, wsRow, constant.BASE_XL_COL_MODE_VAL, bEmitterModeName)
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(ws, wsRow, constant.COMP_XL_COL_MODE_VAL, cEmitterModeName)

        
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
        
        
    def writeEmitter(self, wsEmitters, wsEmittersRow, bElnot, cElnot, emitterWritten=False):
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
        emitterWritten = True
        

    def writeMode(self, wsModes, wsModesRow, bElnot, cElnot, baseEmitterModeName, comparisonEmitterModeName, modeWritten=False):
        self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)

        self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterModeName)
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterModeName)
        modeWritten = True
        
    
    def writeGenerator(self, wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten=False):
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
        
        
    def writePRISequence(self, wsPRISequences, wsPRISequencesRow, bElnot, modeName, generatorNumber, ordinalPos, priSequenceWritten=False):
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_ELNOT_VAL, bElnot)

        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        priSequenceWritten = True
        
    
    def writePRISegment(self, wsPRISequences, wsPRISequencesRow, segmentNumber, priSegmentWritten=False):
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        priSegmentWritten = True
        
    
    def writeFREQSequence(self, wsFREQSequences, wsFREQSequencesRow, bElnot, modeName, generatorNumber, ordinalPos, freqSequenceWritten=False):
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_ELNOT_VAL, bElnot)

        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_MODE_VAL, modeName)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_MODE_VAL, modeName)

        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_GENERATOR_VAL, generatorNumber)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_GENERATOR_VAL, generatorNumber)
        
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, ordinalPos)
        freqSequenceWritten = True
        

    def writeFREQSegment(self, wsFREQSequences, wsFREQSequencesRow, segmentNumber, freqSegmentWritten=False):
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, segmentNumber)
        freqSegmentWritten = True
        
            
    def displayDifferences(self, wb, base_emitter_collection):
         
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
            
            self.setColorMarkBoundary(wsEmittersBkgColorMarks, wsEmittersRow)
            self.setColorMarkBoundary(wsModesBkgColorMarks, wsModesRow)
            self.setColorMarkBoundary(wsGeneratorsBkgColorMarks, wsGeneratorsRow)
            self.setColorMarkBoundary(wsPRISequencesBkgColorMarks, wsPRISequencesRow)
            self.setColorMarkBoundary(wsFREQSequencesBkgColorMarks, wsFREQSequencesRow)
            
            print("emitter has differences: {}".format(baseEmitter.get_hasDifferences()))
            print("emitter bFile: {}".format(baseEmitter.get_bfile()))
            print("emitter cFile: {}".format(baseEmitter.get_cfile()))
            
                  
            if baseEmitter.get_hasDifferences() == True:
                if baseEmitter.get_cfile() == '':
                    self.writeMissingComparisonEmitter(wsEmitters, wsEmittersRow, bElnot)
                    wsEmittersRow += 1
                elif baseEmitter.get_bfile() == '':
                    self.writeMissingBaseEmitter(wsEmitters, wsEmittersRow, bElnot)
                    wsEmittersRow += 1
                else:
                    self.writeEmitter(wsEmitters, wsEmittersRow, bElnot, bElnot)
                    wsEmittersRow += 1

                    for baseAttribute in baseEmitter.get_attributes():
                        if baseAttribute.get_hasDifferences() == True:
                            if baseAttribute.get_cfile() == '':
                                self.writeMissingComparisonEmitterAttribute(wsEmitters, wsEmittersRow, baseAttribute.get_name(), baseAttribute.get_value())
                            elif baseAttribute.get_bfile() == '':
                                self.writeMissingBaseEmitterAttribute(wsEmitters, wsEmittersRow, baseAttribute.get_name(), baseAttribute.get_cvalue())
                            else:
                                self.writeFullEmitterAttribute(wsEmitters, wsEmittersRow, baseAttribute.get_name(), baseAttribute.get_value(), baseAttribute.get_name(), baseAttribute.get_cvalue())
                            wsEmittersRow += 1
                        
                for baseEmitterMode in baseEmitter.get_modes():
                    if baseEmitterMode.get_hasDifferences() == True:
                        if baseEmitterMode.get_cfile() == '':
                            self.writeMissingComparisonMode(wsEmitters, wsEmittersRow, baseEmitterMode.get_name())
                            wsEmittersRow += 1
                        elif baseEmitterMode.get_bfile() == '':
                            self.writeMissingBaseMode(wsEmitters, wsEmittersRow, baseEmitterMode.get_name())
                            wsEmittersRow += 1
                        else:
                            self.writeMode(wsModes, wsModesRow, bElnot, bElnot, baseEmitterMode.get_name(), baseEmitterMode.get_name())
                            wsModesRow += 1
                        
                            for baseModeAttribute in baseEmitterMode.get_attributes():
                                if baseModeAttribute.get_cfile() == '':
                                    self.writeMissingComparisonModeAttribute(wsModes, wsModesRow, baseModeAttribute.get_name(), baseModeAttribute.get_value())
                                elif baseModeAttribute.get_bfile() == '':
                                    self.writeMissingBaseModeAttribute(wsModes, wsModesRow, baseModeAttribute.get_name(), baseModeAttribute.get_cvalue())
                                else:
                                    self.writeFullModeAttribute(wsModes, wsModesRow, baseModeAttribute.get_name(), baseModeAttribute.get_value(), baseModeAttribute.get_name(), baseModeAttribute.get_cvalue())
                                wsModesRow += 1
                            
                        for baseGenerator in baseEmitterMode.get_generators():
                            if baseGenerator.get_hasDifferences() == True:
                                if baseGenerator.get_cfile() == '':
                                    self.writeMissingComparisonGenerator(wsModes, wsModesRow, baseGenerator.get_generator_number())
                                    wsModesRow += 1
                                elif baseGenerator.get_bfile() == '':
                                    self.writeMissingBaseGenerator(wsModes, wsModesRow, baseGenerator.get_generator_number())
                                    wsModesRow += 1
                                else:
                                    self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator)
                                    wsGeneratorsRow += 1                                    
                                
                                    for baseGeneratorAttribute in baseGenerator.get_attributes():
                                        if baseGeneratorAttribute.get_cfile() == '':
                                            self.writeMissingComparisonGeneratorAttribute(wsGenerators, wsGeneratorsRow, baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_value())
                                        elif baseGeneratorAttribute.get_bfile() == '':
                                            self.writeMissingBaseGeneratorAttribute(wsGenerators, wsGeneratorsRow, baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_cvalue())
                                        else:
                                            self.writeFullGeneratorAttribute(wsGenerators, wsGeneratorsRow, baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_value(), baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_cvalue())
                                        wsGeneratorsRow += 1
                                        
                                for basePRISequence in baseGenerator.get_pri_sequences():
                                    if basePRISequence.get_hasDifferences() == True:
                                        if basePRISequence.get_cfile() == '':
                                            self.writeMissingComparisonPRISequence(wsGenerators, wsGeneratorsRow, basePRISequence.get_ordinal_pos())
                                            wsGeneratorsRow += 1
                                        elif basePRISequence.get_bfile() == '':
                                            self.writeMissingBasePRISequence(wsGenerators, wsGeneratorsRow, basePRISequence.get_ordinal_pos())
                                            wsGeneratorsRow += 1
                                        else:
                                            self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos())
                                            wsPRISequencesRow += 1
                                            
                                        
                                            for bPRISeqAttribute in basePRISequence.get_attributes():
                                                if bPRISeqAttribute.get_cfile() == '':
                                                    self.writeMissingComparisonSequenceAttribute(wsPRISequences, wsPRISequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value())
                                                elif bPRISeqAttribute.get_bfile() == '':
                                                    self.writeMissingBaseSequenceAttribute(wsPRISequences, wsPRISequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_cvalue())
                                                else:
                                                    self.writeFullSequenceAttribute(wsPRISequences, wsPRISequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value(), bPRISeqAttribute.get_name(), bPRISeqAttribute.get_cvalue())
                                                wsPRISequencesRow += 1
                                                    
                                        for bPRISegment in basePRISequence.get_segments():
                                            if bPRISegment.get_hasDifferences() == True:
                                                if bPRISegment.get_cfile() == '':
                                                    self.writeMissingComparisonSegment(wsPRISequences, wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    wsPRISequencesRow += 1
                                                elif bPRISegment.get_bfile() == '':
                                                    self.writeMissingBaseSegment(wsPRISequences, wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    wsPRISequencesRow += 1
                                                else:
                                                    self.writePRISegment(wsPRISequences, wsPRISequencesRow, bPRISegment.get_segment_number())
                                                    wsPRISequencesRow += 1
                                                    
                                                    for bSegmentAttribute in bPRISegment.get_attributes():
                                                        if bSegmentAttribute.get_cfile() == '':
                                                            self.writeMissingComparisonSegmentAttribute(wsPRISequences, wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())                                                            
                                                        elif bSegmentAttribute.get_bfile() == '':
                                                            self.writeMissingBaseSegmentAttribute(wsPRISequences, wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())                                                            
                                                        else:
                                                            self.writeFullSegmentAttribute(wsPRISequences, wsPRISequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())
                                                        wsPRISequencesRow += 1
                                                            
                                for baseSequence in baseGenerator.get_freq_sequences():
                                    if baseSequence.get_hasDifferences() == True:
                                        if baseSequence.get_cfile() == '':
                                            self.writeMissingComparisonFREQSequence(wsGenerators, wsGeneratorsRow, baseSequence.get_ordinal_pos())
                                            wsGeneratorsRow += 1
                                        elif basePRISequence.get_bfile() == '':
                                            self.writeMissingBaseFREQSequence(wsGenerators, wsGeneratorsRow, baseSequence.get_ordinal_pos())
                                            wsGeneratorsRow += 1
                                        else:
                                            self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseSequence.get_ordinal_pos())
                                            wsFREQSequencesRow += 1
                                            
                                        
                                            for bSeqAttribute in baseSequence.get_attributes():
                                                if bSeqAttribute.get_cfile() == '':
                                                    self.writeMissingComparisonSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value())
                                                elif bSeqAttribute.get_bfile() == '':
                                                    self.writeMissingBaseSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, bPRISeqAttribute.get_name(), bPRISeqAttribute.get_cvalue())
                                                else:
                                                    self.writeFullSequenceAttribute(wsFREQSequences, wsFREQSequencesRow, bSeqAttribute.get_name(), bSeqAttribute.get_value(), bSeqAttribute.get_name(), bSeqAttribute.get_cvalue())
                                                wsFREQSequencesRow += 1
                                                    
                                        for bSegment in baseSequence.get_segments():
                                            if bSegment.get_hasDifferences() == True:
                                                if bSegment.get_cfile() == '':
                                                    self.writeMissingComparisonSegment(wsFREQSequences, wsFREQSequencesRow, bSegment.get_segment_number())
                                                    wsFREQSequencesRow += 1
                                                elif bPRISegment.get_bfile() == '':
                                                    self.writeMissingBaseSegment(wsFREQSequences, wsFREQSequencesRow, bSegment.get_segment_number())
                                                    wsFREQSequencesRow += 1
                                                else:
                                                    self.writePRISegment(wsFREQSequences, wsFREQSequencesRow, bSegment.get_segment_number())
                                                    wsFREQSequencesRow += 1
                                                    
                                                    for bSegmentAttribute in bSegment.get_attributes():
                                                        if bSegmentAttribute.get_cfile() == '':
                                                            self.writeMissingComparisonSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value())                                                            
                                                        elif bSegmentAttribute.get_bfile() == '':
                                                            self.writeMissingBaseSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())                                                            
                                                        else:
                                                            self.writeFullSegmentAttribute(wsFREQSequences, wsFREQSequencesRow, bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), bSegmentAttribute.get_name(), bSegmentAttribute.get_cvalue())
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
     
        self.displayDifferences(wb, self.base_emitters)
    
       
        wb.save(r'EMT_Differences')
        wb.close()
        xwApp.quit()
    
        print("finished!")
               
    #if __name__ == '__main__':
     #   compareTheFiles()
    
                    
                
                
                       
                   

       