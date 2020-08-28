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
    
    
    def parseBaseFile(self):
        print("baseFileName: {}".format(self.baseFileName))
        self.parseFile(self.baseFileName, self.base_emitters, True)
            
        
    def parseComparisonFile(self):
        self.parseFile(self.comparisonFileName, self.comparison_emitters, False)
        
              
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
                self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_LBL, "ELNOT:")
                self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_VAL, bElnot)
                self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
                wsEmittersRow += 1
               
            else:

                for baseAttribute in baseEmitter.get_attributes():
                    comparisonAttribute = self.findAttribute(baseAttribute.get_name(), comparisonEmitter)
                    
                    if not comparisonAttribute:
                        if emitterWritten == False:
                            self.writeEmitter(wsEmitters, wsEmittersRow, bElnot, comparisonEmitter.get_elnot(), emitterWritten)
                            wsEmittersRow += 1
                            
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, baseAttribute.get_name())
                        self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, baseAttribute.get_value())
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                        wsEmittersRow += 1
                        
                    else:
                        if comparisonAttribute.get_value() != baseAttribute.get_value():

                            if emitterWritten == False:
                                self.writeEmitter(wsEmitters, wsEmittersRow, bElnot, comparisonEmitter.get_elnot(), emitterWritten)
                                wsEmittersRow += 1

                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, baseAttribute.get_name())
                            self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_VAL, baseAttribute.get_value())
                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, comparisonAttribute.get_name())
                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, comparisonAttribute.get_value())
                            wsEmittersRow += 1

                            
                for baseEmitterMode in baseEmitter.get_modes():
                    comparisonEmitterMode = self.findEmitterMode(baseEmitterMode.get_name(), comparisonEmitter)
                    
                    modeWritten = False
                    
                    if not comparisonEmitterMode:
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
                        self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterMode.get_name())
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
                        wsEmittersRow += 1
                    else:
                        if comparisonEmitterMode.get_name() != baseEmitterMode.get_name():
                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_LBL, "MODE:")
                            self.writeValueCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_VAL, baseEmitterMode.get_name())
                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
                            self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterMode.get_name())
                            wsEmittersRow += 1
                        
                        
                        for baseModeAttribute in baseEmitterMode.get_attributes():
                            comparisonModeAttribute = self.findAttribute(baseModeAttribute.get_name(), comparisonEmitterMode)
                            
                            if not comparisonModeAttribute:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                    
                                self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, baseModeAttribute.get_name())
                                self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, baseModeAttribute.get_value())
                                self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                
                                wsModesRow += 1
                            else:
                                if comparisonModeAttribute.get_value() != baseModeAttribute.get_value():
                                    
                                    if modeWritten == False:
                                        self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                        wsModesRow += 1
                                    
                                    self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, baseModeAttribute.get_name())
                                    self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_ATTRIB_VAL, baseModeAttribute.get_value())
                                    self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, comparisonModeAttribute.get_name())
                                    self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, comparisonModeAttribute.get_value())
                                    
                                    wsModesRow += 1
                                    
                        for baseGenerator in baseEmitterMode.get_generators():
                            comparisonGenerator = self.findGenerator(baseGenerator.get_generator_number(), comparisonEmitterMode)
                        
                            generatorWritten = False
                            
                            if not comparisonGenerator:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                
                                self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_GENERATOR_LBL, "GENERATOR:")
                                self.writeValueCell(wsModes, wsModesRow, constant.BASE_XL_COL_GENERATOR_VAL, baseGenerator.get_generator_number())
                                self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)
                                
                                wsModesRow += 1
                            else:
                                
                                for baseGeneratorAttribute in baseGenerator.get_attributes():
                                    comparisonGeneratorAttribute = self.findAttribute(baseGeneratorAttribute.get_name(), comparisonGenerator)
                                    
                                    if not comparisonGeneratorAttribute:
                                        
                                        
                                        if generatorsWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                                
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, baseGeneratorAttribute.get_name())
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, baseGeneratorAttribute.get_value())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                        wsGeneratorsRow += 1
                                        
                                    else:
                                        
                                        if comparisonGeneratorAttribute.get_value() != baseGeneratorAttribute.get_value():

                                            if generatorWritten == False:
                                                self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                                wsGeneratorsRow += 1

                                            self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, baseGeneratorAttribute.get_name())
                                            self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_VAL, baseGeneratorAttribute.get_value())
                                            self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, comparisonGeneratorAttribute.get_name())
                                            self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, comparisonGeneratorAttribute.get_value())
                                            wsGeneratorsRow += 1
                            
                                bgPRISequences = baseGenerator.get_pri_sequences()
                                bgFREQSequences = baseGenerator.get_freq_sequences()
                                cpPRISequences = comparisonGenerator.get_pri_sequences()
                                cpFREQSequences = comparisonGenerator.get_freq_sequences()
                                
                                if len(bgPRISequences) != len(cpPRISequences):

                                    if generatorWritten == False:
                                        self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                        wsGeneratorsRow += 1

                                    self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
                                    self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, len(bgPRISequences))
                                    self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCES[]:")
                                    self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, len(cpPRISequences))
                                    wsGeneratorsRow += 1
        
                                if len(bgFREQSequences) != len(cpFREQSequences):
                                    
                                    if generatorWritten == False:
                                        self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                        wsGeneratorsRow += 1
                                    
                                    self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
                                    self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, len(bgFREQSequences))
                                    self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCES[]:")
                                    self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, len(cpFREQSequences))
                                    wsGeneratorsRow += 1
                                
                                for basePRISequence in baseGenerator.get_pri_sequences():
                                    comparisonPRISequence = self.findPRISequenceByOrdinalPos(basePRISequence.get_ordinal_pos(), comparisonGenerator)
                                    priSequenceWritten = False
                                    
                                    if not comparisonPRISequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, basePRISequence.get_ordinal_pos())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        
                                        wsGeneratorsRow += 1
                                    else:
                                        
                                        for bPRISeqAttribute in basePRISequence.get_attributes(): 
                                            cPRISeqAttribute = self.findAttribute(bPRISeqAttribute.get_name(), comparisonPRISequence)
                                            
                                            if not cPRISeqAttribute:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                    
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bPRISeqAttribute.get_name())
                                                self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bPRISeqAttribute.get_value())
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                wsPRISequencesRow += 1

                                            else:

                                                if cPRISeqAttribute.get_value () != bPRISeqAttribute.get_value():
                                                    
                                                    if priSequenceWritten == False:
                                                        self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                        wsPRISequencesRow += 1
                                                    
                                                    self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bPRISeqAttribute.get_name())
                                                    self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bPRISeqAttribute.get_value())
                                                    self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cPRISeqAttribute.get_name())
                                                    self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cPRISeqAttribute.get_value())
                                                   
                                                    wsPRISequencesRow += 1
        
                                        for bPRISegment in basePRISequence.get_segments():
                                            cPRISegment = self.findSegmentBySegmentNumber(bPRISegment.get_segment_number(), comparisonPRISequence)
                                            priSegmentWritten = False
                                            
                                            if not cPRISegment:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
                                                self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bPRISegment.get_segment_number())
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
                                                
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
                                                            
                                                        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bSegmentAttribute.get_name())
                                                        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bSegmentAttribute.get_value())
                                                        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                        wsPRISequencesRow += 1

                                                    else:

                                                        if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
                                                            if priSequenceWritten == False:
                                                                self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                                wsPRISequencesRow += 1
                                                            
                                                            if priSegmentWritten == False:
                                                                self.writePRISegment(wsPRISequences, wsPRISequencesRow, bPRISegment, cPRISegment, priSegmentWritten)
                                                                wsPRISequencesRow += 1
                                                            
                                                            self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bSegmentAttribute.get_name())
                                                            self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bSegmentAttribute.get_value())
                                                            self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cSegmentAttribute.get_name())
                                                            self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cSegmentAttribute.get_value())
                                                            
                                                            wsPRISequencesRow += 1
                    
                                for baseFREQSequence in baseGenerator.get_freq_sequences():
                                    comparisonFREQSequence = self.findFREQSequenceByOrdinalPos(baseFREQSequence.get_ordinal_pos(), comparisonGenerator)
                                    freqSequenceWritten = False
                                    
                                    if not comparisonFREQSequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_VAL, baseFREQSequence.get_ordinal_pos())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        wsGeneratorsRow += 1

                                    else:
                                        
                                        for bFREQSeqAttribute in baseFREQSequence.get_attributes():
                                            cFREQSeqAttribute = self.findAttribute(bFREQSeqAttribute.get_name(), comparisonFREQSequence)
                                            
                                            if not cFREQSeqAttribute:
                                                
                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1
                                                    
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bFREQSeqAttribute.get_name())
                                                self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bFREQSeqAttribute.get_value())
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                wsFREQSequencesRow += 1

                                            else:
                                                if cFREQSeqAttribute.get_value() != bFREQSeqAttribute.get_value():
                                                    
                                                    if freqSequenceWritten == False:
                                                        self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                        wsFREQSequencesRow += 1
                                                    
                                                    self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, bFREQSeqAttribute.get_name())
                                                    self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_VAL, bFREQSeqAttribute.get_value())
                                                    self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cFREQSeqAttribute.get_name())
                                                    self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cFREQSeqAttribute.get_value())
                                                    wsFREQSequencesRow += 1
        
                                        for bFREQSegment in baseFREQSequence.get_segments():
                                            cFREQSegment = self.findSegmentBySegmentNumber(bFREQSegment.get_segment_number(), comparisonFREQSequence)
                                            freqSegmentWritten = False
                                            
                                            if not cFREQSegment:

                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1

                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
                                                self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_VAL, bFREQSegment.get_segment_number())
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
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
                                                            
                                                        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bSegmentAttribute.get_name())
                                                        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bSegmentAttribute.get_value())
                                                        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                        wsFREQSequencesRow += 1

                                                    else:

                                                        if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                            
                                                            if freqSequenceWritten == False:
                                                                self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                                wsFREQSequencesRow += 1
                                                            
                                                            if freqSegmentWritten == False:
                                                                self.writeFREQSegment(wsFREQSequences, wsFREQSequencesRow, bFREQSegment, cFREQSegment, freqSegmentWritten)
                                                                wsFREQSequencesRow += 1
                                                            
                                                            self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, bSegmentAttribute.get_name())
                                                            self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_VAL, bSegmentAttribute.get_value())
                                                            self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cSegmentAttribute.get_name())
                                                            self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cSegmentAttribute.get_value())
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
                wsEmittersRow += 1
                self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_ELNOT_LBL, constant.XL_MISSING_TEXT)
                self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_LBL, "ELNOT:")
                self.writeValueCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_ELNOT_VAL, cElnot)
                wsEmittersRow += 1
               
            else:

                for comparisonAttribute in comparisonEmitter.get_attributes():
                    baseAttribute = self.findAttribute(comparisonAttribute.get_name(), baseEmitter)
                    
                    if not baseAttribute:
                        if emitterWritten == False:
                            self.writeEmitter(wsEmitters, wsEmittersRow, cElnot, baseEmitter.get_elnot(), emitterWritten)
                            wsEmittersRow += 1
                            
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_EMITTER_ATTRIB_LBL, comparisonAttribute.get_name())
                        self.writeValueCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_EMITTER_ATTRIB_VAL, comparisonAttribute.get_value())
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_EMITTER_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                        wsEmittersRow += 1
                        
                            
                for comparisonEmitterMode in comparisonEmitter.get_modes():
                    baseEmitterMode = self.findEmitterMode(comparisonEmitterMode.get_name(), baseEmitter)
                    
                    modeWritten = False
                    
                    if not baseEmitterMode:
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_LBL, "MODE:")
                        self.writeValueCell(wsEmitters, wsEmittersRow, constant.COMP_XL_COL_MODE_VAL, comparisonEmitterMode.get_name())
                        self.writeLabelCell(wsEmitters, wsEmittersRow, constant.BASE_XL_COL_MODE_LBL, constant.XL_MISSING_TEXT)
                        wsEmittersRow += 1
                    else:
                        
                        for comparisonModeAttribute in comparisonEmitterMode.get_attributes():
                            baseModeAttribute = self.findAttribute(comparisonModeAttribute.get_name(), baseEmitterMode)
                            
                            if not baseModeAttribute:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, cElnot, baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                    
                                self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_ATTRIB_LBL, comparisonModeAttribute.get_name())
                                self.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_MODE_ATTRIB_VAL, comparisonModeAttribute.get_value())
                                self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_MODE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                
                                wsModesRow += 1
                                    
                        for comparisonGenerator in comparisonEmitterMode.get_generators():
                            baseGenerator = self.findGenerator(comparisonGenerator.get_generator_number(), baseEmitterMode)
                        
                            generatorWritten = False
                            
                            if not baseGenerator:
                                
                                if modeWritten == False:
                                    self.writeMode(wsModes, wsModesRow, bElnot, comparisonEmitter.get_elnot(), baseEmitterMode.get_name(), comparisonEmitterMode.get_name(), modeWritten)
                                    wsModesRow += 1
                                
                                self.writeLabelCell(wsModes, wsModesRow, constant.COMP_XL_COL_GENERATOR_LBL, "GENERATOR:")
                                self.writeValueCell(wsModes, wsModesRow, constant.COMP_XL_COL_GENERATOR_VAL, comparisonGenerator.get_generator_number())
                                self.writeLabelCell(wsModes, wsModesRow, constant.BASE_XL_COL_GENERATOR_LBL, constant.XL_MISSING_TEXT)
                                
                                wsModesRow += 1
                            else:
                                
                                for comparisonGeneratorAttribute in comparisonGenerator.get_attributes():
                                    baseGeneratorAttribute = self.findAttribute(comparisonGeneratorAttribute.get_name(), baseGenerator)
                                    
                                    if not baseGeneratorAttribute:
                                        
                                        if generatorsWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                                
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_LBL, comparisonGeneratorAttribute.get_name())
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_GENERATOR_ATTRIB_VAL, comparisonGeneratorAttribute.get_value())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_GENERATOR_ATTRIB_LBL, constant.XL_MISSING_TEXT)
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
                                        
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "PRI_SEQUENCE:")
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, comparisonPRISequence.get_ordinal_pos())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        
                                        wsGeneratorsRow += 1
                                    else:
                                        
                                        for cPRISeqAttribute in comparisonPRISequence.get_attributes(): 
                                            bPRISeqAttribute = self.findAttribute(cPRISeqAttribute.get_name(), basePRISequence)
                                            
                                            if not bPRISeqAttribute:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                    
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cPRISeqAttribute.get_name())
                                                self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cPRISeqAttribute.get_value())
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                wsPRISequencesRow += 1

                                        for cPRISegment in comparisonPRISequence.get_segments():
                                            bPRISegment = self.findSegmentBySegmentNumber(cPRISegment.get_segment_number(), basePRISequence)
                                            priSegmentWritten = False
                                            
                                            if not bPRISegment:
                                                
                                                if priSequenceWritten == False:
                                                    self.writePRISequence(wsPRISequences, wsPRISequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, basePRISequence, comparisonPRISequence, priSequenceWritten)
                                                    wsPRISequencesRow += 1
                                                
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "PRI_SEGMENT:")
                                                self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cPRISegment.get_segment_number())
                                                self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
                                                
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
                                                            
                                                        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cSegmentAttribute.get_name())
                                                        self.writeValueCell(wsPRISequences, wsPRISequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cSegmentAttribute.get_value())
                                                        self.writeLabelCell(wsPRISequences, wsPRISequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                        wsPRISequencesRow += 1

                                for comparisonFREQSequence in comparisonGenerator.get_freq_sequences():
                                    baseFREQSequence = self.findFREQSequenceByOrdinalPos(comparisonFREQSequence.get_ordinal_pos(), baseGenerator)
                                    freqSequenceWritten = False
                                    
                                    if not baseFREQSequence:
                                        
                                        if generatorWritten == False:
                                            self.writeGenerator(wsGenerators, wsGeneratorsRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, generatorWritten)
                                            wsGeneratorsRow += 1
                                        
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_LBL, "FREQ_SEQUENCE:")
                                        self.writeValueCell(wsGenerators, wsGeneratorsRow, constant.COMP_XL_COL_PF_SEQUENCE_VAL, comparisonFREQSequence.get_ordinal_pos())
                                        self.writeLabelCell(wsGenerators, wsGeneratorsRow, constant.BASE_XL_COL_PF_SEQUENCE_LBL, constant.XL_MISSING_TEXT)
                                        wsGeneratorsRow += 1

                                    else:
                                        
                                        for cFREQSeqAttribute in comparisonFREQSequence.get_attributes():
                                            bFREQSeqAttribute = self.findAttribute(cFREQSeqAttribute.get_name(), baseFREQSequence)
                                            
                                            if not bFREQSeqAttribute:
                                                
                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1
                                                    
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_LBL, cFREQSeqAttribute.get_name())
                                                self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEQUENCE_ATTRIB_VAL, cFREQSeqAttribute.get_value())
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEQUENCE_ATTRIB_LBL, constant.XL_MISSING_TEXT)
                                                wsFREQSequencesRow += 1

                                        for cFREQSegment in comparisonFREQSequence.get_segments():
                                            bFREQSegment = self.findSegmentBySegmentNumber(cFREQSegment.get_segment_number(), baseFREQSequence)
                                            freqSegmentWritten = False
                                            
                                            if not bFREQSegment:

                                                if freqSequenceWritten == False:
                                                    self.writeFREQSequence(wsFREQSequences, wsFREQSequencesRow, bElnot, comparisonEmitter, baseEmitterMode, comparisonEmitterMode, baseGenerator, comparisonGenerator, baseFREQSequence, comparisonFREQSequence, freqSequenceWritten)
                                                    wsFREQSequencesRow += 1

                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_LBL, "FREQ_SEGMENT:")
                                                self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_VAL, cFREQSegment.get_segment_number())
                                                self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_LBL, constant.XL_MISSING_TEXT)
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
                                                            
                                                        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_LBL, cSegmentAttribute.get_name())
                                                        self.writeValueCell(wsFREQSequences, wsFREQSequencesRow, constant.COMP_XL_COL_PF_SEGMENT_ATTRIB_VAL, cSegmentAttribute.get_value())
                                                        self.writeLabelCell(wsFREQSequences, wsFREQSequencesRow, constant.BASE_XL_COL_PF_SEGMENT_ATTRIB_LBL, constant.XL_MISSING_TEXT)
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

            
# =============================================================================
#             wsEmittersUsedRangeRows = wsEmitters.api.UsedRange.Rows.Count + 1
#             wsAutofitRange = "A{}".format(wsEmittersUsedRangeRows)
#             wsEmitters.range("A1",wsAutofitRange).autofit()
#             
#             wsModesUsedRangeRows = wsModes.api.UsedRange.Rows.Count + 1
#             wsAutofitRange = "A{}".format(wsModesUsedRangeRows)
#             wsModes.range("A1",wsAutofitRange).autofit()
#             
#             wsGeneratorsUsedRangeRows = wsGenerators.api.UsedRange.Rows.Count + 1
#             wsAutofitRange = "A{}".format(wsGeneratorsUsedRangeRows)
#             wsGenerators.range("A1",wsAutofitRange).autofit()
# 
#             wsPRISequencesUsedRangeRows = wsPRISequences.api.UsedRange.Rows.Count + 1
#             wsAutofitRange = "A{}".format(wsPRISequencesUsedRangeRows)
#             wsPRISequences.range("A1",wsAutofitRange).autofit()
# 
#             wsFREQSequencesUsedRangeRows = wsFREQSequences.api.UsedRange.Rows.Count + 1
#             wsAutofitRange = "A{}".format(wsFREQSequencesUsedRangeRows)
#             wsFREQSequences.range("A1",wsAutofitRange).autofit()
#                                 
# =============================================================================
                                
                                
     
    def compareTheFiles(self):
    
        self.getShortFileNames()
        
        print("parsing Base File ...")
        self.parseBaseFile()
        
        print("parsing Comparison File ...")
        self.parseComparisonFile()
    
        wb = xw.Book()
     
       
        self.compareEMTFiles(wb, self.base_emitters, self.comparison_emitters, constant.COMPARISON_ARRAY, constant.BASE_ARRAY)
    
        
        #self.compareEMTFiles(wf, self.comparison_emitters, self.base_emitters, constant.BASE_ARRAY, self.comparisonFileName, self.baseFileName)
    
       
        wb.save(r'EMT_Differences')
        wb.close()
    
               
    #if __name__ == '__main__':
     #   compareTheFiles()
    
                    
                
                
                       
                   

       