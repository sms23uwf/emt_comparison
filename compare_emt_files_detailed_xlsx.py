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
import list_utility
from itertools import filterfalse

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
            
            
    def cleanTheList(self):
        self.base_emitters = list(filterfalse(list_utility.filtertrue, self.base_emitters))
        
        for emitter in self.base_emitters:
            emitter.clean_attributes()
            emitter.clean_modes()
                
        
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
                                                bPRISegment.print_pri_segment(wsPRISequences)
                                                
                                for baseSequence in baseGenerator.get_freq_sequences():
                                    if baseSequence.get_hasDifferences() == True:
                                        baseSequence.print_freq_sequence(wsFREQSequences, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number())
                                        
                                        for bSegment in baseSequence.get_segments():
                                            if bSegment.get_hasDifferences() == True:
                                                bSegment.print_freq_segment(wsFREQSequences)
                                                            

            
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
    
    
    def listToDict(self):
        emitter_holder = []
        mode_holder = []
        
        for emitter in self.base_emitters:
            emitter_holder.append(emitter)
                
        for dictCandidateEmitter in emitter_holder:
            dictCandidateEmitter.modes_to_dict()
            dictCandidateEmitter.attributes_to_dict()
            
            self.output_emitters.append(dictCandidateEmitter.__dict__)
            
            
    def writeDFGenerators(self, bfTitle, cfTitle):
        
        dfSubCols = ['ELNOT', 'MODE', 'GENERATOR', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'PRI_SEQUENCE', 'FREQ_SEQUENCE', 'ELNOT', 'MODE', 'GENERATOR', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'PRI_SEQUENCE', 'FREQ_SEQUENCE']
        dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','','','', cfTitle,'','','','','',''], dfSubCols))
        
        d = []
        for emitter in self.base_emitters:
            if emitter.get_hasDifferences() == True:
                bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
                cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
                
                for mode in emitter.get_modes():
                    if mode.get_hasDifferences() == True:
                        bMode = constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name()
                        cMode = constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name()

                        
                        for generator in mode.get_generators():
                            if generator.get_hasDifferences() == True:
                                g = []
                                g.append(bElnot)
                                g.append(bMode)
                                g.append(constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number())
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append('')
                                
                                g.append(cElnot)
                                g.append(cMode)
                                g.append(constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number())
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append('')
                                d.append(g)
        
                                for attribute in generator.get_attributes():
                                    if attribute.get_hasDifferences() == True:
                                        a = []
                                        a.append('')
                                        a.append('')
                                        a.append('')
                                        a.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name())
                                        a.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value())
                                        a.append('')
                                        a.append('')
        
                                        a.append('')
                                        a.append('')
                                        a.append('')
                                        a.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name())
                                        a.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue())
                                        a.append('')
                                        a.append('')
                                        
                                        d.append(a)

                                for sequence in generator.get_pri_sequences():
                                    if sequence.get_bfile() == False or sequence.get_cfile() == False:
                                        ps = []
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append(constant.XL_MISSING_TEXT if generator.get_bfile() == False else sequence.get_ordinal_pos())
                                        ps.append('')
                                        
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append('')
                                        ps.append(constant.XL_MISSING_TEXT if generator.get_cfile() == False else sequence.get_ordinal_pos())
                                        ps.append('')
                
                                        d.append(ps)

                                for sequence in generator.get_freq_sequences():
                                    if sequence.get_bfile() == False or sequence.get_cfile() == False:
                                        fs = []
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append(constant.XL_MISSING_TEXT if generator.get_bfile() == False else sequence.get_ordinal_pos())
                                        fs.append('')
                                        
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append('')
                                        fs.append(constant.XL_MISSING_TEXT if generator.get_cfile() == False else sequence.get_ordinal_pos())
                                        fs.append('')
                
                                        d.append(fs)

                        divider = []
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        d.append(divider)
                        
        
        dfModes = pd.DataFrame(d, columns = dfCols)
        
        return dfModes
            
            
    def writeDFModes(self, bfTitle, cfTitle):
        
        dfSubCols = ['ELNOT', 'MODE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'GENERATOR', 'ELNOT', 'MODE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'GENERATOR']
        dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','', cfTitle,'','','',''], dfSubCols))
        
        d = []
        for emitter in self.base_emitters:
            if emitter.get_hasDifferences() == True:
                bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
                cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
                
                for mode in emitter.get_modes():
                    if mode.get_hasDifferences() == True:
                        m = []
                        m.append(bElnot)
                        m.append(constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name())
                        m.append('')
                        m.append('')
                        m.append('')
                        
                        m.append(cElnot)
                        m.append(constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name())
                        m.append('')
                        m.append('')
                        m.append('')
                        d.append(m)

                        for attribute in mode.get_attributes():
                            if attribute.get_hasDifferences() == True:
                                a = []
                                a.append('')
                                a.append('')
                                a.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name())
                                a.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value())
                                a.append('')

                                a.append('')
                                a.append('')
                                a.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name())
                                a.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue())
                                a.append('')
                                
                                d.append(a)

                        for generator in mode.get_generators():
                            if generator.get_bfile() == False or generator.get_cfile() == False:
                                g = []
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append(constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number())
                                
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append('')
                                g.append(constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number())
        
                                d.append(g)

                        divider = []
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        divider.append('')
                        d.append(divider)
                        
        
        dfModes = pd.DataFrame(d, columns = dfCols)
        
        return dfModes
            
    
    def writeDFEmitters(self, bfTitle, cfTitle):
        
        dfEmittersSubCols = ['ELNOT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'MODE', 'ELNOT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'MODE']
        dfEmittersCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','', cfTitle,'','',''], dfEmittersSubCols))
        
        d = []
        for emitter in self.base_emitters:
            
            if emitter.get_hasDifferences() == True:
                e = []
                e.append(constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot())
                e.append('')
                e.append('')
                e.append('')
                e.append(constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot())
                e.append('')
                e.append('')
                e.append('')
                d.append(e)

                for attribute in emitter.get_attributes():
                    if attribute.get_hasDifferences() == True:
                        r = []
                        r.append('')
                        r.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name())
                        r.append(constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value())
                        r.append('')
                        r.append('')
                        r.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name())
                        r.append(constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue())
                        r.append('')
                        
                        d.append(r)

                for mode in emitter.get_modes():
                    if mode.get_bfile() == False or mode.get_cfile() == False:
                        m = []
                        m.append('')
                        m.append('')
                        m.append('')
                        m.append(constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name())
                        m.append('')
                        m.append('')
                        m.append('')
                        m.append(constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name())

                        d.append(m)
                        
                divider = []
                divider.append('')
                divider.append('')
                divider.append('')
                divider.append('')
                divider.append('')
                divider.append('')
                divider.append('')
                divider.append('')
                d.append(divider)
                        
        dfEmitters = pd.DataFrame(d, columns = dfEmittersCols)
        
        return dfEmitters
    
    
    def dictsToPandasDFs(self):
        
        #df = pd.DataFrame.from_records(self.output_emitters)
        bfTitle = "Base File: {}".format(self.bfDisplay)
        cfTitle = "Comparison File: {}".format(self.cfDisplay)
        

        dfsWB = []
        dfEmitters = self.writeDFEmitters(bfTitle, cfTitle)
        dfModes = self.writeDFModes(bfTitle, cfTitle)
        dfGenerators = self.writeDFGenerators(bfTitle, cfTitle)
        
        dfPRISequences = pd.DataFrame()
        dfFREQSequences = pd.DataFrame()
        
        dfsWB.append(dfEmitters)
        dfsWB.append(dfModes)
        dfsWB.append(dfGenerators)
        dfsWB.append(dfPRISequences)
        dfsWB.append(dfFREQSequences)
        
        return dfsWB
    
    
    def compareTheFiles(self):
    
        print("starting file comparison.")
        
        self.getShortFileNames()
        
        print("parsing Base File ...")
        self.parseBaseFile()
        
        print("parsing Comparison File ...")
        self.parseComparisonFile()
    
        #print("cleaning the list, leaving only records with a difference ...")
        #self.cleanTheList()
        
        #print("converting objects to dictionaries for dataframe...")
        #self.listToDict()
        
        dfsWB = self.dictsToPandasDFs()
        writer = pd.ExcelWriter('EMT_Differences.xlsx', engine='xlsxwriter')
        
        dfEmitters = dfsWB[0]
        dfModes = dfsWB[1]
        dfGenerators = dfsWB[2]
        dfPRISequences = dfsWB[3]
        dfFREQSequences = dfsWB[4]
        
        
        #columns = ['ELNOT','MODE','GENERATOR','SEQUENCE','SEGMENT']
        #xwApp = xw.App(visible=False)
        
        #wb = xw.Book()
     
        print("outputting excel display of differences ...")
        dfEmitters.to_excel(writer, sheet_name='Emitters')
        dfModes.to_excel(writer, sheet_name='Modes')
        dfGenerators.to_excel(writer, sheet_name='Generators')
        writer.save()
        
        #self.displayDifferences(wb)
    
       
        #wb.save(r'EMT_Differences')
        #wb.close()
        #xwApp.quit()
    
        print("finished!")
               
    #if __name__ == '__main__':
     #   compareTheFiles()
       