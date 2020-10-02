# -*- coding: utf-8 -*-
"""
    
    print_utility.py
   ----------------------
   
   This module contains functions for printing to a cell using xlwings.

"""

__author__ = "Steven M. Satterfield"

import constant
from emitter import Emitter
from attribute import Attribute
from emitter_mode import EmitterMode
from generator import Generator
from sequence import Sequence
from segment import Segment
import pandas as pd


def writeDFFREQSequences(base_emitters, bfTitle, cfTitle):
    
    dfSubCols = ['ELNOT', 'MODE', 'GENERATOR', 'FREQ_SEQUENCE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'SEGMENT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'ELNOT', 'MODE', 'GENERATOR', 'FREQ_SEQUENCE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'SEGMENT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE']
    dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','','','','','', cfTitle,'','','','','','','',''], dfSubCols))
    
    d = []
    for emitter in base_emitters:
        if emitter.get_hasDifferences() == True:
            bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
            cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
            
            for mode in emitter.get_modes():
                if mode.get_hasDifferences() == True:
                    bMode = constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name()
                    cMode = constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name()

                    
                    for generator in mode.get_generators():
                        if generator.get_hasDifferences() == True:
                            bGenerator = constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number()
                            cGenerator = constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number()
                            
                            for sequence in generator.get_freq_sequences():
                                if sequence.get_hasDifferences() == True:
                                    s = [bElnot,
                                         bMode,
                                         bGenerator,
                                         constant.XL_MISSING_TEXT if sequence.get_bfile() == False else sequence.get_ordinal_pos(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         cElnot,
                                         cMode,
                                         cGenerator,
                                         constant.XL_MISSING_TEXT if sequence.get_cfile() == False else sequence.get_ordinal_pos(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER]
                                    d.append(s)
    
                                    for attribute in sequence.get_attributes():
                                        if attribute.get_hasDifferences() == True:
                                            a = [constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name(),
                                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name(),
                                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER]
                                            d.append(a)

                                    for segment in sequence.get_segments():
                                        if segment.get_hasDifferences() == True:
                                            sg = [constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if segment.get_bfile() == False else segment.get_segment_number(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if segment.get_cfile() == False else segment.get_segment_number(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER]
                                            d.append(sg)


                                            for s_attribute in segment.get_attributes():
                                                if s_attribute.get_hasDifferences() == True:
                                                    sga = [constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.XL_MISSING_TEXT if s_attribute.get_bfile() == False else s_attribute.get_name(),
                                                           constant.XL_MISSING_TEXT if s_attribute.get_bfile() == False else s_attribute.get_value(),
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.XL_MISSING_TEXT if s_attribute.get_cfile() == False else s_attribute.get_name(),
                                                           constant.XL_MISSING_TEXT if s_attribute.get_cfile() == False else s_attribute.get_cvalue()]
                                                    d.append(sga)

                                    divider = [constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER]
                                    d.append(divider)
                                        
    
    dfSequences = pd.DataFrame(d, columns = dfCols)
    
    return dfSequences
       

def writeDFPRISequences(base_emitters, bfTitle, cfTitle):
    
    dfSubCols = ['ELNOT', 'MODE', 'GENERATOR', 'PRI_SEQUENCE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'SEGMENT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'ELNOT', 'MODE', 'GENERATOR', 'PRI_SEQUENCE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'SEGMENT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE']
    dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','','','','','', cfTitle,'','','','','','','',''], dfSubCols))
    
    d = []
    for emitter in base_emitters:
        if emitter.get_hasDifferences() == True:
            bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
            cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
            
            for mode in emitter.get_modes():
                if mode.get_hasDifferences() == True:
                    bMode = constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name()
                    cMode = constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name()

                    
                    for generator in mode.get_generators():
                        if generator.get_hasDifferences() == True:
                            bGenerator = constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number()
                            cGenerator = constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number()
                            
                            for sequence in generator.get_pri_sequences():
                                if sequence.get_hasDifferences() == True:
                                    s = [bElnot,
                                         bMode,
                                         bGenerator,
                                         constant.XL_MISSING_TEXT if sequence.get_bfile() == False else sequence.get_ordinal_pos(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         cElnot,
                                         cMode,
                                         cGenerator,
                                         constant.XL_MISSING_TEXT if sequence.get_cfile() == False else sequence.get_ordinal_pos(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER]
                                    d.append(s)
    
                                    for attribute in sequence.get_attributes():
                                        if attribute.get_hasDifferences() == True:
                                            a = [constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name(),
                                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name(),
                                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER]
                                            d.append(a)

                                    for segment in sequence.get_segments():
                                        if segment.get_hasDifferences() == True:
                                            sg = [constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if segment.get_bfile() == False else segment.get_segment_number(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER,
                                                 constant.XL_MISSING_TEXT if segment.get_cfile() == False else segment.get_segment_number(),
                                                 constant.PLACEHOLDER,
                                                 constant.PLACEHOLDER]
                                            d.append(sg)


                                            for s_attribute in segment.get_attributes():
                                                if s_attribute.get_hasDifferences() == True:
                                                    sga = [constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.XL_MISSING_TEXT if s_attribute.get_bfile() == False else s_attribute.get_name(),
                                                           constant.XL_MISSING_TEXT if s_attribute.get_bfile() == False else s_attribute.get_value(),
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.PLACEHOLDER,
                                                           constant.XL_MISSING_TEXT if s_attribute.get_cfile() == False else s_attribute.get_name(),
                                                           constant.XL_MISSING_TEXT if s_attribute.get_cfile() == False else s_attribute.get_cvalue()]
                                                    d.append(sga)

                                    divider = [constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER,
                                               constant.PLACEHOLDER]
                                    d.append(divider)
                                        
    
    dfSequences = pd.DataFrame(d, columns = dfCols)
    
    return dfSequences

        
def writeDFGenerators(base_emitters, bfTitle, cfTitle):
    
    dfSubCols = ['ELNOT', 'MODE', 'GENERATOR', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'PRI_SEQUENCE', 'FREQ_SEQUENCE', 'ELNOT', 'MODE', 'GENERATOR', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'PRI_SEQUENCE', 'FREQ_SEQUENCE']
    dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','','','', cfTitle,'','','','','',''], dfSubCols))
    
    d = []
    for emitter in base_emitters:
        if emitter.get_hasDifferences() == True:
            bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
            cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
            
            for mode in emitter.get_modes():
                if mode.get_hasDifferences() == True:
                    bMode = constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name()
                    cMode = constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name()

                    
                    for generator in mode.get_generators():
                        if generator.get_hasDifferences() == True:
                            g = [bElnot,
                                 bMode,
                                 constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number(),
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 cElnot,
                                 cMode,
                                 constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number(),
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER]
                            d.append(g)
    
                            for attribute in generator.get_attributes():
                                if attribute.get_hasDifferences() == True:
                                    a = [constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name(),
                                         constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER,
                                         constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name(),
                                         constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue(),
                                         constant.PLACEHOLDER,
                                         constant.PLACEHOLDER]
                                    d.append(a)

                            for sequence in generator.get_pri_sequences():
                                if sequence.get_bfile() == False or sequence.get_cfile() == False:
                                    ps = [constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.XL_MISSING_TEXT if sequence.get_bfile() == False else sequence.get_ordinal_pos(),
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.XL_MISSING_TEXT if sequence.get_cfile() == False else sequence.get_ordinal_pos(),
                                          constant.PLACEHOLDER]
                                    d.append(ps)

                            for sequence in generator.get_freq_sequences():
                                if sequence.get_bfile() == False or sequence.get_cfile() == False:
                                    fs = [constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.XL_MISSING_TEXT if sequence.get_bfile() == False else sequence.get_ordinal_pos(),
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.PLACEHOLDER,
                                          constant.XL_MISSING_TEXT if sequence.get_bfile() == False else sequence.get_ordinal_pos(),
                                          constant.PLACEHOLDER]
                                    d.append(fs)

                    divider = [constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER]
                    d.append(divider)
                    
    
    dfModes = pd.DataFrame(d, columns = dfCols)
    
    return dfModes
        
        
def writeDFModes(base_emitters, bfTitle, cfTitle):
    
    dfSubCols = ['ELNOT', 'MODE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'GENERATOR', 'ELNOT', 'MODE', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'GENERATOR']
    dfCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','','', cfTitle,'','','',''], dfSubCols))
    
    d = []
    for emitter in base_emitters:
        if emitter.get_hasDifferences() == True:
            bElnot = constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot()
            cElnot = constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot()
            
            for mode in emitter.get_modes():
                if mode.get_hasDifferences() == True:
                    m = [bElnot,
                         constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name(),
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         cElnot,
                         constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name(),
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER]
                    d.append(m)

                    for attribute in mode.get_attributes():
                        if attribute.get_hasDifferences() == True:
                            a = [constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name(),
                                 constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value(),
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name(),
                                 constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue(),
                                 constant.PLACEHOLDER]
                            d.append(a)

                    for generator in mode.get_generators():
                        if generator.get_bfile() == False or generator.get_cfile() == False:
                            g = [constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.XL_MISSING_TEXT if generator.get_bfile() == False else generator.get_generator_number(),
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.PLACEHOLDER,
                                 constant.XL_MISSING_TEXT if generator.get_cfile() == False else generator.get_generator_number()]
                            d.append(g)

                    divider = [constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER,
                               constant.PLACEHOLDER]
                    d.append(divider)
                    
    
    dfModes = pd.DataFrame(d, columns = dfCols)
    
    return dfModes
        

def writeDFEmitters(base_emitters, bfTitle, cfTitle):
    
    dfEmittersSubCols = ['ELNOT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'MODE', 'ELNOT', 'ATTRIBUTE NAME', 'ATTRIBUTE VALUE', 'MODE']
    dfEmittersCols = pd.MultiIndex.from_tuples(zip([bfTitle, '','','', cfTitle,'','',''], dfEmittersSubCols))
    
    d = []
    for emitter in base_emitters:
        
        if emitter.get_hasDifferences() == True:
            e = [constant.XL_MISSING_TEXT if emitter.get_bfile() == False else emitter.get_elnot(),
                 constant.PLACEHOLDER,
                 constant.PLACEHOLDER,
                 constant.PLACEHOLDER,
                 constant.XL_MISSING_TEXT if emitter.get_cfile() == False else emitter.get_elnot(),
                 constant.PLACEHOLDER,
                 constant.PLACEHOLDER,
                 constant.PLACEHOLDER]
            d.append(e)

            for attribute in emitter.get_attributes():
                if attribute.get_hasDifferences() == True:
                    r = [constant.PLACEHOLDER,
                         constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_name(),
                         constant.XL_MISSING_TEXT if attribute.get_bfile() == False else attribute.get_value(),
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_name(),
                         constant.XL_MISSING_TEXT if attribute.get_cfile() == False else attribute.get_cvalue(),
                         constant.PLACEHOLDER]
                    d.append(r)

            for mode in emitter.get_modes():
                if mode.get_bfile() == False or mode.get_cfile() == False:
                    m = [constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.XL_MISSING_TEXT if mode.get_bfile() == False else mode.get_name(),
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.PLACEHOLDER,
                         constant.XL_MISSING_TEXT if mode.get_cfile() == False else mode.get_name()]
                    d.append(m)
                    
            divider = [constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER,
                       constant.PLACEHOLDER]
            d.append(divider)
                    
    dfEmitters = pd.DataFrame(d, columns = dfEmittersCols)
    
    return dfEmitters


def dictsToPandasDFs(base_emitters, bfDisplay, cfDisplay):
    
    bfTitle = "Base File: {}".format(bfDisplay)
    cfTitle = "Comparison File: {}".format(cfDisplay)
    

    dfsWB = []
    dfEmitters = writeDFEmitters(base_emitters, bfTitle, cfTitle)
    dfModes = writeDFModes(base_emitters, bfTitle, cfTitle)
    dfGenerators = writeDFGenerators(base_emitters, bfTitle, cfTitle)
    dfPRISequences = writeDFPRISequences(base_emitters, bfTitle, cfTitle)
    dfFREQSequences = writeDFFREQSequences(base_emitters, bfTitle, cfTitle)
    
    dfsWB.append(dfEmitters)
    dfsWB.append(dfModes)
    dfsWB.append(dfGenerators)
    dfsWB.append(dfPRISequences)
    dfsWB.append(dfFREQSequences)
    
    return dfsWB


