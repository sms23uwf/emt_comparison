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



writeFile = 'diff.txt'
baseFileName = 'SPECTRE_OUTPUT_07032020.emt'
comparisonFileName = 'SPECTRE_OUTPUT_08032020.emt'
  
base_emitters = []
comparison_emitters = []

currentEntity = constant.EMITTER

def lineAsKeyValue(line):
    line_kv = line.split(constant.VALUE_SEPARATOR)
    return line_kv

    
def parseFile(fName, emitter_collection):
    emitter = Emitter()
    emitter_mode = EmitterMode()
    generator = Generator()
    pri_sequence = Pri_Sequence()
    freq_sequence = Freq_Sequence()
    pri_segment = Pri_Segment
    freq_segment = Freq_Segment
        
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
                    emitter_collection.append(emitter)
                    
                emitter = Emitter()
                passNumber += 1
                modePass = 0
                generatorPass = 0
                priSeqPass = 0
                freqSeqPass = 0
                priSegmentPass = 0
                freqSegmentPass = 0
                
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
                
            elif line.strip() == constant.PRI_SEQUENCE:
                currentEntity = constant.PRI_SEQUENCE
                if priSeqPass > 0:
                    generator.add_pri_sequence(pri_sequence)
                    
                pri_sequence = Pri_Sequence()
                pri_sequence.set_ordinal_pos(priSeqPass)
                priSeqPass += 1
                priSegmentPass = 0
                
            elif line.strip() == constant.PRI_SEGMENT:
                currentEntity = constant.PRI_SEGMENT
                if priSegmentPass > 0:
                    pri_sequence.add_segment(pri_segment)
                    
                pri_segment = Pri_Segment()
                priSegmentPass += 1

            elif line.strip() == constant.FREQ_SEQUENCE:
                currentEntity = constant.FREQ_SEQUENCE
                if freqSeqPass > 0:
                    generator.add_freq_sequence(freq_sequence)
                    
                freq_sequence = Freq_Sequence()
                freq_sequence.set_ordinal_pos(freqSeqPass)
                freqSeqPass += 1
                freqSegmentPass = 0
                
            elif line.strip() == constant.FREQ_SEGMENT:
                currentEntity = constant.FREQ_SEGMENT
                if freqSegmentPass > 0:
                    freq_sequence.add_segment(freq_segment)
                    
                freq_segment = Freq_Segment()
                freqSegmentPass += 1
                
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
                            
                    elif currentEntity == constant.PRI_SEQUENCE:
                        
                        if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                            pri_sequence.set_number_of_segments(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            pri_sequence.add_attribute(attrib)

                    elif currentEntity == constant.PRI_SEGMENT:
                        
                        if line_key.strip() == constant.PRI_SEGMENT_NUMBER:
                            pri_segment.set_segment_number(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            pri_segment.add_attribute(attrib)

                    elif currentEntity == constant.FREQ_SEQUENCE:
                        
                        if line_key.strip() == constant.NUMBER_OF_SEGMENTS:
                            freq_sequence.set_number_of_segments(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            freq_sequence.add_attribute(attrib)

                    elif currentEntity == constant.FREQ_SEGMENT:
                        
                        if line_key.strip() == constant.FREQ_SEGMENT_NUMBER:
                            freq_segment.set_segment_number(line_value)
                        else:
                            attrib = Attribute()
                            attrib.set_name(line_key)
                            attrib.set_value(line_value)
                            freq_segment.add_attribute(attrib)
                            
        else:
            emitter_collection.append(emitter)


def parseBaseFile():
    parseFile(baseFileName, base_emitters)
        
    
def parseComparisonFile():
    parseFile(comparisonFileName, comparison_emitters)
    
                    
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
                return comparisonGenerator
          
        return []


def findPRISequenceByOrdinalPos(ordinalPos, comparisonGenerator):
    for cPRISequence in comparisonGenerator.get_pri_sequences():
        if cPRISequence.get_ordinal_pos() == ordinalPos:
            return cPRISequence
        
    return []


def findFREQSequenceByOrdinalPos(ordinalPos, comparisonGenerator):
    for cFREQSequence in comparisonGenerator.get_freq_sequences():
        if cFREQSequence.get_ordinal_pos() == ordinalPos:
            return cFREQSequence 
        
    return []
    

def findSegmentBySegmentNumber(segmentNumber, comparisonSequence):
    for cSegment in comparisonSequence.get_segments():
        if cSegment.get_segment_number() == segmentNumber:
            return cSegment
        
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
                            wf.write("{} mode {} contains generator: {} that is missing from the mode in {}.\n".format(bFileName, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), cFileName))
                        else:
                            for baseGeneratorAttribute in baseGenerator.get_attributes():
                                comparisonGeneratorAttribute = findAttribute(baseGeneratorAttribute.get_name(), comparisonGenerator)
                                
                                if not comparisonGeneratorAttribute:
                                    wf.write("{} emitter:{}.mode:{}.generator:{} contains attribute:{} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseGeneratorAttribute.get_name(), cFileName))
                                else:
                                    if comparisonGeneratorAttribute.get_value() != baseGeneratorAttribute.get_value():
                                        wf.write("{} emitter:{}.mode:{}.generator:{} contains attribute:{} with value:{} that is different than the value:{} in the same path attribute in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseGeneratorAttribute.get_name(), baseGeneratorAttribute.get_value(), comparisonGeneratorAttribute.get_value(), cFileName))
                        
                                    bgPRISequences = baseGenerator.get_pri_sequences()
                                    bgFREQSequences = baseGenerator.get_freq_sequences()
                                    cpPRISequences = comparisonGenerator.get_pri_sequences()
                                    cpFREQSequences = comparisonGenerator.get_freq_sequences()
                                    
                                    if len(bgPRISequences) != len(cpPRISequences):
                                        wf.write("{} emitter({}).mode({}).generator({}) contains {} PRI Sequences - but the same path generator in {} contains {} PRI Sequences.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), len(bgPRISequences), cFileName, len(cpPRISequences)))
            
                                    if len(bgFREQSequences) != len(cpFREQSequences):
                                        wf.write("{} emitter({}).mode({}).generator({}) contains {} FREQ Sequences - but the same path generator in {} contains {} FREQ Sequences.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), len(bgFREQSequences), cFileName, len(cpFREQSequences)))
                                    
                                    for basePRISequence in baseGenerator.get_pri_sequences():
                                        comparisonPRISequence = findPRISequenceByOrdinalPos(basePRISequence.get_ordinal_pos(), comparisonGenerator)
                                        
                                        if not comparisonPRISequence:
                                            wf.write("{} emitter({}).mode({}).generator({}) contains PRISequence in ordinal position {} that is missing from the same path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), cFileName))
                                        else:
                                            for bPRISeqAttribute in basePRISequence.get_attributes():
                                                cPRISeqAttribute = findAttribute(bPRISeqAttribute.get_name(), comparisonPRISequence)
                                                
                                                if not cPRISeqAttribute:
                                                    wf.write("{} emitter({}).mode({}).generator({}).PRISequence({}) contains attribute: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), bPRISeqAttribute.get_name(), cFileName))
                                                else:
                                                    if cPRISeqAttribute.get_value() != bPRISeqAttribute.get_value():
                                                        wf.write("{} emitter({}).mode({}).generator({}).PRISequence({}) contains attribute:{} with value:{} that is different than the value:{} in the same path attribute in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), bPRISeqAttribute.get_name(), bPRISeqAttribute.get_value(), cPRISeqAttribute.get_value(), cFileName))
            
                                            for bPRISegment in basePRISequence.get_segments():
                                                cPRISegment = findSegmentBySegmentNumber(bPRISegment.get_segment_number(), comparisonPRISequence)
                                                
                                                if not cPRISegment:
                                                    wf.write("{} emitter({}).mode({}).generator({}).PRISequence({}) contains PRI Segment number: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), bPRISegment.get_segment_number(), cFileName))
                                                else:
                                                    for bSegmentAttribute in bPRISegment.get_attributes():
                                                        cSegmentAttribute = findAttribute(bSegmentAttribute.get_name(), cPRISegment)
                                                        
                                                        if not cSegmentAttribute:
                                                            wf.write("{} emitter({}).mode({}).generator({}).PRISequence({}).PRISegment({}) contains attribute: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), bPRISegment.get_segment_number(), bSegmentAttribute.get_name(), cFileName))
                                                        else:
                                                            if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                                wf.write("{} emitter({}).mode({}).generator({}).PRISequence({}).PRISegment({}) contains attribute: {} with value:{} that is different than the value:{} in the same path attribute in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), basePRISequence.get_ordinal_pos(), bPRISegment.get_segment_number(), bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), cSegmentAttribute.get_value(), cFileName))
                        
                                    for baseFREQSequence in baseGenerator.get_freq_sequences():
                                        comparisonFREQSequence = findFREQSequenceByOrdinalPos(baseFREQSequence.get_ordinal_pos(), comparisonGenerator)
                                        
                                        if not comparisonFREQSequence:
                                            wf.write("{} emitter({}).mode({}).generator({}) contains FREQSequence in ordinal position {} that is missing from the same path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), cFileName))
                                        else:
                                            for bFREQSeqAttribute in baseFREQSequence.get_attributes():
                                                cFREQSeqAttribute = findAttribute(bFREQSeqAttribute.get_name(), comparisonFREQSequence)
                                                
                                                if not cFREQSeqAttribute:
                                                    wf.write("{} emitter({}).mode({}).generator({}).FREQSequence({}) contains attribute: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), bFREQSeqAttribute.get_name(), cFileName))
                                                else:
                                                    if cFREQSeqAttribute.get_value() != bFREQSeqAttribute.get_value():
                                                        wf.write("{} emitter({}).mode({}).generator({}).FREQSequence({}) contains attribute:{} with value:{} that is different than the value:{} in the same path attribute in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), bFREQSeqAttribute.get_name(), bFREQSeqAttribute.get_value(), cFREQSeqAttribute.get_value(), cFileName))
            
                                            for bFREQSegment in baseFREQSequence.get_segments():
                                                cFREQSegment = findSegmentBySegmentNumber(bFREQSegment.get_segment_number(), comparisonFREQSequence)
                                                
                                                if not cFREQSegment:
                                                    wf.write("{} emitter({}).mode({}).generator({}).FREQSequence({}) contains FREQ Segment number: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), bFREQSegment.get_segment_number(), cFileName))
                                                else:
                                                    for bSegmentAttribute in bFREQSegment.get_attributes():
                                                        cSegmentAttribute = findAttribute(bSegmentAttribute.get_name(), cFREQSegment)
                                                        
                                                        if not cSegmentAttribute:
                                                            wf.write("{} emitter({}).mode({}).generator({}).FREQSequence({}).FREQSegment({}) contains attribute: {} that is missing from this path in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), bFREQSegment.get_segment_number(), bSegmentAttribute.get_name(), cFileName))
                                                        else:
                                                            if cSegmentAttribute.get_value() != bSegmentAttribute.get_value():
                                                                wf.write("{} emitter({}).mode({}).generator({}).FREQSequence({}).FREQSegment({}) contains attribute: {} with value:{} that is different than the value:{} in the same path attribute in {}.\n".format(bFileName, bElnot, baseEmitterMode.get_name(), baseGenerator.get_generator_number(), baseFREQSequence.get_ordinal_pos(), bFREQSegment.get_segment_number(), bSegmentAttribute.get_name(), bSegmentAttribute.get_value(), cSegmentAttribute.get_value(), cFileName))
                            
                            
                            
 
def compareTheFiles():

    print("parsing Base File ...")
    parseBaseFile()
    
    print("parsing Comparison File ...")
    parseComparisonFile()

    wf = open(writeFile, "w+")

    compareEMTFiles(wf, base_emitters, comparison_emitters, constant.COMPARISON_ARRAY, baseFileName, comparisonFileName)

    wf.write("\n\n")
    
    compareEMTFiles(wf, comparison_emitters, base_emitters, constant.BASE_ARRAY, comparisonFileName, baseFileName)

    wf.close()

           
if __name__ == '__main__':
    compareTheFiles()
    
                    
                
                
                       
                   

       