#!/usr/bin/env python
# pylint:disable=invalid-name, line-too-long, pointless-string-statement

'''
A script to create HL7 Pathology messages (orders and results) that comply with one of two Australian
HL7 Pathology Messaging Profiles (ADRM or IHENPAACR6) from synthetic data

SYNOPSIS
$ python mkHL7LabMessages.py [-D dataDir|--dataDir=dataDir]
                             [-P profile|--profile=profile]
                             [-S|--useSnomedSpecimenCodes]
                             [-R requestCode|--requestCode=requestCode]
                             [-A|--All]
                             [-N numberOfMessages|--numberOfMessages=numberOfMessages]
                             [-O outputDir|--outputDir=outputDir]
                             [-v loggingLevel|--verbose=loggingLevel]
                             [-L logDir|--logDir=logDir] [-l logfile|--logfile=logfile]


OPTIONS
-D dataDir|--dataDir=dataDir
The directory containing the synthetic demographic and clinical data (default='data')

-P profile|--profile=profile
The Australian HL7 Pathology Messageing Profile (default='ADRM')
Only ADRM or IHENPAACR6 allowed

-S|--useSnomedSpecimenCodes
Use SNOMED codes for specimen types (default=False)

-R requestCode|--requestCode=requestCode
The request code for the HL7 message (default=None)

-A|--All
Generate HL7 messages for all top 100 SPIA request codes (default=False)

-N numberOfMessages|--numberOfMessages=numberOfMessages
The number of HL7 messages pairs (Order and Results) to generate (default=1)

-O outputDir|--outputDir=outputDir
The directory in which the output file will be created (default='output')
[mkHealthPopulation.cfg will be read from this directory]

-v loggingLevel|--verbose=loggingLevel
Set the level of logging that you want.

-L logDir|--logDir=logDir
The name of the folder for the logging file

-l logfile|--logfile=logfile
The name of a logging file where you want all messages captured.

'''

from dataclasses import field
import os
import sys
import argparse
import logging
import configparser
import json
import isodate
import random
import csv
import pandas as pd
import base64
import zoneinfo
from openpyxl import load_workbook
from datetime import datetime as dt, timedelta as td
from dateutil import relativedelta as rd
from pytz import timezone as tz
from fpdf import FPDF

# This next section is plagurised from /usr/include/sysexits.h
EX_OK = 0                # successful termination
EX_WARN = 1                # non-fatal termination with warnings

EX_USAGE = 64            # command line usage error
EX_DATAERR = 65            # data format error
EX_NOINPUT = 66            # cannot open input
EX_NOUSER = 67            # addressee unknown
EX_NOHOST = 68            # host name unknown
EX_UNAVAILABLE = 69        # service unavailable
EX_SOFTWARE = 70        # internal software error
EX_OSERR = 71            # system error (e.g., can't fork)
EX_OSFILE = 72            # critical OS file missing
EX_CANTCREAT = 73        # can't create (user) output file
EX_IOERR = 74            # input/output error
EX_TEMPFAIL = 75        # temp failure; user is invited to retry
EX_PROTOCOL = 76        # remote error in protocol
EX_NOPERM = 77            # permission denied
EX_CONFIG = 78            # configuration error

fieldSep = '|'
compSep = '^'
repSep = '~'
subCompSep = '&'
escapeChar = '\\'
messageControlId = random.randint(100000, 999999)
specimenMapping = {
    "Serum": ["SER","119364003"],
    "Plasma": ["PLAS","119361006"],
    "Platelet poor plasma": ["PPP","119361006"],
    "Blood": ["BLD","119297000"],
    "Whole blood": ["BDY","258591005"],
    "Arterial blood": ["BLDA","122552005"],
    "Venous blood": ["BLDV","122555007"],
    "Urine": ["UR","122575003"],
    "24 hr urine": ["UR","122575003"],
    "Fluid": ["FLU","309051001"],
    "Bone marrow": ["MAR","119359002"],
    "Tissue": ["TISS","119376003"],
    "Lymph node": ["LYM","309078004"],
    "CSF": ["CSF","258450006"],
    "Faeces": ["STL","119339001"],
    "Breath": ["BRTH","119336008"],
    "Nasopharyngeal aspirate": ["NOS","258467004"],
    "Nasopharyngeal swab": ["NOS","257261003"],
    "Genital Swab": ["GEN","258508008"],
    "Cervix swab": ["CVX","258521001"]
}
disciplineMapping = {
    "Chemical pathology": "CH",
    "Cytopathology": "CP",
    "Haematology": "HM",
    "Immunopathology": "IMM",
    "Microbiology": "MB",
    "Molecular": "MB",
    "Serology": "SR"
}


def mkXCN(doc, sep, withRepeats):

    '''Create an XCN (Extended Composite ID Number and Name for Persons) field for a doctor.
       Or a CN if the sep character is not the component separator.'''

    global compSep, repSep

    if sep != compSep:
        withRepeats = False
    doctor = ''
    if doc['Provider No'] != '':
        if sep == compSep:
            drBits = [doc['Provider No'], doc[drMappings['familyName']], doc[drMappings['givenName']], '', '', doc['Title'], '', '', 'AUSHICPR', 'L', '', '', 'UPIN']
        else:               #  CN doesn't have a name type code
            drBits = [doc['Provider No'], doc[drMappings['familyName']], doc[drMappings['givenName']], '', '', doc['Title'], '', '', 'AUSHICPR']
        doctor += sep.join(drBits)
    if sep != compSep:      # Can't encode Prescriber or HPI-I in a CN field - both AUSHIC and no identifier type code
        return doctor
    if withRepeats or (doctor == ''):
        if doc['Prescriber No'] != '':
                drBits = [doc['Prescriber No'], doc[drMappings['familyName']], doc[drMappings['givenName']], '', '', doc['Title'], '', '', 'AUSHIC', 'L', '', '', 'PRES']
                if doctor != '':
                    doctor += repSep
                doctor += sep.join(drBits)
    if withRepeats or (doctor == ''):
        if doc['HPI-I'] != '':
            drBits = [doc['HPI-I'], doc[drMappings['familyName']], doc[drMappings['givenName']], '', '', doc['Title'], '', '', 'AUSHIC', 'L', '', '', 'NPI']
            if doctor != '':
                doctor += repSep
            doctor += sep.join(drBits)
    return doctor


def mkMSH(thisLab, thisLocalHospitalHPIO, thisClinic, messageType):

    '''Create the MSH (Message Header) segment for an HL7 message, from this clinic/lab to this lab/clinic (based on the message type).'''

    global fieldSep, compSep, repSep, subCompSep, escapeChar, messageControlId, clinics, Labs

    Fields = ['MSH']
    Fields.append(compSep + repSep + escapeChar + subCompSep)       # MSH-2: Encoding Characters
    if messageType == 'ORM':        # An order message - MSH-3 to MSH-6 are sender and receiver information
        Fields.append('mkHL7LabMessages')
        Fields.append(thisClinic[1] + compSep + thisClinic[0] + compSep + 'AUSHIC')
        Fields.append('LabWare Solutions')
        Fields.append(Labs[thisLab] + compSep + thisLab + compSep + 'AUSNATA')
    else:                           # A results message
        Fields.append('LabWare Solutions')
        Fields.append(Labs[thisLab] + compSep + thisLab + compSep + 'AUSNATA')
        Fields.append('mkHL7LabMessages')
        Fields.append(thisClinic[1] + compSep + thisClinic[0] + compSep + 'AUSHIC')
    melbourne = tz('Australia/Melbourne')
    if messageType == 'ORM':        # For order messages, backdate the timestamp by 3 days
        timeStamp = dt.now(melbourne) - td(days=3)
    else:                           # For results messages, use the current timestamp
        timeStamp = dt.now(melbourne)
    HL7timestamp = timeStamp.isoformat()
    HL7timestamp = HL7timestamp.replace('T', '').replace('-', '', 2).replace(':', '')
    HL7timestamp = HL7timestamp[0:19] + HL7timestamp[-5:]
    Fields.append(HL7timestamp)     # MSH-7: Date/Time of Message
    Fields.append('')               # Skip MSH-8: Security
    if messageType == 'ORM':        # For order messages, set the message type to ORM_O01 - MSH-9: Message Type
        Fields.append('ORM' + compSep + 'O01' + compSep + 'ORM_O01')
    else:                           # For results messages, set the message type to ORU_R01
        Fields.append('ORU' + compSep + 'R01' + compSep + 'ORU_R01')
    messageControlId += 1
    messageControlId %= 1000000
    Fields.append(f'{HL7timestamp[0:14]}{messageControlId:06d}')    # MSH-10: Message Control ID
    Fields.append('P')              # MSH-11: Processing ID
    if messageType == 'ORM':        # MSH-12: Version ID for order messages
        Fields.append(f'2.4{compSep}AUS{subCompSep}Australia{subCompSep}ISO3166_1{compSep}HL7AU-OO-ORM-201701{subCompSep}{subCompSep}L')
    else:                           # MSH-12: Version ID for results messages
        Fields.append(f'2.4{compSep}AUS{subCompSep}Australia{subCompSep}ISO3166_1{compSep}HL7AU-OO-ORU-201701{subCompSep}{subCompSep}L')
    Fields.append('')               # Skip MSH-13: Sequence Number
    Fields.append('')               # Skip MSH-14: Continuation Pointer
    Fields.append('AL')             # MSH-15: Accept Acknowledgment Type
    Fields.append('AL')             # MSH-16: Application Acknowledgment Type
    Fields.append('AUS')            # MSH-17: Country Code
    Fields.append('ASCII')           # MSH-18: Character Set
    Fields.append(f'en{compSep}English{compSep}ISO639')   # MSH-19: Principal Language
    Fields.append('')                 # Skip MSH-20: Alternate Character Set Handling Scheme
    Fields.append('HL7AUSD-STD-OO-ADRM-2018.1')   # MSH-21: Message Profile Identifier
    return fieldSep.join(Fields)


usedFillerNos = set()
usedPlacerNos = set()
thisPlacerNo = 0
thisFillerNo = 0
nextLN = {}
thisFiller1 = None
serviceID = None
changedTime = None
changedTimeStamp = None
receivedTime = None

def mkOBR(messageType, obrIndex, thisClinic, thisLab):

    global fieldSep, OBRs, usedFillerNos, usedPlacerNos, thisPlacerNo, thisFillerNo, orderObs, thisFiller1, serviceID
    global specimenMapping, specimenSnomed, profile, thisGPxcn, nextLN, thisSpecialistxcn, thisLabStaffcn
    global changedTime, changedTimeStamp, orderPIT, resultPIT
    global receivedTime, Labs

    Fields = ['OBR']
    Fields.append(str(obrIndex + 1))                # OBR-1: Set ID - OBR
    if messageType == 'ORM':
        thisPlacerNo = random.randint(1000, 1000000)
        while thisPlacerNo in usedPlacerNos:
            thisPlacerNo = random.randint(1000, 1000000)
        usedPlacerNos.add(thisPlacerNo)
        Fields.append(f'{str(thisPlacerNo)}{compSep}{thisClinic[0]}{compSep}{thisClinic[1]}{compSep}AUSHIC')            # OBR-2: Placer Order Number
        Fields.append('')                           # OBR-3: Filler Order Number
    else:
        thisFillerNo = random.randint(1000, 1000000)
        while thisFillerNo in usedFillerNos:
            thisFillerNo = random.randint(1000, 1000000)
        usedFillerNos.add(thisFillerNo)
        Fields.append(f'{str(thisPlacerNo)}{compSep}{thisClinic[0]}{compSep}{thisClinic[1]}{compSep}AUSHIC')            # OBR-2: Placer Order Number
        Fields.append(f'{str(thisFillerNo)}{compSep}{thisLab}{compSep}{Labs[thisLab]}{compSep}AUSNATA')            # OBR-3: Filler Order Number
    orderCode = OBRs[obrIndex][0]
    Fields.append(f'{orderCode}{compSep}{orderObs[orderCode]['Description']}{compSep}SCT')               # OBR-4: Universal Service Identifier
    Fields.append('')               # OBR-5: Priority
    Fields.append('')               # OBR-6: Requested Date/Time
    Fields.append('')               # OBR-7: Observation Date/Time
    Fields.append('')               # OBR-8: Observation End Date/Time
    Fields.append('')               # OBR-9: Collection Volume
    Fields.append('')               # OBR-10: Collector Identifier
    Fields.append('')               # OBR-11: Specimen Action Code
    Fields.append('')               # OBR-12: Danger Code
    Fields.append('')               # OBR-13: Relevant Clinical Information
    melbourne = tz('Australia/Melbourne')
    receivedTime = dt.now(melbourne) - td(days=3) + td(hours=random.randint(4, 12))
    receivedTimeStamp = receivedTime.isoformat()
    receivedTimeStamp = receivedTimeStamp.replace('T', '').replace('-', '', 2).replace(':', '')
    receivedTimeStamp = receivedTimeStamp[0:19] + receivedTimeStamp[-5:]
    Fields.append(receivedTimeStamp)                # OBR-14: Specimen Received Date/Time
    specimen = OBRs[obrIndex][1]
    if useSnomedSpecimenCodes:
        specimen = specimenMapping[specimen][1]
        Fields.append(f'{specimen}{subCompSep}{specimenSnomed[specimen]}{subCompSep}SCT')               # OBR-15: Specimen Source
    elif profile == 'ADRM':
        specimen = specimenMapping[specimen][0]
        Fields.append(f'{specimen}{subCompSep}{hl7Table0070[specimen]}{subCompSep}HL70070')           # OBR-15: Specimen Source
    else:
        Fields.append(f'{specimen}{subCompSep}{subCompSep}L')                                                 # OBR-15: Specimen Source
    Fields.append(thisGPxcn)                        # OBR-16: Ordering Provider
    Fields.append('')               # OBR-17: Order Callback Phone Number
    Fields.append('')               # OBR-18: Placer Field 1
    Fields.append('')               # OBR-19: Placer Field 2
    if messageType == 'ORM':
        serviceID = disciplineMapping[orderObs[thisOrder]['Discipline']]
        if serviceID not in nextLN:
            nextLN[serviceID] = random.randint(50, 100)
        nextLN[serviceID] += 1
        thisFiller1 = f'{dt.now(melbourne).year}-{nextLN[serviceID]}-{serviceID}'
        if thisSpecialistxcn != '':
            specialistNames = thisSpecialistxcn.split(repSep)
            for specialistName in specialistNames:
                thisName = specialistName.split(compSep)
                if thisName[8] == 'AUSHICPR':
                    thisFiller1 += f',DR={thisName[0]},CP=Y'
                    break
    Fields.append(thisFiller1)                      # OBR-20: Filler Field 1
    Fields.append('')               # OBR-21: Filler Field 2
    changedTime = dt.now(melbourne) - td(hours=random.randint(1, 3))
    changedTimeStamp = changedTime.isoformat()
    changedTimeStamp = changedTimeStamp.replace('T', '').replace('-', '', 2).replace(':', '')
    changedTimeStamp = changedTimeStamp[0:19] + changedTimeStamp[-5:]
    Fields.append(changedTimeStamp)                 # OBR-22: Results Rpt/Status Chng - Date/Time
    Fields.append('')               # OBR-23: Charge to Practice
    Fields.append(serviceID)                        # OBR-24: Diagnostic Serv Sect ID
    Fields.append('F')                              # OBR-25: Result Status
    Fields.append('')               # OBR-26: Parent Result
    Fields.append(f'{compSep}{compSep}{compSep}{receivedTimeStamp}{compSep}{compSep}RT')               # OBR-27: Quantity/Timing
    Fields.append(thisSpecialistxcn)                       # OBR-28: Result Copies To
    Fields.append('')               # OBR-29: Parent
    Fields.append('')               # OBR-30: Transportation Mode
    Fields.append('')               # OBR-31: Reason for Study
    Fields.append(thisLabStaffcn)                   # OBR-32: Principal Result Interpreter
    # Fields.append('')               # OBR-33: Assistant Result Interpreter
    # Fields.append('')               # OBR-34: Technician
    # Fields.append('')               # OBR-35: Transcriptionist
    # Fields.append('')               # OBR-36: Scheduled Date/Time
    # Fields.append('')               # OBR-37: Number of Sample Containers
    # Fields.append('')               # OBR-38: Transport Logistics of Collected Sample
    # Fields.append('')               # OBR-39: Collector's Comment
    # Fields.append('')               # OBR-40: Transport Arrangement Responsibility
    # Fields.append('')               # OBR-41: Transport Arranged
    # Fields.append('')               # OBR-42: Escort Required
    # Fields.append('')               # OBR-43: Planned Patient Transport Comment
    # Fields.append('')               # OBR-44: Procedure Code
    # Fields.append('')               # OBR-45: Procedure Code Modifier
    # Fields.append('')               # OBR-46: Placer supplemental service information
    # Fields.append('')               # OBR-47: Filler supplemental service information
    return fieldSep.join(Fields)


def mkOBX(obrIndex, obxIndex):

    global OBXs, spia, UCUM, UCUMunit, thisLabStaffxcn

    Fields = ['OBX']
    Fields.append(str(obxIndex + 1))                # OBX-1: Set ID - OBX
    Fields.append(OBXs[obrIndex][obxIndex][0])      # OBX-2: Value Type
    testCode = OBXs[obrIndex][obxIndex][1]
    identifier = f'{testCode}{compSep}{spia[testCode][0]}{compSep}LN'
    Fields.append(identifier)                       # OBX-3: Observation Identifier
    Fields.append('')               # OBX-4: Observation Sub-ID
    Fields.append(str(OBXs[obrIndex][obxIndex][2])) # OBX-5: Observation Value
    if (spia[testCode][2] is not None) and (spia[testCode][2] != 'No unit') and (spia[testCode][1] != 'no unit'):   # We have UCUM
        code = spia[testCode][2].replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
        display = UCUM[spia[testCode][2]].replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
        units = f'{code}{compSep}{display}{compSep}UCUM'
    elif (spia[testCode][1] is not None) and (spia[testCode][1] != 'No unit') and (spia[testCode][1] != 'no unit'):   # We have a non-UCUM unit
        if spia[testCode][1] in UCUMunit:       # But it's actually a UCUM unit
            thisUCUM = UCUMunit[spia[testCode][1]]
            code = thisUCUM.replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
            display = UCUM[thisUCUM].replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
            units = f'{code}{compSep}{display}{compSep}UCUM'
        else:
            code = spia[testCode][1].replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
            display = spia[testCode][1].replace(escapeChar, '\\E\\').replace(fieldSep, '\\F\\').replace(repSep, '\\R\\').replace(compSep, '\\S\\').replace(subCompSep, '\\T\\')
            units = f'{code}{compSep}{display}{compSep}L'
    else:
        units = ''
    Fields.append(units)                            # OBX-6: Units
    if OBXs[obrIndex][obxIndex][4] != '':
        Fields.append(f'{OBXs[obrIndex][obxIndex][4]}')               # OBX-7: References Range
    else:
        Fields.append('')               # OBX-7: References Range
    Fields.append(OBXs[obrIndex][obxIndex][3])      # OBX-8: Abnormal Flags
    Fields.append('')               # OBX-9: Probability
    Fields.append('')               # OBX-10: Nature of Abnormal Test
    Fields.append('F')               # OBX-11: Observation Result Status
    Fields.append('')               # OBX-12: Date Last Obs Normal Values
    Fields.append('')               # OBX-13: User Defined Access Checks
    Fields.append(changedTimeStamp)                 # OBX-14: Date/Time of the Observation
    Fields.append('')               # OBX-15: Producer's ID
    Fields.append(thisLabStaffxcn)                  # OBX-16: Responsible Observer
    # Fields.append('')               # OBX-17: Observation Method
    # Fields.append('')               # OBX-18: Equipment Instance Identifier
    # Fields.append('')               # OBX-19: Date/Time of the Analysis
    return fieldSep.join(Fields)


def mkePIT(obrIndex):

    global OBXs, compSep, fieldSep, dataDir
    global Labs, thisLab, thisGPxcn, thisSpecialistxcn, PIDfields, patientBirthdate, patientSex
    global collectionTime, useSnomedSpecimenCodes, profile

    pdf = FPDF(format='A4', orientation='L', enforce_compliance="PDF/A-1b")
    pdf.add_page()
    pdf.add_font('Courier New', fname=os.path.join(dataDir, 'cour.ttf'))
    pdf.set_font('Courier New', size=10)
    pdf.cell(0, 5, f'001 {Labs[thisLab]:59}07 01/07/1996', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '002                                                  ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'003 {"Report Run Number:":24}{"Created":26}at', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'004 {"Surgery:":16}{"Reports:":30}to', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '006 Hospital:', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '009 -------------------------------------------------', new_x="LMARGIN", new_y="NEXT")
    GPnames = thisGPxcn.split(repSep)
    for GPname in GPnames:
        thisName = GPname.split(compSep)
        if thisName[8] == 'AUSHICPR':
            pdf.cell(0, 5, f'010 {(thisName[5] + " " +thisName[2] + " " + thisName[1])[:32]:32}{" ":13}{thisName[0]}', new_x="LMARGIN", new_y="NEXT")
            break
    else:
        pdf.cell(0, 5, '010 ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '019 --------------------------------------------------', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'020 {"Your Ref.":12}{"Patient Name":31}{"Lab Ref.":13}Test Code', new_x="LMARGIN", new_y="NEXT")
    patientName = PIDfields[5].split(compSep)
    if patientName[2] != '':
        pdf.cell(0, 5, f'021 {" ":12}{(patientName[1] + " " + patientName[2] + " " + patientName[0])[:32]}', new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.cell(0, 5, f'021 {" ":12}{(patientName[1] + " " + patientName[0])[:32]}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '029 --------------------------------------------------', new_x="LMARGIN", new_y="NEXT")
    patientName = PIDfields[5].split(compSep)
    thisName = ''
    if patientName[4] != '':
        thisName = patientName[4] + ' '
    thisName += patientName[1] + ' '
    if patientName[2] != '':
        thisName += patientName[2] + ' '
    thisName += patientName[0]
    if patientName[3] != '':
        thisName += ' ' + patientName[3]
    pdf.cell(0, 5, f'100 {"Start Patient :":22}{thisName[:32]}', new_x="LMARGIN", new_y="NEXT")
    patientAddress = PIDfields[11].split(compSep)
    pdf.cell(0, 5, f'101 {" ":22}{patientAddress[0]},{patientAddress[2]},{patientAddress[3]},{patientAddress[4]}', new_x="LMARGIN", new_y="NEXT")
    thisAge = rd.relativedelta(dt.now(), patientBirthdate)
    if thisAge.years > 0:
        age = f'Y{thisAge.years}'
    elif thisAge.months > 0:
        age = f'M{thisAge.months:02d}'
    else:
        age = f'D{thisAge.days}'
    pdf.cell(0, 5, f'104 {" ":22}{"Birthdate:":11}{patientBirthdate.strftime("%d/%m/%Y")}    Age:  {age:4}  Sex: {patientSex}', new_x="LMARGIN", new_y="NEXT")
    phone = PIDfields[13]
    if phone == '':
        phone = PIDfields[14]
    if phone != '':
        thesePhones = phone.split(repSep)
        phone = ''
        for thisPhone in thesePhones:
            thisPhone = thisPhone.split(compSep)
            if thisPhone[2] == 'PH':
                phone = thisPhone[5] + ' ' + thisPhone[6]
            elif thisPhone[2] == 'CP':
                phone = thisPhone[6]
    pdf.cell(0, 5, f'105 {" ":22}{"Telephone:":11}{phone}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '109                                                  ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'110 {"Your Reference:":22}{thisPlacerNo}', new_x="LMARGIN", new_y="NEXT")
    labNo = thisFiller1.split(',')[0]
    pdf.cell(0, 5, f'111 {Labs[thisLab][:11] + ' Reference:':22}{labNo}', new_x="LMARGIN", new_y="NEXT")
    patientNumbers = PIDfields[3]
    for thisNumber in patientNumbers.split(repSep):
        thisNumber = thisNumber.split(compSep)
        if (thisNumber[3] == 'AUSHIC') and (thisNumber[4] == 'MC'):
            patientMedicare = thisNumber[0]
            break
    else:
        patientMedicare = ''
    pdf.cell(0, 5, f'112 {"Medicare Number:":22}{patientMedicare}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '115 Phone Enquiries:', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '119                                                  ', new_x="LMARGIN", new_y="NEXT")
    thisGPbits = thisGPxcn.split(compSep)
    pdf.cell(0, 5, f'121 {"Referred By... :":22}{thisGPbits[5]} {thisGPbits[2]} {thisGPbits[1]}', new_x="LMARGIN", new_y="NEXT")
    if thisSpecialistxcn != '':
        specialistBits = thisSpecialistxcn.split(compSep)
        pdf.cell(0, 5, f'122 {"Specialist:":22}{specialistBits[5]} {specialistBits[2]} {specialistBits[1]}', new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.cell(0, 5, f'122 {"Specialist:":22}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '123 Addressee:', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '129                                                    ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '130 Ward:', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '131 Auto Ward Print:', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '139                                                    ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '200 Start of Result:', new_x="LMARGIN", new_y="NEXT")
    specimen = OBRs[obrIndex][1]
    if useSnomedSpecimenCodes:
        specimen = specimenMapping[specimen][1]
        specimen = specimenSnomed[specimen]
    elif profile == 'ADRM':
        specimen = specimenMapping[specimen][0]
        specimen = hl7Table0070[specimen]
    pdf.cell(0, 5, f'201 {"Specimen:":22}{specimen}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'203 {"Request Date:":22}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'204 {"Collection Date:":22}{receivedTime.strftime("%d/%m/%Y %H:%M")}', new_x="LMARGIN", new_y="NEXT")
    orderCode = OBRs[obrIndex][0]
    pdf.cell(0, 5, f'205 {"Name of Test:":22}{orderObs[orderCode]['Description']}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'206 {"Reported:":22}{changedTime.strftime("%d/%m/%Y %H:%M")}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'207 {"Confidential:":22}{"N"}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'208 {"Test Category:":22}R', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '209                                                    ', new_x="LMARGIN", new_y="NEXT")
    normal = 'N'
    requestedTests = ''
    for i in range(len(OBXs[obrIndex])):
        if requestedTests != '':
            requestedTests += ', '
        requestedTests += spia[OBXs[obrIndex][i][1]][0]
        if OBXs[obrIndex][i][3] != 'N':
            normal = 'A'
    pdf.cell(0, 5, f'210 {"Normal Result:":22}{normal}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'211 {"Requested Tests:":22}{requestedTests}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'212 {"Request Complete:":22}{"Y"}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '299                                                    ', new_x="LMARGIN", new_y="NEXT")
    for i in range(len(OBXs[obrIndex])):
        testCode = OBXs[obrIndex][i][1]
        test = spia[testCode][0]
        value = OBXs[obrIndex][i][2]
        if (spia[testCode][2] is not None) and (spia[testCode][2] != 'No unit') and (spia[testCode][2] != 'no unit'):   # We have UCUM
            display = UCUM[spia[testCode][2]]
        elif (spia[testCode][1] is not None) and (spia[testCode][1] != 'No unit') and (spia[testCode][1] != 'no unit'):   # We have a non-UCUM unit
            if spia[testCode][1] in UCUMunit:       # But it's actually a UCUM unit
                thisUCUM = UCUMunit[spia[testCode][1]]
                display = UCUM[thisUCUM]
            else:
                display = spia[testCode][1]
        else:
            display = ''
        pdf.cell(0, 5, f'301 {test:55} {value:8} {OBXs[obrIndex][i][4]:10} ({OBXs[obrIndex][i][3]}) {display:8}', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '309                                                    ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '311 ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '319                                                    ', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '390 End Of Report :', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, '399 ---------------------------------------------------', new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f'999 {"END OF LISTING - Run Number:":34}{changedTime.strftime("%d/%m/%Y  %H:%M:%S")}', new_x="LMARGIN", new_y="NEXT")

    pdf_byte = pdf.output()
    base64_pdf = base64.b64encode(pdf_byte).decode('utf-8')

    Fields = ["OBX"]
    Fields.append(str(len(OBXs[obrIndex]) + 1))               # OBX-1: Set ID
    Fields.append("ED")                                       # OBX-2: Value Type
    Fields.append(f'PDF{compSep}Display formant in PDF{compSep}AUSPDI')  # OBX-3: Observation Identifier
    Fields.append('')                                         # OBX-4: Observation Sub-ID
    Fields.append(f'{compSep}application{compSep}pdf{compSep}Base64{compSep}{base64_pdf}')  # OBX-5: Value
    Fields.append('')                                         # OBX-6: Units
    Fields.append('')                                         # OBX-7: References Range
    Fields.append('')                                         # OBX-8: Abnormal Flags
    Fields.append('')                                         # OBX-9: Probability
    Fields.append('')                                         # OBX-10: Nature of Abnormal Test
    Fields.append('F')                                         # OBX-11: Observation Result Status

    return fieldSep.join(Fields)


bloods = ['Blood','Whole blood','Venous blood','Arterial blood','Plasma','Serum']
urine = ['Urine','Midstream urine','Catheter urine','24 hr urine']
def buildOrderObs(thisOrder):

    global logger, referenceData, orderObs

    # Create a dictionary of tests for each order specimen type
    # When we create a message we are going to randomly choose an order code.
    # Each order code has a list of order specimen types - orderObs[orderCode]['orderSpecimenList']
    # However, each order code also has a list of tests - orderObs[orderCode]['Tests']
    # and each test has list of test specimen types because different tests can only be run on specific specimen types.
    # So the list of possible tests for each order specimen type may be less than all tests for this order code.
    # To confound things, we have reference data for each test, but only for specific specimen types.
    # We may be missing reference data for a test specimen type even if the test is valid for the order specimen type.
    # Hence, we create a dictionary of tests for each order specimen type where the value is the reference data specimen type
    # that will enable us to create valid test results data (OBX-5) for each order specimen type.
    possibleOrderSpecimens = set(orderObs[thisOrder]['orderSpecimenList'])
    if 'Any' in possibleOrderSpecimens:          # Use the test specimen types
        possibleOrderSpecimens = set()
        for testCode in orderObs[thisOrder]['Tests']:
            possibleOrderSpecimens = possibleOrderSpecimens | set(orderObs[thisOrder]['Tests'][testCode])
        if 'Any' in possibleOrderSpecimens:         # for Any/Any use 'Blood' as the default specimen type
            possibleOrderSpecimens.remove('Any')
            possibleOrderSpecimens.add('Blood')
    testedTests = set()
    foundTests = set()
    for orderSpecimen in possibleOrderSpecimens:      # For each order specimen type we need a dictionary of matching tests and reference data specimen type
        for testCode in orderObs[thisOrder]['Tests']:
            testedTests.add(testCode)
            thisSpecimen = None
            for testSpecimen in orderObs[thisOrder]['Tests'][testCode]:
                if testSpecimen == orderSpecimen:
                    thisSpecimen = orderSpecimen
                    break
                elif testSpecimen == 'Any':
                    thisSpecimen = orderSpecimen
                    break
                elif (orderSpecimen in bloods) and (testSpecimen in bloods):
                    thisSpecimen = testSpecimen
                    break
                elif (orderSpecimen in urine) and (testSpecimen in urine):
                    thisSpecimen = testSpecimen
                    break
            else:
                continue
            # Find a reference data specimen that matches thisSpecimen
            for refSpecimen in referenceData[testCode]['specimen']:
                isMatched = False
                if (refSpecimen == thisSpecimen) or (refSpecimen == 'all'):
                    isMatched = True
                elif (thisSpecimen in bloods) and (refSpecimen in bloods) :
                    isMatched = True
                elif (thisSpecimen in urine) and (refSpecimen in urine):
                    isMatched = True
                if isMatched:
                    if orderSpecimen not in orderObs[thisOrder]['Specimen']:
                        orderObs[thisOrder]['Specimen'][orderSpecimen] = {}
                    if thisSpecimen not in orderObs[thisOrder]['Specimen'][orderSpecimen]:
                        orderObs[thisOrder]['Specimen'][orderSpecimen][thisSpecimen] = {}
                    orderObs[thisOrder]['Specimen'][orderSpecimen][thisSpecimen][testCode] = refSpecimen
                    foundTests.add(testCode)
                    break
        # Now check how many test we have for this order specimen
        if len(orderObs[thisOrder]['Specimen']) == 0:
            logger.warning(f'Order {thisOrder} has no matching test specimen for order specimen "{orderSpecimen}" in file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
            continue
        if len(foundTests) < (len(testedTests) / 2):
            logger.warning(f'Order {thisOrder} - less than half of the tests for order specimen "{orderSpecimen}" have a matching test specimen - file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
            logger.warning(f'Order {thisOrder} has the following specimens/tests: {orderObs[thisOrder]["Specimen"]}')
            logger.warning(f'out of {list(orderObs[thisOrder]["Tests"].keys())}')
            continue
    return

# Define a logging Formatter class
# To suppress program/level/datetime at the start of a logging record use
# logger.xxxx('message', extra={'raw_message':True})
class ConditionalFormatter(logging.Formatter):
    def format(self, record):
        if hasattr(record, 'raw_message') and record.raw_message:
            return record.getMessage()
        else:
            return super().format(record)


if __name__ == '__main__':
    '''
The main code
    '''

    # Save the program name
    progName = sys.argv[0]
    progName = progName[0:-3]        # Strip off the .py ending

    parser = argparse.ArgumentParser()
    parser.add_argument('-D', '--dataDir', dest='dataDir', default='data',
                        help='The name of the directory containing synthetic demographic and clinical data(default="data")')
    parser.add_argument('-P', '--profile', dest='profile', default='ADRM', choices=['ADRM', 'IHENPAACR6'],
                        help='The Australian HL7 Laboratory Messaging Profile (default="ADRM")')
    parser.add_argument('-S', '--useSnomedSpecimenCodes', dest='useSnomedSpecimenCodes', action='store_true',
                        help='Use SNOMED codes for specimen types (default=False)')
    parser.add_argument('-R', '--requestCode', dest='requestCode', default=None,
                        help='The SNOMED CT request code for the laboratory messages')
    parser.add_argument('-A', '--All', dest='allRequests', action='store_true',
                        help='Include all SPIA top 100 request codes in the laboratory messages (default=False)')
    parser.add_argument('-N', '--numberOfMessages', dest='numberOfMessages', type=int, default=0,
                        help='The number of random laboratory messages to generate')
    parser.add_argument('-O', '--outputDir', dest='outputDir', default='output',
                        help='The name of the output directory (default="output")')
    parser.add_argument('-v', '--verbose', dest='verbose', type=int, choices=range(0, 5),
                        help='The level of logging\n\t0=CRITICAL,1=ERROR,2=WARNING,3=INFO,4=DEBUG')
    parser.add_argument('-L', '--logDir', dest='logDir', default='logs',
                        help='The name of a directory for the logging file(default="logs")')
    parser.add_argument('-l', '--logfile', metavar='logfile',
                        dest='logfile', help='The name of a logging file')
    args = parser.parse_args()

    # Set up logging
    loggingLevels = {0:logging.CRITICAL, 1:logging.ERROR, 2:logging.WARNING, 3:logging.INFO, 4:logging.DEBUG}
    logger = logging.getLogger(__name__)
    if args.verbose:
        logger.setLevel(loggingLevels[args.verbose])
    else:
        logger.setLevel(logging.WARNING)
    if args.logfile:
        handler = logging.FileHandler(os.path.join(args.logDir, args.logfile), mode='w')
    else:
        handler = logging.StreamHandler()
    logformat = progName + ' %(levelname)s[%(asctime)s]: %(message)s'
    dateformat = '%d/%m/%y'
    formatter = ConditionalFormatter(logformat, dateformat)
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    # Parse the optional arguments
    dataDir = args.dataDir
    profile = args.profile
    useSnomedSpecimenCodes = args.useSnomedSpecimenCodes
    requestCode = args.requestCode
    allRequests = args.allRequests
    numberOfMessages = args.numberOfMessages
    outputDir = args.outputDir

    if allRequests and requestCode:
        logger.fatal('Cannot specify both --allRequests and --requestCode')
        logging.shutdown()
        sys.exit(EX_CONFIG)
    if allRequests and (numberOfMessages > 0):
        logger.fatal('Cannot specify both --allRequests and --numberOfMessages')
        logging.shutdown()
        sys.exit(EX_CONFIG)
    if requestCode and (numberOfMessages > 0):
        logger.fatal('Cannot specify both --requestCode and --numberOfMessages')
        logging.shutdown()
        sys.exit(EX_CONFIG)

    # Start by reading in the synthetic demographic data - we need the heading mappings for the name fields
    drMappings = {}
    config = configparser.ConfigParser(allow_no_value=True)
    config.optionxform = str
    try:
        config.read(os.path.join(dataDir, 'mkHealthPopulation.cfg'))
        drFields = config.get('Fields', 'drFields')
        for row in csv.reader([drFields], csv.excel):
            drFields = row
        drFieldNames = config.get('Fields', 'drFieldNames')
        for row in csv.reader([drFieldNames], csv.excel):
            drFieldNames = row
        if len(drFields) != len(drFieldNames):
            logger.fatal('Incorrect number of drFieldNames for drFields')
            logging.shutdown()
            sys.exit(EX_CONFIG)
        drMappings = {}
        for i in range(len(drFields)):
            drMappings[drFields[i]] = drFieldNames[i]
        if (not 'familyName' in drMappings) or (not 'givenName' in drMappings):
            logger.fatal('Missing required doctor field mappings for familyName or givenName')
            logging.shutdown()
            sys.exit(EX_CONFIG)
    except (configparser.MissingSectionHeaderError, configparser.NoSectionError,
            configparser.NoOptionError, configparser.ParsingError) as detail:
        logger.fatal('%s', detail)
        logging.shutdown()
        sys.exit(EX_CONFIG)

    # Now load the Excel workbook containing the synthetic healthcare population data
    wb = load_workbook(filename=os.path.join(dataDir, 'healthPopulation.xlsx'))
    # Start with the Specialist Services associated with each local hospital and the Specialists associated with each Specialist Service
    ws = wb['Specialist Services']
    specialistServices = {}
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisSpecialistService = dict(zip(headers, row))
        localHospitalHPIO = thisSpecialistService['Local Public Hospital HPI-O']
        if localHospitalHPIO not in specialistServices:
            specialistServices[localHospitalHPIO] = []
        specialistServices[localHospitalHPIO].append(thisSpecialistService['Specialist Service HPI-O'])        # Specialist Service HPI-O
    ws = wb['Specialists']
    specialists = {}              # A dictionary where the key is the Clinic HPI-O and the value is a list of specialists associated with that clinic
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisSpecialist = dict(zip(headers, row))
        noi = thisSpecialist['Specialist Service HPI-O']            # Specialist Service HPI-O
        if noi not in specialists:
            specialists[noi] = []
        thisSpecialist['HPI-I'] = thisSpecialist['Specialist HPI-I']
        specialists[noi].append(mkXCN(thisSpecialist, compSep, True))

    # Next, load the GP clinics associated with each local hospital and the GPs associated with each GP clinic
    clinics = {}                # A dictionary where the key is the Clinic HPI-O and the value is the clinic data
    ws = wb['GP Clinics']
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisClinic = dict(zip(headers, row))
        localHospitalHPIO = thisClinic['Local Public Hospital HPI-O']
        if localHospitalHPIO not in clinics:
            clinics[localHospitalHPIO] = []
        noi = thisClinic['GP Clinic HPI-O']
        clinicName = thisClinic['Clinic Name']
        clinics[localHospitalHPIO].append([noi, clinicName])
    ws = wb['GP Clinic Staff']
    GPs = {}                    # A dictionary where the key is the Clinic HPI-O and the value is a list of GPs associated with that clinic
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisGP = dict(zip(headers, row))
        noi = thisGP['GP Clinic HPI-O']            # The Clinic HPI-O to which this GP belongs
        if noi not in GPs:
            GPs[noi] = []
        npi = thisGP['GP HPI-I']
        thisGP['HPI-I'] = thisGP['GP HPI-I']
        GPs[noi].append([npi, mkXCN(thisGP, compSep, True)])

    # And finally the patients - we need to associate each patient with their GP.
    ws = wb['Patients']
    GPpatients = {}
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisPatient = dict(zip(headers, row))
        npi = thisPatient['GP HPI-I']                    # The GP's HPI-I
        if npi not in GPpatients:
            GPpatients[npi] = []
        ihi = thisPatient['Patient IHI']
        GPpatients[npi].append(ihi)

    # And for each patient, we need to get their PID segment from the HL7_PID sheet
    ws = wb['HL7_PID']
    patientPID = {}
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    for row in data:
        if row[0] is None:
            break
        thisPatient = dict(zip(headers, row))
        ihi = thisPatient['IHI']                    # The patient's IHI
        patientPID[ihi] = thisPatient['PID']

    # Then the Labs
    labWB = load_workbook(os.path.join(dataDir,'NATAdata.xlsx'))
    ws = labWB['Organisation']
    data = ws.iter_rows(values_only=True)
    headers = next(data)
    Labs = {}
    for row in data:
        if row[0] is None:
            break
        thisLab = dict(zip(headers, row))
        if thisLab['Organisation Name'] not in Labs:
            nataAccred = thisLab['NATA Accreditation Number']
            Labs[nataAccred] = thisLab['Organisation Name']

    # And some lab staff
    labStaff = []
    for sheet in ['Public Hospital Staff', 'Private Hospital Staff']:
        ws = wb[sheet]
        data = ws.iter_rows(values_only=True)
        headers = next(data)
        for row in data:
            if row[0] is None:
                break
            thisStaff = dict(zip(headers, row))
            thisStaff['HPI-I'] = thisStaff['Hospital Staff HPI-I']
            XCN = mkXCN(thisStaff, compSep, True)
            CN = mkXCN(thisStaff, subCompSep, False)
            if CN == '':
                continue
            labStaff.append([XCN, CN])

    # Now for the clinical stuff
    # Firstly the SNOMED_CT Specimen Type Codes
    with open(os.path.join(dataDir,'Specimen-type-reference-set-20260930-expansion.json')) as f:
        snomedSpecimenTypes = json.load(f)
    specimenSnomed = {}
    for item in snomedSpecimenTypes['expansion']['contains']:
        if ('code' in item) and ('display' in item) and ('system' in item) and (item['system'] == 'http://snomed.info/sct'):
            specimenSnomed[item['code']] = item['display']
    # Then get HL7 Tables 0070 and 0074 from Appendix A
    hl7Table0002 = {}
    hl7Table0005 = {}
    hl7Table0070 = {}
    hl7Table0074 = {}
    dTypes = {'Type':str, 'Table':str, 'Value':str, 'Description':str, 'Seg':str, 'Seq#':str, 'Len':str}
    dfSheets = pd.read_excel(os.path.join(dataDir, "Appendix A.xlsx"), sheet_name=None, dtype=dTypes, na_filter=False)
    if "Appendix A.6 Tables Numeric" not in dfSheets:
        logger.critical('Excel Workbook "Appendix A.xlsx" in schemaDir folder(%s/xsd) does not have a Worksheet named "Appendx A.6 Tables Numeric"', dataDir)
        logging.shutdown()
        sys.exit(EX_CONFIG)
    tablesDf = dfSheets["Appendix A.6 Tables Numeric"]
    tablesDf['Table'] = tablesDf['Table'].str.zfill(4)
    tablesDf.loc[tablesDf['Table'].eq("0000"), 'Table'] = ''
    if "Type" not in tablesDf.columns.values.tolist():
        logger.critical('Missing column "Type" in Worksheet "Appendx A.6 Tables Numeric" in Excel Workbook "Appendix A.xlsx" in schemaDir folder(%s/xsd)', dataDir)
        logging.shutdown()
        sys.exit(EX_CONFIG)
    if "Table" not in tablesDf.columns.values.tolist():
        logger.critical('Missing column "Table" in Worksheet "Appendx A.6 Tables Numeric" in Excel Workbook "Appendix A.xlsx" in schemaDir folder(%s/xsd)', dataDir)
        logging.shutdown()
        sys.exit(EX_CONFIG)
    if "Value" not in tablesDf.columns.values.tolist():
        logger.critical('Missing column "Value" in Worksheet "Appendx A.6 Tables Numeric" in Excel Workbook "Appendix A.xlsx" in schemaDir folder(%s/xsd)', dataDir)
        logging.shutdown()
        sys.exit(EX_CONFIG)
    if "Description" not in tablesDf.columns.values.tolist():
        logger.critical('Missing column "Description" in Worksheet "Appendx A.6 Tables Numeric" in Excel Workbook "Appendix A.xlsx" in schemaDir folder(%s/xsd)', dataDir)
        logging.shutdown()
        sys.exit(EX_CONFIG)
    tablesData = tablesDf.to_dict(orient='records')
    tableNumber = None
    tableCode = None
    for row in tablesData:
        if row['Table'] == '':
            continue                    # A "Type" only row
        if row['Table'] != tableNumber:
            tableNumber = row['Table']
            tableCode = f'HL7{tableNumber:04}'
            if tableCode not in ['HL70002', 'HL70005', 'HL70070', 'HL70074']:
                continue
        tableValue = row['Value'].strip()
        if tableValue == '':
            continue
        valueDescription = row['Description'].strip()
        if tableCode == 'HL70002':
            hl7Table0002[tableValue] = valueDescription
        elif tableCode == 'HL70005':
            hl7Table0005[tableValue] = valueDescription
        elif tableCode == 'HL70070':
            hl7Table0070[tableValue] = valueDescription
        elif tableCode == 'HL70074':
            hl7Table0074[tableValue] = valueDescription

    # Now check the specimenMappings
    for specimen in specimenMapping:
        HL7specimenCode = specimenMapping[specimen][0]
        snomedSpecimenCode = specimenMapping[specimen][1]
        if HL7specimenCode not in hl7Table0070:
            logger.fatal(f"HL7 specimen mapping code {HL7specimenCode} for SPIA specimen {specimen} is not in HL70070 table")
            logging.shutdown()
            sys.exit(EX_CONFIG)
        if snomedSpecimenCode  not in specimenSnomed:
            logger.fatal(f"SNOMED specimen mapping code {snomedSpecimenCode} for SPIA specimen {specimen} is not in SNOMED specimen codes")
            logging.shutdown()
            sys.exit(EX_CONFIG)

    # Then the reference material
    with open(os.path.join(dataDir,'referenceIntervals.json')) as f:
        referenceIntervals = json.load(f)
    with open(os.path.join(dataDir,'referenceCodes.json')) as f:
        referenceCodes = json.load(f)
    # Check we haven't doubled up
    for LOINC in referenceIntervals:
        if LOINC in referenceCodes:
            logger.fatal(f"LOINC {LOINC} is present in both referenceIntervals and referenceCodes")
            logging.shutdown()
            sys.exit(EX_CONFIG)
        referenceIntervals[LOINC]['type'] = 'Intervals'
        for specimen in referenceIntervals[LOINC]['specimen']:
            for sex in referenceIntervals[LOINC]['specimen'][specimen]:
                for i in range(len(referenceIntervals[LOINC]['specimen'][specimen][sex])):
                    minAge = referenceIntervals[LOINC]['specimen'][specimen][sex][i]['minAge']
                    try:
                        referenceIntervals[LOINC]['specimen'][specimen][sex][i]['minAge'] = isodate.parse_duration(minAge)
                    except Exception as e:
                        logger.fatal(f"Error processing reference interval for LOINC {LOINC}, specimen {specimen}, sex {sex}, reference interval {referenceIntervals[LOINC]['specimen'][specimen][sex][i]}, minAge {minAge}: {e}")
                        logging.shutdown()
                        sys.exit(EX_CONFIG)
                    maxAge = referenceIntervals[LOINC]['specimen'][specimen][sex][i]['maxAge']
                    try:
                        referenceIntervals[LOINC]['specimen'][specimen][sex][i]['maxAge'] = isodate.parse_duration(maxAge)
                    except Exception as e:
                        logger.fatal(f"Error processing reference interval for LOINC {LOINC}, specimen {specimen}, sex {sex}, reference interval {referenceIntervals[LOINC]['specimen'][specimen][sex][i]}, maxAge {maxAge}: {e}")
                        logging.shutdown()
                        sys.exit(EX_CONFIG)
    for LOINC in referenceCodes:
        if LOINC in referenceIntervals:
            logger.fatal(f"LOINC {LOINC} is present in both referenceIntervals and referenceCodes")
            logging.shutdown()
            sys.exit(EX_CONFIG)
        referenceCodes[LOINC]['type'] = 'Codes'
        for specimen in referenceCodes[LOINC]['specimen']:
            for sex in referenceCodes[LOINC]['specimen'][specimen]:
                for i in range(len(referenceCodes[LOINC]['specimen'][specimen][sex])):
                    minAge = referenceCodes[LOINC]['specimen'][specimen][sex][i]['minAge']
                    try:
                        referenceCodes[LOINC]['specimen'][specimen][sex][i]['minAge'] = isodate.parse_duration(minAge)
                    except Exception as e:
                        logger.fatal(f"Error processing reference interval for LOINC {LOINC}, specimen {specimen}, sex {sex}, reference codes {referenceCodes[LOINC]['specimen'][specimen][sex][i]}, minAge {minAge}: {e}")
                        logging.shutdown()
                        sys.exit(EX_CONFIG)
                    maxAge = referenceCodes[LOINC]['specimen'][specimen][sex][i]['maxAge']
                    try:
                        referenceCodes[LOINC]['specimen'][specimen][sex][i]['maxAge'] = isodate.parse_duration(maxAge)
                    except Exception as e:
                        logger.fatal(f"Error processing reference interval for LOINC {LOINC}, specimen {specimen}, sex {sex}, reference codes {referenceCodes[LOINC]['specimen'][specimen][sex][i]}, maxAge {maxAge}: {e}")
                        logging.shutdown()
                        sys.exit(EX_CONFIG)
    referenceData = referenceIntervals | referenceCodes

    # Then SPIA stuff
    wb = load_workbook(filename=os.path.join(dataDir, 'RCPA SPIA Preferred Units Table December 2025.xlsx'))
    ws = wb['Preferred Units_Dec 2025']
    rows = ws.iter_rows(values_only=True)
    headers = next(rows)
    preferredUnits = [dict(zip(headers, row)) for row in rows]
    UCUM = {}
    UCUMunit = {}
    for row in preferredUnits:
        if row['Description'] is None:
            break
        if row['UCUM Unit'] in UCUM:
            logger.fatal(f"UCUM unit {row['UCUM Unit']} is duplicated in the preferred units table")
            logging.shutdown()
            sys.exit(EX_CONFIG)
        if row['Preferred Display '] != 'no unit':
            UCUM[row['UCUM Unit']] = row['Preferred Display ']
            UCUMunit[row['Preferred Display ']] = row['UCUM Unit']
        else:
            UCUM[row['UCUM Unit']] = ''
    wb = load_workbook(filename=os.path.join(dataDir, 'RCPA SPIA Core Data.xlsx'))
    ws = wb['Sheet1']
    rows = ws.iter_rows(values_only=True)
    headers = next(rows)
    spiaData = [dict(zip(headers, row)) for row in rows]
    spia = {}
    for row in spiaData:
        LOINC = row['LOINC']
        if LOINC is None:
            break
        if (LOINC not in referenceCodes) and (LOINC not in referenceIntervals):
            continue
        if row['UCUM'] is not None:
            if (row['UCUM'] != 'No unit') and (row['UCUM'] not in UCUM):
                logger.fatal(f"UCUM unit {row['UCUM']} for LOINC {LOINC} in file {os.path.join(dataDir, 'RCPA SPIA Core Data.xlsx')} is not in the preferred units table")
                logging.shutdown()
                sys.exit(EX_CONFIG)
        elif (row['Unit'] is not None) and (row['Unit'] != 'No unit') and (row['Unit'] not in UCUMunit):
            logger.fatal(f"Unit {row['Unit']} for LOINC {LOINC} in file {os.path.join(dataDir, 'RCPA SPIA Core Data.xlsx')} is not in the preferred units table")
            logging.shutdown()
            sys.exit(EX_CONFIG)
        spia[LOINC] = [row['RCPA SPIA Preferred term'], row['Unit'], row['UCUM']]
    for LOINC in referenceData:
        if LOINC not in spia:
            logger.fatal(f"LOINC {LOINC} is present in reference{referenceData[LOINC]['type']} but not in SPIA")
            logging.shutdown()
            sys.exit(EX_CONFIG)

    # And finally the orders and assocated tests
    wb = load_workbook(filename=os.path.join(dataDir, 'Order-Test Sets.xlsx'))
    ws = wb['OrderObs']
    rows = ws.iter_rows(values_only=True)
    headers = next(rows)
    orderTestSets = [dict(zip(headers, row)) for row in rows]
    orderObs = {}
    thisOrder = None
    for row in orderTestSets:
        if row['LOINC'] is None:
            break
        if row['SNOMED CT'] is not None:
            if thisOrder is not None:
                buildOrderObs(thisOrder)        # Build orderObs for the previous order code
            thisOrder = str(row['SNOMED CT']).strip()
            orderObs[thisOrder] = {}
            orderObs[thisOrder]['Description'] = row['Description']
            thisDiscipline = row['Discipline']
            if thisDiscipline not in disciplineMapping:
                logger.fatal(f'Encountered an unknown Discipline "{thisDiscipline}" for order {thisOrder} - file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
                logging.shutdown()
                sys.exit(EX_CONFIG)
            if disciplineMapping[thisDiscipline] not in hl7Table0074:
                logger.fatal(f'Encountered an unmapped Discipline code "{disciplineMapping[thisDiscipline]}" for order {thisOrder} - file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
                logging.shutdown()
                sys.exit(EX_CONFIG)
            orderObs[thisOrder]['Discipline'] = thisDiscipline
            orderObs[thisOrder]['Specimen'] = {}
            orderSpecimenList = str(row['Order Specimen']).split(';')
            for i in range(len(orderSpecimenList) -1, -1, -1):
                orderSpecimenList[i] = orderSpecimenList[i].strip()
                if orderSpecimenList[i] == '':
                    del orderSpecimenList[i]
            orderObs[thisOrder]['orderSpecimenList'] = orderSpecimenList
            orderObs[thisOrder]['Tests'] = {}
        elif thisOrder is None:
            logger.fatal(f'Encountered a row with no SNOMED CT at start of file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
            logging.shutdown()
            sys.exit(EX_CONFIG)
        testCode = row['LOINC'].strip()
        if testCode not in referenceData:
            logger.fatal(f'Encountered a test {testCode} for order {thisOrder} with no reference data - file {os.path.join(dataDir, "Order-Test Sets.xlsx")}')
            logging.shutdown()
            sys.exit(EX_CONFIG)
        testSpecimenList = str(row['Test Specimen']).split(';')
        for i in range(len(testSpecimenList) -1, -1, -1):
            testSpecimenList[i] = testSpecimenList[i].strip()
            if testSpecimenList[i] == '':
                del testSpecimenList[i]
        orderObs[thisOrder]['Tests'][testCode] = testSpecimenList
    buildOrderObs(thisOrder)        # Build orderObs for the last order code


    # Now make some messages
    segments = []
    if requestCode is not None:
        requestCodes = [requestCode]
    else:
        if allRequests:
            requestCodes = list(orderObs.keys())
        elif numberOfMessages > 0:
            requestCodes = random.sample(list(orderObs.keys()), k=numberOfMessages)
        else:
            requestCodes = random.sample(list(orderObs.keys()), k=10)
        remainingRequestCodes = set(orderObs.keys()) - set(requestCodes)
        Fields = ['FHS']
        Fields.append(f'{compSep}{repSep}{escapeChar}{subCompSep}')
        segments.append(fieldSep.join(Fields))
        Fields = ['BHS']
        Fields.append(f'{compSep}{repSep}{escapeChar}{subCompSep}')
        segments.append(fieldSep.join(Fields))
    batchMessages = 0
    print(requestCodes)
    for eachOrderCode in requestCodes:
        orderCode = eachOrderCode
        patientFound = False
        patientCount = 0
        while not patientFound and patientCount < 20:
            # We may need to try multiple patients in case we select a male  patient but randomly select female only tests
            # or select a child and randomly select adult only tests
            patientCount += 1
            thisLocalHospitalHPIO = random.choice(list(clinics.keys()))       # local hospital HPI-O
            thisClinic = random.choice(clinics[thisLocalHospitalHPIO])        # returns the [clinic HPI-O, clinic name]
            thisLab = random.choice(list(Labs.keys()))              # returns the NATA Accreditation Number of the Lab
            thisGP = random.choice(GPs[thisClinic[0]])                 # returns a list of [npi, XCN] for the GP
            thisGPnpi = thisGP[0]
            thisGPxcn = thisGP[1]
            if thisLocalHospitalHPIO in specialistServices:
                thisSpecialistService = random.choice(specialistServices[thisLocalHospitalHPIO])
                thisSpecialistxcn = random.choice(specialists[thisSpecialistService])          # returns a XCN for the Specialist
            else:
                thisSpecialistxcn = ''
            thisLabStaff = random.choice(labStaff)          # returns a list of [npi, XCN] for the Lab Staff
            thisLabStaffxcn = thisLabStaff[0]
            thisLabStaffcn = thisLabStaff[1]
            thisPatient = random.choice(GPpatients[thisGPnpi])      # returns the IHI of a patient associated with this GP
            thisPID = patientPID[thisPatient]                       # The patient's PID segment
            PIDfields = thisPID.split(fieldSep)
            patientBirthdate = dt.strptime(PIDfields[7][:8], '%Y%m%d').date()
            patientAge = (dt.today().date() - patientBirthdate)
            patientSex = PIDfields[8]
            PIDfields[7] = PIDfields[7][0:8] + '+1000'
            if PIDfields[10] in hl7Table0005:
                PIDfields[10] = f'{PIDfields[10]}{compSep}{hl7Table0005[PIDfields[10]]}{compSep}{"HL70005"}'
            if PIDfields[16] in hl7Table0002:
                PIDfields[16] = f'{PIDfields[16]}{compSep}{hl7Table0002[PIDfields[16]]}{compSep}{"HL70002"}'
            newPID = fieldSep.join(PIDfields)
            if (requestCode is not None) or allRequests:
                noOfOrders = 1 
            else:
                noOfOrders = random.randint(1, 3)                       # One to three orders
            OBRs = []
            OBXs = []
            for obrIndex in range(noOfOrders):
                if obrIndex > 0:
                    orderCode = random.choice(list(remainingRequestCodes))
                    remainingRequestCodes.remove(orderCode)
                OBXs.append([])
                orderCount = 0
                while obrIndex == len(OBRs) and (orderCount < 20):
                    # We may need to try multiple orders in case we select an order that is not suitable for this patient
                    orderCount += 1
                    # Pick the specimen with most number of tests
                    orderSpecimen = random.choice(list(orderObs[orderCode]['Specimen']))
                    testSpecimen = random.choice(list(orderObs[orderCode]['Specimen'][orderSpecimen]))
                    for testCode in orderObs[orderCode]['Specimen'][orderSpecimen][testSpecimen]:
                        refDataSpecimen = orderObs[orderCode]['Specimen'][orderSpecimen][testSpecimen][testCode]
                        testSex = None
                        if 'all' in referenceData[testCode]['specimen'][refDataSpecimen]:
                            testSex = 'all'
                        elif patientSex in referenceData[testCode]['specimen'][refDataSpecimen]:
                            testSex = patientSex
                        if testSex is None:
                            continue
                        for i in range(len(referenceData[testCode]['specimen'][refDataSpecimen][testSex])):
                            startDate = referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['minAge']
                            if isinstance(startDate, isodate.duration.Duration):
                                startDate = startDate.totimedelta(start=patientBirthdate)
                            endDate = referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['maxAge']
                            if isinstance(endDate, isodate.duration.Duration):
                                endDate = endDate.totimedelta(start=patientBirthdate)
                            if startDate > patientAge:
                                continue
                            if endDate <= patientAge:
                                continue
                            if referenceData[testCode]['type'] == 'Intervals':
                                minValue = referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['minValue']
                                maxValue = referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['maxValue']
                                precision = referenceData[testCode]['precision']
                                value = round(random.uniform(minValue, maxValue), precision)
                                if len(OBRs) == obrIndex:
                                    OBRs.append([orderCode, orderSpecimen])
                                OBXs[obrIndex].append(['NM', testCode, f'{value:.{precision}f}', 'N', f'{minValue:.{precision}f}-{maxValue:.{precision}f}'])
                            else:
                                probability = random.random()*100
                                for j in range(len(referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['codes'])):
                                    if probability <= referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['codes'][j][1]:
                                        value = referenceData[testCode]['specimen'][refDataSpecimen][testSex][i]['codes'][j][0]
                                        if len(OBRs) == obrIndex:
                                            OBRs.append([orderCode, orderSpecimen])
                                        if value == referenceData[testCode]['normal']:
                                            OBXs[obrIndex].append(['ST', testCode, value, 'N', ''])
                                        else:
                                            OBXs[obrIndex].append(['ST', testCode, value, 'A', ''])
                                        break
                    if len(OBXs[obrIndex]) == 0:
                        continue
                    patientFound = True
                    continue
        if patientFound:
            segments.append(mkMSH(thisLab, thisLocalHospitalHPIO, thisClinic, 'ORM'))       # Start with an Order Message
            segments.append(newPID)                                # Add the patient's PID segment for the ORM message
            Fields = ["PV1", "1","N"]
            PV1seg = fieldSep.join(Fields)
            segments.append(PV1seg)
            for obrIndex in range(len(OBRs)):
                Fields = ["ORC", "NW"]
                ORCseg = fieldSep.join(Fields)
                segments.append(ORCseg)
                segments.append(mkOBR('ORM', obrIndex, thisClinic, thisLab))                # Add the OBR segment for this order
            if requestCode is not None:
                filename = f'ORM_{dt.now().strftime("%Y%m%d%H%M%S")}.hl7'
                with open(os.path.join(outputDir, filename), 'w', newline='\r') as f:
                    for seg in segments:
                        print(seg, file=f)
                segments = []
            segments.append(mkMSH(thisLab, thisLocalHospitalHPIO, thisClinic, 'ORU'))       # Followed by an Observation Result Message
            segments.append(newPID)                                # Add the patient's PID segment for the ORU message
            segments.append(PV1seg)                                 # Add the PV1 segment
            for obrIndex in range(len(OBRs)):
                segments.append(mkOBR('ORU', obrIndex, thisClinic, thisLab))                 # Add the OBR segment for these results
                for obxIndex in range(len(OBXs[obrIndex])):
                    segments.append(mkOBX(obrIndex, obxIndex))              # Add each OBX segment for this result
                segments.append(mkePIT(obrIndex))                       # Add the display segment
            if requestCode is not None:
                filename = f'ORU_{dt.now().strftime("%Y%m%d%H%M%S")}.hl7'
                with open(os.path.join(outputDir, filename), 'w', newline='\r') as f:
                    for seg in segments:
                        print(seg, file=f)
                sys.exit(EX_OK)
            else:
                batchMessages += 1
    Fields = ['BTS']
    Fields.append(str(batchMessages * 2))
    segments.append(fieldSep.join(Fields))
    Fields = ['FTS']
    Fields.append('1')
    segments.append(fieldSep.join(Fields))
    filename = f'Batch_{dt.now().strftime("%Y%m%d%H%M%S")}.hl7'
    with open(os.path.join(outputDir, filename), 'w', newline='\r') as f:
        for seg in segments:
            print(seg, file=f)
    sys.exit(EX_OK)