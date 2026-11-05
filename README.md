# mkPathology-HL7-Messages-Australia

Create HL7 v2.4 Pathology Order and Results messages from SPIA order and result codes

Uses a synthetic Australian healthcare population (data/healthPopulation.xlsx) for GPs, patients, copy to specialist and lab staff.

Order-Test Sets.xlsx is a cleaned up version of the SPIA Top 100 requesting procedures; 81 of which were matched with SPIA LOINC result code(s).

RCPA SPIA Core Data.xlsx is a cleaned up version of a concatination of the SPIA Information Model and Reference Sets workbooks.

referenceintervals.json are upper and lower bounds for numeric SPIA LOINC code. **mkHL7LabMessages.py** only creates normal results for numeric SPIA LOINC codes.

referenceCode.json are population distributions for non-numeric LOINC codes; codes where the result is a code such as "Positive" or "Negative". **mkHL7LabMessages.py** will create some abnormal result messages.

## SYNOPSIS

$ python mkHL7LabMessages.py [-D dataDir|--dataDir=dataDir]  
                             [-P profile|--profile=profile]  
                             [-S|--useSnomedSpecimenCodes]  
                             [-R requestCode|--requestCode=requestCode]  
                             [-A|--All]  
                             [-N numberOfMessages|--numberOfMessages=numberOfMessages]  
                             [-O outputDir|--outputDir=outputDir]  
                             [-v loggingLevel|--verbose=loggingLevel]  
                             [-L logDir|--logDir=logDir] [-l logfile|--logfile=logfile]  

## OPTIONS

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
Generate HL7 messages for all SPIA request codes (default=False)

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

## Output

Where a specific request code is specified two files will be careated (one order message and a matching results message).

In all other scenarios a batch of messages will be created, being an order message followed by the matching result(s) message.
