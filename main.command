#!/usr/bin/env python

from openpyxl import Workbook
import datetime
from os.path import dirname

from Source.spreadsheetHandler import verifyWB, verifyWorksheets, cleanFiles, createWorksheet
from Source.inputWSHandler import processInput
from Source.cancerStatusHandler import validateCancerStatus, validateSiteLocations, getSubjectStatus
from Source.invoiceStatusHandler import validateInvoiceStatus
from Source.patientObj import PatientObj
from Source.subjectDataHandler import copySubjectData
from Source.queryWSHandler import getQueryData

print('Started Program.')

# clean up previous files
print('Cleaning files...')
cleanFiles()
print('Cleaned files')

# grab work books after verification
print('Verifying workbooks...')
workbooks = verifyWB()
print('Workbooks verified.')

# verify worksheets
print('Verifying worksheets...')
worksheets = verifyWorksheets(workbooks)
print('Worksheets verified.')

# load workbook objects
print('Loading data...')
queryWS = worksheets[0]
subjectDataWS = worksheets[1]
invoicesWS = worksheets[2]
cancerStatusWS = worksheets[3]
inputWS = worksheets[4]
print('Data loaded.')

# create output workbook
writeWB = Workbook()
writeWS = writeWB.active

# get list of patients
print('Getting patient data...')
patientDict = processInput(inputWS, writeWS)
print('Patient data gathered.')

# check invoice status
print('Validating invoice status...')
validateInvoiceStatus(patientDict, invoicesWS, writeWS)
print('Invoice status validated.')

# check cancer status
print('Validating patient cancer status...')
validateCancerStatus(patientDict,cancerStatusWS, writeWS)
print('Patient cancer status validated.')

# check site locations
print('Gathering site locations...')
validateSiteLocations(patientDict, cancerStatusWS, writeWS)
print('Site locations gathered.')

# get subject status
print('Gathering subject status...')
getSubjectStatus(patientDict, cancerStatusWS, writeWS)
print('Subject status gathered.')

# get subject data
print('Gathering subject data...')
subjectDataOutputWS = createWorksheet(writeWB, 'Subject Data')
copySubjectData(subjectDataWS, subjectDataOutputWS, patientDict)
print('Subject data gathered.')

print('Gathering subject query...')
queryOutputWS = createWorksheet(writeWB, 'Query')
getQueryData(patientDict, queryWS, queryOutputWS)
print('Gathered subject query.')

# save output workbook
# date time
now = datetime.datetime.now()
print('Saving output file')
outputPath = dirname(dirname(__file__)) + '/InvoiceReviewPY/invoice_' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '_LK.xlsx'
writeWB.save(outputPath)
print('Saved output file at: ' + outputPath)
