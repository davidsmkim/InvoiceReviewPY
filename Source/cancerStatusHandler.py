from Source.errorLog import writeErrorLog
from Source.errorTypeConstants import errorTypeCancerStatus, errorTypeSiteStatus, errorTypeSubjectStatus
from Source.columnNamesConstants import cancerColumn, siteColumn, subjectStatusColumn

# method to compare cancer status from patient list to cancer status worksheet
def validateCancerStatus(patientDict, cancerStatusWS, writeWS):
    # track errors
    errorLogList = []

    # make copy of patient ids to make sure all patient id's are accounted for
    patientList = list(patientDict.keys())

    # iterate through worksheet to find appropriate patient id's and status
    for row in range(2, cancerStatusWS.max_row + 1):
        wsPatientIDCell = cancerStatusWS.cell(column = 3, row = row)

        # found patient id match
        if wsPatientIDCell.value in patientDict:

            # grab cancer status and remove from patientList
            wsPatientCancerStatusCell = cancerStatusWS.cell(column = 6, row = row)
            print('Checking Patient ID: ' + str(wsPatientIDCell.value))
            try:
                patientList.remove(wsPatientIDCell.value)
            except ValueError:
                print('Duplicate ID Found: ' + str(wsPatientIDCell.value))
                writeWS.cell(column = cancerColumn, row = patientDict[wsPatientIDCell.value].cellLocation[1],
                    value = 'Duplicate ID Found: ' + str(patientDict[wsPatientIDCell.value].cancerStatus))
                continue
            print('Matched ID: ' + str(wsPatientIDCell.value))
            # compare cancer status
            if patientDict[wsPatientIDCell.value].cancerStatus.lower() == wsPatientCancerStatusCell.value.lower():
                print('Cancer Status matched for patient ' + str(wsPatientIDCell.value))
            else:
                writeWS.cell(column = cancerColumn, row = patientDict[wsPatientIDCell.value].cellLocation[1],
                    value = wsPatientCancerStatusCell.value)
                print('Cancer Status mismatched for patient ' + str(wsPatientIDCell.value) + ':Found - ' +
                    str(wsPatientCancerStatusCell.value) + ' Expected - ' + 
                    patientDict[wsPatientIDCell.value].cancerStatus)
                errorLogList.append('Cancer Status mismatched for patient ' + str(wsPatientIDCell.value) + ':Found - ' +
                    str(wsPatientCancerStatusCell.value) + ' Expected - ' + 
                    patientDict[wsPatientIDCell.value].cancerStatus)
    
    # verify all patients have been found
    if len(patientList) != 0:
        print('Not all patients have been matched.')
        for patient in patientList:
            print('Patient ID: ' + str(patient) + ' has not been found in the Cancer/Non-cancer workbook')
            writeWS.cell(column = cancerColumn, row = patientDict[patient].cellLocation[1],
            value = 'Patient not found in Cancer/Non Cancer Medrio Report')
            errorLogList.append('Patient ID: ' + str(patient) + ' has not been found in the Cancer/Non-cancer workbook')

    # create/append error log if needed
    writeErrorLog(errorTypeCancerStatus, errorLogList)

def validateSiteLocations(patientDict, cancerStatusWS, writeWS):
    # error log list
    errorLog = []

    # make copy of patient ids to make sure all patient id's are accounted for
    patientList = list(patientDict.keys())

    for row in range(2, cancerStatusWS.max_row + 1):
        patientIDCell = cancerStatusWS.cell(column = 3, row = row)

        # add site location to output worksheet and remove from patient list
        if patientIDCell.value in patientDict:
            siteLocationCell = cancerStatusWS.cell(column = 4, row = row)
            writeWS.cell(column = siteColumn, row = patientDict[patientIDCell.value].cellLocation[1],
            value = siteLocationCell.value)
            try:
                patientList.remove(patientIDCell.value)
            except ValueError:
                writeWS.cell(column = siteColumn, row = patientDict[patientIDCell.value].cellLocation[1],
                value = 'Duplicate ID Found: ' + str(patientIDCell.value))
            
            if siteLocationCell.value is None:
                errorLog.append('Patient ID ' + str(patientIDCell.value) + ' does not have a site location')
                print('Patient ID ' + str(patientIDCell.value) + ' does not have a site location')
    
    # check for left over patients
    if len(patientList) != 0:
        for patient in patientList:
            errorLog.append('Patient ID ' + str(patient) + ' was not found in the Cancer/Non-cancer workbook.')
    
    # call error logging if needed
    writeErrorLog(errorTypeCancerStatus, errorLog)

def getSubjectStatus(patientDict, cancerStatusWS, writeWS):
    # error log list
    errorLog = []

    # make copy of patient ids to make sure all patient id's are accounted for
    patientList = list(patientDict.keys())

    for row in range(2, cancerStatusWS.max_row + 1):
        patientIDCell = cancerStatusWS.cell(column = 3, row = row)

        # add subject status to output worksheet and remove from patient list
        if patientIDCell.value in patientDict:
            subjectStatusCell = cancerStatusWS.cell(column = 9, row = row)
            writeWS.cell(column = subjectStatusColumn, row = patientDict[patientIDCell.value].cellLocation[1],
            value = subjectStatusCell.value)
            try:
                patientList.remove(patientIDCell.value)
            except ValueError:
                writeWS.cell(column = siteColumn, row = patientDict[patientIDCell.value].cellLocation[1],
                value = 'Duplicate ID Found: ' + str(patientIDCell.value))
            
    # check for left over patients
    if len(patientList) != 0:
        for patient in patientList:
            errorLog.append('Patient ID ' + str(patient) + ' was not found in the Cancer/Non-cancer workbook.')
    
    # call error logging if needed
    writeErrorLog(errorTypeCancerStatus, errorLog)