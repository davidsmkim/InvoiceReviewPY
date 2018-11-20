from Source.errorLog import writeErrorLog
from Source.errorTypeConstants import errorTypeInvoiceStatus
from Source.columnNamesConstants import invoiceColumn

def validateInvoiceStatus(patientDict, invoiceStatusWS, writeWS):
    # track errors
    errorLog = []
    patientList = list(patientDict.keys())

    # check if patient already exists in invoice tracker
    for row in range(3, invoiceStatusWS.max_row + 1):
        wsPatientIDCell = invoiceStatusWS.cell(column = 3, row = row)
        if wsPatientIDCell.value in patientDict:
            writeWS.cell(column = invoiceColumn, row = patientDict[wsPatientIDCell.value].cellLocation[1], value = 'Already paid')
            errorLog.append('Patient ID ' + str(wsPatientIDCell.value) + ' already exists in the invoice tracker.')
            print('Patient ID ' + str(wsPatientIDCell.value) + ' already exists in the invoice tracker.')
            patientList.remove(wsPatientIDCell.value)

    if len(patientList) != 0:
        for patient in patientList:
            writeWS.cell(column = invoiceColumn, row = patientDict[patient].cellLocation[1], value = 'Not paid')
            print('Patient ID ' + str(patient) + ' has not been paid.')
    
    # add to error log if exists
    writeErrorLog(errorTypeInvoiceStatus, errorLog)
            
