import sys

from Source.errorLog import writeErrorLog
from Source.errorTypeConstants import errorTypeInvoiceStatus
from Source.columnNamesConstants import invoiceColumn


def validateInvoiceStatus(patientDict, invoiceStatusWS, writeWS):
    # get columns
    medrioIDColumn = None
    enrollmentColumn = None
    year1Column = None
    year2Column = None
    for column in range(1, invoiceStatusWS.max_column + 1):
        if invoiceStatusWS.cell(column=column, row=1).value == "Per Subject Fee  (Overhead Inclusive, at 90{} of total cost at enrollment / Initial Enrollment)".format("%"):
            enrollmentColumn = column
        elif invoiceStatusWS.cell(column=column, row=1).value == "Outcome Data Collection Fee (Y1) (5{} or Per Contract Fee)".format("%"):
            year1Column = column
        elif invoiceStatusWS.cell(column=column, row=1).value == "Outcome Data Collection Fee (Y2) (5{} or Per Contract Fee)".format("%"):
            year2Column = column
        elif invoiceStatusWS.cell(column=column, row=1).value == "Medrio Subject ID":
            medrioIDColumn = column

    errorLog = []
    invoiceSet = set()

    # insert header
    writeWS.cell(column=1, row=1, value="Medrio ID")
    writeWS.cell(column=2, row=1, value="Amount")
    writeWS.cell(column=3, row=1, value="Payment")
    writeWS.cell(column=4, row=1, value="Cancer/Non-Cancer")
    writeWS.cell(column=5, row=1, value="Site")
    writeWS.cell(column=6, row=1, value="Subject Status")
    writeWS.cell(column=7, row=1, value="Time Period")
    writeWS.cell(column=8, row=1, value="Comments")

    # check if patient already exists in invoice tracker
    for row in range(2, invoiceStatusWS.max_row + 1):
        wsPatientIDCell = invoiceStatusWS.cell(column=medrioIDColumn, row=row)

        if wsPatientIDCell.value in patientDict:
            patient = patientDict[wsPatientIDCell.value]

            for timePeriod in patient.patientPeriod:
                # check if id + time period already present
                if str(wsPatientIDCell.value) + timePeriod in invoiceSet:
                    for timePeriod in patient.idLocations:
                        writeWS.cell(column=invoiceColumn,
                                     row=patient.idLocations[timePeriod][1] + 1, value="Duplicate IDs and Time Period in Invoice Tracker")
                else:
                    invoiceSet.add(str(wsPatientIDCell.value) + timePeriod)
                    # check for time period
                    if timePeriod == "Enrollment":
                        timePeriodColumn = enrollmentColumn
                        patient.patientPeriod.remove("Enrollment")
                    elif timePeriod == "Y1":
                        timePeriodColumn = year1Column
                        patient.patientPeriod.remove("Y1")
                    elif timePeriod == "Y2":
                        timePeriodColumn = year2Column
                        patient.patientPeriod.remove("Y2")

                    # mark payment status
                    if invoiceStatusWS.cell(column=timePeriodColumn, row=row).value is not None:
                        writeWS.cell(column=invoiceColumn,
                                     row=patientDict[wsPatientIDCell.value].cellLocation[1] + 1, value='Already paid')
                    else:
                        writeWS.cell(column=invoiceColumn,
                                     row=patientDict[wsPatientIDCell.value].cellLocation[1] + 1, value='Not paid')

    # mark status for ids not present in invoice tracker
    for patient in patientDict:
        for timePeriod in patientDict[patient].patientPeriod:
            writeWS.cell(column=invoiceColumn,
                         row=patientDict[patient].idLocations[timePeriod][1] + 1, value='Not paid')
        print('Patient ID ' + str(patient) + ' has not been paid.')

    # add to error log if exists
    writeErrorLog(errorTypeInvoiceStatus, errorLog)
