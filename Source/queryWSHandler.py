from Source.queryObj import QueryObj


def getQueryData(patientDict, queryWS, outputWS):
    # insert header
    for column in range(1, 13):
        outputWS.cell(column=column, row=1,
                      value=queryWS.cell(column=column, row=1).value)

    outputWS.cell(column=13, row=1, value=queryWS.cell(column=21, row=1).value)

    # form list of id's to be sorted
    patientList = []

    for row in range(2, queryWS.max_row + 1):
        patientID = queryWS.cell(column=4, row=row).value
        print("Checking patient for query list: " + str(patientID))
        if int(patientID) in patientDict:
            print("    Adding patient for query list: " + str(patientID))
            # grab data and store into obj
            columnA = queryWS.cell(column=1, row=row).value
            columnB = queryWS.cell(column=2, row=row).value
            columnC = queryWS.cell(column=3, row=row).value
            columnD = queryWS.cell(column=4, row=row).value
            columnE = queryWS.cell(column=5, row=row).value
            columnF = queryWS.cell(column=6, row=row).value
            columnG = queryWS.cell(column=7, row=row).value
            columnH = queryWS.cell(column=8, row=row).value
            columnI = queryWS.cell(column=9, row=row).value
            columnJ = queryWS.cell(column=10, row=row).value
            columnK = queryWS.cell(column=11, row=row).value
            columnL = queryWS.cell(column=12, row=row).value
            columnM = queryWS.cell(column=21, row=row).value

            patient = QueryObj(columnA, columnB, columnC, columnD, columnE, columnF,
                               columnG, columnH, columnI, columnJ, columnK, columnL, columnM)
            patientList.append(patient)

    # split list into two based on column i value (missing, entered)
    missingList = []
    enteredList = []
    notApplicableList = []
    for patient in patientList:
        if patient.columnI == 'Missing':
            missingList.append(patient)
        elif patient.columnI == 'Entered':
            enteredList.append(patient)
        elif patient.columnI == 'Not applicable':
            notApplicableList.append(patient)
        else:
            print('Query Worksheet: Patient ' + str(patient.columnD) +
                  ' has an invalid Field Status: ' + str(patient.columnI))

    # sort by column L and store to
    sortedMissingList = sorted(missingList, key=lambda x: x.columnL)
    sortedEnteredList = sorted(enteredList, key=lambda x: x.columnL)
    sortedNotApplicableList = sorted(
        notApplicableList, key=lambda x: x.columnL)
    fullList = sortedMissingList + sortedEnteredList + sortedNotApplicableList

    # enter data to spread sheet
    for index, patient in enumerate(fullList):
        # offset index by 2 due to 0 indexing + header
        outputWS.cell(column=1, row=index + 2, value=patient.columnA)
        outputWS.cell(column=2, row=index + 2, value=patient.columnB)
        outputWS.cell(column=3, row=index + 2, value=patient.columnC)
        outputWS.cell(column=4, row=index + 2, value=patient.columnD)
        outputWS.cell(column=5, row=index + 2, value=patient.columnE)
        outputWS.cell(column=6, row=index + 2, value=patient.columnF)
        outputWS.cell(column=7, row=index + 2, value=patient.columnG)
        outputWS.cell(column=8, row=index + 2, value=patient.columnH)
        outputWS.cell(column=9, row=index + 2, value=patient.columnI)
        outputWS.cell(column=10, row=index + 2, value=patient.columnJ)
        outputWS.cell(column=11, row=index + 2, value=patient.columnK)
        outputWS.cell(column=12, row=index + 2, value=patient.columnL)
        outputWS.cell(column=13, row=index + 2, value=patient.columnM)
        print("Adding query to query worksheet: " + str(patient.columnA))
