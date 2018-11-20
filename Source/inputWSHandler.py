from Source.patientObj import PatientObj

def processInput(inputWS, writeWS):
    patientDict = {}
    for row in range(2, inputWS.max_row + 1):
        # store patient info into dict
        patientIDCell = inputWS.cell(column = 1, row = row)
        patientCancerStatusCell = inputWS.cell(column = 2, row = row)
        patientCellLocation = [1, row - 1]
        patient = PatientObj(patientIDCell.value, patientCancerStatusCell.value, patientCellLocation)
        patientDict[patient.patientID] = patient

        # write patientID to output file
        writeWS.cell(column = patientDict[patientIDCell.value].cellLocation[0], row = patientDict[patientIDCell.value].cellLocation[1],
        value = patientDict[patientIDCell.value].patientID)

        print('Patient ID: ' + str(patientIDCell.value) + ' added')

    return patientDict
