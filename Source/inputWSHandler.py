from Source.patientObj import PatientObj


def processInput(inputWS, writeWS):
    patientDict = {}
    for row in range(2, inputWS.max_row + 1):
        # create patient obj and store in dict
        patientIDCell = inputWS.cell(column=1, row=row)
        patientTimePeriod = inputWS.cell(column=4, row=row)
        patientCancerStatus = inputWS.cell(column=2, row=row).value.lower()
        patientCellLocation = [1, row - 1]
        paymentAmount = inputWS.cell(column=3, row=row).value

        # update patient info in dict if exists or store in dict
        if patientIDCell.value in patientDict:
            patient = patientDict[patientIDCell.value]
            patient.idLocations[patientTimePeriod.value] = patientCellLocation

            print("Duplicate input found for: " + str(patientIDCell.value))

            # check if cancer statuses match
            if patient.cancerStatus != patientCancerStatus:
                patient.cancerStatus = "Error: Cancer/NonCancer"

            # mark duplicates and add new time period to list
            if patientTimePeriod.value in patient.patientPeriod:
                for timePeriod in patientDict[patientIDCell.value].idLocations:
                    writeWS.cell(
                        column=8, row=patientDict[patientIDCell.value].idLocations[timePeriod][1] + 1, value='Duplicate in Input')
            patient.patientPeriod.append(patientTimePeriod.value)

            # mark mismatched cancer status
            if patient.cancerStatus == "Error: Cancer/NonCancer":
                for locations in patientDict[patientIDCell.value].idLocations:
                    writeWS.cell(
                        column=4, row=patientDict[patientIDCell.value].idLocations[locations][1] + 1, value=patient.cancerStatus)

        else:
            patient = PatientObj(
                patientIDCell.value, patientCancerStatus, patientCellLocation)
            patientDict[patientIDCell.value] = patient
            patientDict[patientIDCell.value].idLocations[patientTimePeriod.value] = patientCellLocation
            patient.patientPeriod.append(patientTimePeriod.value)

        # write patientID to output file
        writeWS.cell(column=patientCellLocation[0], row=patientCellLocation[1] + 1,
                     value=patientDict[patientIDCell.value].patientID)
        # write time period to output file
        writeWS.cell(
            column=7, row=patientCellLocation[1] + 1, value=patientTimePeriod.value)

        # write payment amount
        writeWS.cell(
            column=2, row=patientCellLocation[1] + 1, value=paymentAmount)

        print('Patient ID: ' + str(patientIDCell.value) + ' added')

    return patientDict
