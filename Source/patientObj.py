class PatientObj:
    patientID = None
    cancerStatus = None
    cellLocation = None

    def __init__(self, patientID, cancerStatus, cellLocation):
        self.patientID = patientID
        self.cancerStatus = cancerStatus
        self.cellLocation = cellLocation
