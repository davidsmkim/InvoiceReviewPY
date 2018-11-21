class PatientObj:

    def __init__(self, patientID, cancerStatus, cellLocation):
        self.patientPeriod = []
        self.patientID = patientID
        self.cancerStatus = cancerStatus
        self.cellLocation = cellLocation
        self.idLocations = {}  # timeperiod : cell location
        self.numTimesPaid = None
        self.numPaymentRequested = None
