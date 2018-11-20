from Source.errorLog import writeErrorLog
from Source.errorTypeConstants import errorTypeCancerStatus, errorTypeSiteStatus, errorTypeSubjectStatus
from Source.columnNamesConstants import cancerColumn, siteColumn, subjectStatusColumn

def copySubjectData(subjectDataWS, outputWS, patientDict):
  # insert header
  for column in range(1, 10):
    outputWS.cell(column = column, row = 1,
      value = subjectDataWS.cell(column = column, row = 1).value)

  currentRow = 2 # keep track of output row

  # add data if match
  for row in range(2, subjectDataWS.max_row + 1):
    checkCell = subjectDataWS.cell(column = 2, row = row)
    if int(checkCell.value) in patientDict:
      for column in range(1, 10):
        outputWS.cell(column = column, row = currentRow,
            value = subjectDataWS.cell(column = column, row = row).value)
      currentRow += 1
