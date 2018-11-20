from os.path import dirname
import datetime

# function to compile error log
def writeErrorLog(errorType, errorLogList):
    now = datetime.datetime.now()
    outputPath = dirname(dirname(__file__)) + '/error_log_' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '_LK.txt'
    errorLog = open(outputPath,"a")
    errorLog.write('\n\n*****' + errorType + '*****\n')
    for error in errorLogList:
        errorLog.write(error + '\n')
    errorLog.close()

    