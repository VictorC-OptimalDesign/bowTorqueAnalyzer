# === IMPORTS ==================================================================
import enum
import glob
import math
import operator
import os
import shot
import shotOutput
import shotPlot
import shutil
import string
import typing
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================

DATA_FOLDER = '_DATA'

# === FUNCTIONS ================================================================

def __initRawDataLog() -> typing.List[shotOutput.xlsxData]:
    GYRO_NAME = 'gyro'
    ACCEL_NAME = 'accel'
    HIG_NAME = 'hiG'
    path : str = os.path.join(os.getcwd(), DATA_FOLDER)
    try:
        os.makedirs(path)
    except:
        pass
    NAME_LUT = {
        shotOutput.xlsxData.DataType.Gyro : GYRO_NAME,
        shotOutput.xlsxData.DataType.Accel : ACCEL_NAME,
        shotOutput.xlsxData.DataType.HiG : HIG_NAME
        }
    logs : typing.List[shotOutput.xlsxData] = []
    for t in shotOutput.xlsxData.DataType:
        log : shotOutput.xlsxData = shotOutput.xlsxData(NAME_LUT[t], path, t)
        logs.append(log)
    return logs

def __getShotRange(i : int, length : int) -> typing.List[int]:
    MINUS: int = -5
    PLUS: int = 5
    start: int = i + MINUS
    end: int = i + PLUS
    if (start < 0):
        start = 0
    if (end > length):
        end = length
    return [start, end]

def __plot(data : shot.data):
    range: typing.List[int] = __getShotRange(data.hiGShot.datum.index, len(data.hiG))
    shotPlot.vector_plot(data.getHiGList(range[0], range[1]))
    range = __getShotRange(data.shot.datum.index, len(data.accel))
    shotPlot.vector_plot(data.getAccelList(range[0], range[1]))

def __process():
    output : shotOutput.xlsx = shotOutput.xlsx(shotOutput.xlsx.Mode.Abbreviated)
    logs : typing.List[shotOutput.xlsxData] = __initRawDataLog()
    allLog : shotOutput.xlsxAllData = shotOutput.xlsxAllData('all', DATA_FOLDER)
    for fileName in glob.glob('*.csv'):
        print('processing {0}...'.format(fileName))
        datum : shot.data = shot.data(fileName)
        confidenceIndex = datum.shot.confidence.value
        output.writeShotData(datum)
        for l in logs:
            l.addData(datum)
        allLog.addData(datum)
        #__plot(datum)
    for l in logs:
        l.finalize()
    allLog.finalize()
    output.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    __process()
else:
    print("ERROR: bowTorqueAnalyzer needs to be the calling python module!")
    