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


# === FUNCTIONS ================================================================

def __initRawDataLog() -> typing.List[shotOutput.xlsxData]:
    DATA_FOLDER = '_DATA'
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

def __plot(data : shot.data):
    shotPlot.vector_plot(data.getHiGList(0, 30))
    shotIndex : int = data.shot.datum.index
    shotPlot.vector_plot(data.getAccelList(shotIndex - 3, shotIndex + 7))

def __process():
    output : shotOutput.xlsx = shotOutput.xlsx(shotOutput.xlsx.Mode.Abbreviated)
    logs : typing.List[shotOutput.xlsxData] = __initRawDataLog()
    for fileName in glob.glob('*.csv'):
        print('processing {0}...'.format(fileName))
        datum : shot.data = shot.data(fileName)
        confidenceIndex = datum.shot.confidence.value
        output.writeShotData(datum)
        for l in logs:
            l.addData(datum)
        __plot(datum)
    for l in logs:
        l.finalize()
    output.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    __process()
else:
    print("ERROR: bowTorqueAnalyzer needs to be the calling python module!")
    