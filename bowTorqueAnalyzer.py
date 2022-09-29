# === IMPORTS ==================================================================
import enum
import glob
import math
import operator
import os
import shot
import shotOutput
import shutil
import string
import typing
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================


# === FUNCTIONS ================================================================

def process():
    output = shotOutput.xlsx()
    for fileName in glob.glob('*.csv'):
        print('processing {0}...'.format(fileName))
        datum : shot.data = shot.data(fileName)
        confidenceIndex = datum.shot.confidence.value
        output.writeShotData(datum)
    output.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    process()
else:
    print("ERROR: bowTorqueAnalyzer needs to be the calling python module!")
    