# === IMPORTS ==================================================================

import enum
from io import FileIO
import os
import shot
import string
import typing
import xlsxwriter


# === GLOBAL CONSTANTS =========================================================

DEFAULT_FILE_NAME = "shotParser"

RANGE = [ -4, -3, -2, -1, 0, 1, 2, 3, 4 ]
RANGE_LENGTH = len(RANGE)


# === ENUM =====================================================================

class Row(enum.Enum):
    Header = 0
    Min = 1
    Max = 2
    Ave = 3
    AbsMin = 4
    AbsMax = 5
    AbsAve = 6
    HeaderRepeat = 7
    Data = 8
    
class VectorOffset(enum.Enum):
    Magnitude = 0
    Index = 1
    X = 2
    Y = 3
    Z = 4
VECTOR_OFFSET_LENGTH = len(VectorOffset)
    
class Col(enum.Enum):
    Name = 0
    Samples = Name + 1
    V = Samples + 1
    VRange = V + VECTOR_OFFSET_LENGTH + 1
    X = VRange + RANGE_LENGTH + 1
    Y = X + VECTOR_OFFSET_LENGTH + 1
    Z = Y + VECTOR_OFFSET_LENGTH + 1
    Shot = Z + VECTOR_OFFSET_LENGTH + 1
    ShotConfidence = Shot + VECTOR_OFFSET_LENGTH
    ShotRange = 1 + ShotConfidence + 1
    AltShot = ShotRange + RANGE_LENGTH + 1
    AltShotConfidence = AltShot + VECTOR_OFFSET_LENGTH
    AltShotRange = 1 + AltShotConfidence + 1


# === CLASSES ==================================================================

class xlsx:
    __EXTENSION = '.xlsx'
    
    __ALL_SHEET = 'ALL'
    
    __DEFAULT_COLUMN_WIDTH = 20
    
    __DEFAULT_SHEET_NAMES = (
        'No Shot',
        'Very Low',
        'Low',
        'Medium',
        'High',
        'Very High',
    )
    
    __ROW_LABELS : typing.List[str] = (
        'NAME',
        'MIN',
        'MAX',
        'AVE',
        '|MIN|',
        '|MAX|',
        '|AVE|',
        'NAME',
    )
    __ROW_LABELS_LENGTH = len(__ROW_LABELS)
    
    __ROW_FORMULAS : typing.List[str] = (
        '',
        '{{=MIN({0}${1}:{0}${2})}}',
        '{{=MAX({0}${1}:{0}${2})}}',
        '{{=AVERAGE({0}${1}:{0}${2})}}',
        '{{=MIN(ABS({0}${1}:{0}${2}))}}',
        '{{=MAX(ABS({0}${1}:{0}${2}))}}',
        '{{=AVERAGE(ABS({0}${1}:{0}${2}))}}',
        '',
    )
    __ROW_FORMULAS_LENGTH = len(__ROW_FORMULAS)
    
    __HEADER_LABELS : typing.List[str] = (
        '',
        'Samples',
        'V',
        'V[]',
        'V-x',
        'V-y',
        'V-z',
        '',
        'V[-4]',
        'V[-3]',
        'V[-2]',
        'V[-1]',
        'V[ 0]',
        'V[+1]',
        'V[+2]',
        'V[+3]',
        'V[+4]',
        '',
        '|X|',
        '|X|[]',
        '|X|-x',
        '|X|-y',
        '|X|-z',
        '',
        '|Y|',
        '|Y|[]',
        '|Y|-x',
        '|Y|-y',
        '|Y|-z',
        '',
        '|Z|',
        '|Z|[]',
        '|Z|-x',
        '|Z|-y',
        '|Z|-z',
        '',
        'Shot',
        'Shot[]',
        'Shot-x',
        'Shot-y',
        'Shot-z',
        'Confidence',
        '',
        'Shot[-4]',
        'Shot[-3]',
        'Shot[-2]',
        'Shot[-1]',
        'Shot[ 0]',
        'Shot[+1]',
        'Shot[+2]',
        'Shot[+3]',
        'Shot[+4]',
        '',
        'Alt',
        'Alt[]',
        'Alt-x',
        'Alt-y',
        'Alt-z',
        'Confidence',
        '',
        'Alt[-4]',
        'Alt[-3]',
        'Alt[-2]',
        'Alt[-1]',
        'Alt[ 0]',
        'Alt[+1]',
        'Alt[+2]',
        'Alt[+3]',
        'Alt[+4]',
    )
    __HEADER_LABELS_LENGTH = len(__HEADER_LABELS)
    
    __DATA_ROW_START = Row.Data.value + 1
    __DATA_ROW_END = 2000
        
    class sheet:
        def __init__(self, name: str, ws : xlsxwriter.Workbook.worksheet_class):
            self.name : str = name
            self.row : int = 0
            self.ws: xlsxwriter.Workbook.worksheet_class = ws
        
    def __init__(self, fileName: str = DEFAULT_FILE_NAME, sheetNames: typing.List[str] = __DEFAULT_SHEET_NAMES):
        self.fileName: str = str(fileName)
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.fileName + self.__EXTENSION)
        self.rankedSheets: typing.List[self.sheet] = []
        self.allSheet: self.sheet = self.sheet(self.__ALL_SHEET, self.wb.add_worksheet(self.__ALL_SHEET))
        self.__initSheet(self.allSheet)
        for name in sheetNames:
            s = self.sheet(name, self.wb.add_worksheet(name))
            self.__initSheet(s)
            self.rankedSheets.append(s)

    def __initSheet(self, s : sheet):
        MAX_ROW = 2000
        for i, field in enumerate(self.__HEADER_LABELS):
            s.ws.write(Row.Header.value, i, field)
            s.ws.write(Row.HeaderRepeat.value, i, field)
        for i, field in enumerate(self.__ROW_LABELS):
            s.ws.write(i, Col.Name.value, self.__ROW_LABELS[i])
        # Set the column width.
        s.ws.set_column(Row.Header.value, Row.Header.value, self.__DEFAULT_COLUMN_WIDTH)
        # Freeze the header rows and columns.
        s.ws.freeze_panes(Row.Data.value, Col.Samples.value)
        s.row = Row.Data.value
            
    def __writeVectorDatum(self, ws : xlsxwriter.Workbook.worksheet_class, row : int, col :int, datum : shot.vectorDatum, ) -> int:
        ws.write(row, col + VectorOffset.Magnitude.value, datum.v.magnitude)
        ws.write(row, col + VectorOffset.Index.value, datum.index)
        ws.write(row, col + VectorOffset.X.value, datum.v.x)
        ws.write(row, col + VectorOffset.Y.value, datum.v.y)
        ws.write(row, col + VectorOffset.Z.value, datum.v.z)
        return col + VECTOR_OFFSET_LENGTH
    
    def __writeRange(self, ws : xlsxwriter.Workbook.worksheet_class, row : int, col :int, accel : typing.List[shot.vector], index : int) -> int:
        samples = len(accel)
        for i, j in enumerate(RANGE):
            j += index
            if (j >= 0) and (j < samples):
                ws.write(row, col + i, accel[j].magnitude)
        return col + RANGE_LENGTH
    
    def __writeShotData(self, s : sheet, data : shot.data):
        row: int = s.row
        ws : xlsxwriter.Workbook.worksheet_class = s.ws
        ws.write(row, Col.Name.value, data.name)
        ws.write(row, Col.Samples.value, len(data.accel))
        self.__writeVectorDatum(ws, row, Col.V.value, data.maxAccel)
        self.__writeRange(ws, row, Col.VRange.value, data.accel, data.maxAccel.index)
        self.__writeVectorDatum(ws, row, Col.X.value, data.maxAccelX)
        self.__writeVectorDatum(ws, row, Col.Y.value, data.maxAccelY)
        self.__writeVectorDatum(ws, row, Col.Z.value, data.maxAccelZ)
        self.__writeVectorDatum(ws, row, Col.Shot.value, data.shot.datum)
        ws.write(row, Col.ShotConfidence.value, data.shot.confidence.value)
        self.__writeRange(ws, row, Col.ShotRange.value, data.accel, data.shot.datum.index)
        self.__writeVectorDatum(ws, row, Col.AltShot.value, data.altShot.datum)
        ws.write(row, Col.AltShotConfidence.value, data.altShot.confidence.value)
        self.__writeRange(ws, row, Col.AltShotRange.value, data.accel, data.altShot.datum.index)
        s.row += 1
        
    def __getXlsxColStr(self, col : int) -> str:
        NUM_LETTERS  : int = len(string.ascii_uppercase)
        pre : int = int(col / NUM_LETTERS)
        post : int = int(col % NUM_LETTERS)
        preChar : str = ''
        if (pre > NUM_LETTERS):
            pre = NUM_LETTERS
        if (pre > 0):
            preChar = string.ascii_uppercase[pre - 1]
        postChar : str = string.ascii_uppercase[post]
        return preChar + postChar
    
    def __writeStatistics(self, s : sheet):
        if s.row > Row.Data.value:
            for i, field in enumerate(self.__HEADER_LABELS):
                if field:
                    for j in range(Row.Min.value, Row.HeaderRepeat.value):
                        colStr : str = self.__getXlsxColStr(i)
                        s.ws.write_array_formula(j, i, j, i, self.__ROW_FORMULAS[j].format(colStr, self.__DATA_ROW_START, s.row))
 
    
    def writeShotData(self, data : shot.data):
        self.__writeShotData(self.rankedSheets[data.shot.confidence.value], data)
        self.__writeShotData(self.allSheet, data)
        
    def finalize(self):
        self.__writeStatistics(self.allSheet)
        for s in self.rankedSheets:
            self.__writeStatistics(s)
        self.wb.close()
        
        
class log:
    __EXTENSION : str = 'csv'
    __OPEN_MODE : str = 'w'
    
    def __init__(self, name : str, folderPath : str):
        self.name : str = name
        self.folderPath : str = folderPath
        self.file : FileIO = open(os.path.join(self.folderPath, '{0}.{1}'.format(self.name, self.__EXTENSION)), self.__OPEN_MODE)
        
    def logAccel(self, data : shot.data, start : int, end : int):
        for s in data.accel[start:end]:
            self.file.write('{0}\n'.format(s.accelEntryString()))
        gyroEnd = end
        if end == len(data.accel):
            gyroEnd =  len(data.gyro)
        for s in data.gyro[start:gyroEnd]:
            self.file.write('{0}\n'.format(s.gyroEntryString()))
        for s in data.hiG:
            self.file.write('{0}\n'.format(s.hiGEntryString()))
        self.file.write('{0}\n'.format(data.calibration.calibEntryString()))
        if data.handedness == shot.Handedness.Left:
            self.file.write('{0}\n'.format(shot.vector.leftHandEntryString()))
        else:
            self.file.write('{0}\n'.format(shot.vector.rightHandEntryString()))
        
    def finalize(self):
        self.file.close()
    