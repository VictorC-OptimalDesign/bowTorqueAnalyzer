# === IMPORTS ==================================================================

import enum
from io import FileIO
import os
from statistics import NormalDist
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
    HiGShot = AltShotConfidence + RANGE_LENGTH + 1
    HiGShotConfidence = HiGShot + VECTOR_OFFSET_LENGTH
    HiGShotRange = 1 + HiGShotConfidence + 1
    
class AbbreviatedCol(enum.Enum):
    Name = 0
    Samples = Name + 1
    V = Samples + 1
    VRange = V + VECTOR_OFFSET_LENGTH + 1
    Shot = VRange + RANGE_LENGTH + 1
    ShotConfidence = Shot + VECTOR_OFFSET_LENGTH
    ShotRange = 1 + ShotConfidence + 1
    HiGShot = ShotRange + RANGE_LENGTH + 1
    HiGShotConfidence = HiGShot + VECTOR_OFFSET_LENGTH
    HiGShotRange = 1 + HiGShotConfidence + 1


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
        '',
        'HiG',
        'HiG[]',
        'HiG-x',
        'HiG-y',
        'HiG-z',
        'Confidence',
        '',
        'HiG[-4]',
        'HiG[-3]',
        'HiG[-2]',
        'HiG[-1]',
        'HiG[ 0]',
        'HiG[+1]',
        'HiG[+2]',
        'HiG[+3]',
        'HiG[+4]',
    )
    __HEADER_LABELS_LENGTH = len(__HEADER_LABELS)
    
    
    __ABBREVIATED_HEADER_LABELS : typing.List[str] = (
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
        'HiG',
        'HiG[]',
        'HiG-x',
        'HiG-y',
        'HiG-z',
        'Confidence',
        '',
        'HiG[-4]',
        'HiG[-3]',
        'HiG[-2]',
        'HiG[-1]',
        'HiG[ 0]',
        'HiG[+1]',
        'HiG[+2]',
        'HiG[+3]',
        'HiG[+4]',
    )
    __ABBREVIATED_HEADER_LABELS_LENGTH = len(__HEADER_LABELS)
    
    __DATA_ROW_START = Row.Data.value + 1
    __DATA_ROW_END = 2000
    
    class Mode(enum.Enum):
        Normal = 0
        Abbreviated = 1
        
    class sheet:
        def __init__(self, name: str, ws : xlsxwriter.Workbook.worksheet_class):
            self.name : str = name
            self.row : int = 0
            self.ws: xlsxwriter.Workbook.worksheet_class = ws
        
    def __init__(self, mode : Mode = Mode.Normal, fileName: str = DEFAULT_FILE_NAME, sheetNames: typing.List[str] = __DEFAULT_SHEET_NAMES):
        self.mode : self.Mode = mode
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
        labels : typing.List[str] = self.__HEADER_LABELS
        if self.mode is self.Mode.Abbreviated:
            labels = self.__ABBREVIATED_HEADER_LABELS
        for i, field in enumerate(labels):
            s.ws.write(Row.Header.value, i, field)
            s.ws.write(Row.HeaderRepeat.value, i, field)
        for i, field in enumerate(self.__ROW_LABELS):
            s.ws.write(i, Col.Name.value, self.__ROW_LABELS[i])
        # Set the column width.
        s.ws.set_column(Row.Header.value, Row.Header.value, self.__DEFAULT_COLUMN_WIDTH)
        # Freeze the header rows and columns.
        s.ws.freeze_panes(Row.Data.value, Col.Samples.value)
        s.row = Row.Data.value
            
    def __writeVectorDatum(self, ws : xlsxwriter.Workbook.worksheet_class, row : int, col :int, datum : shot.vectorDatum) -> int:
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
        if self.mode is self.Mode.Normal:
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
            self.__writeVectorDatum(ws, row, Col.HiGShot.value, data.hiGShot.datum)
            ws.write(row, Col.HiGShotConfidence.value, data.hiGShot.confidence.value)
            self.__writeRange(ws, row, Col.HiGShotRange.value, data.accel, data.hiGShot.datum.index)
        elif self.mode is self.Mode.Abbreviated:
            ws.write(row, AbbreviatedCol.Name.value, data.name)
            ws.write(row, AbbreviatedCol.Samples.value, len(data.accel))
            self.__writeVectorDatum(ws, row, AbbreviatedCol.V.value, data.maxAccel)
            self.__writeRange(ws, row, AbbreviatedCol.VRange.value, data.accel, data.maxAccel.index)
            self.__writeVectorDatum(ws, row, AbbreviatedCol.Shot.value, data.shot.datum)
            ws.write(row, AbbreviatedCol.ShotConfidence.value, data.shot.confidence.value)
            self.__writeRange(ws, row, AbbreviatedCol.ShotRange.value, data.accel, data.shot.datum.index)
            self.__writeVectorDatum(ws, row, Col.HiGShot.value, data.hiGShot.datum)
            ws.write(row, Col.HiGShotConfidence.value, data.hiGShot.confidence.value)
            self.__writeRange(ws, row, Col.HiGShotRange.value, data.accel, data.hiGShot.datum.index)
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
            labels : typing.List[str] = self.__HEADER_LABELS
            if self.mode is self.Mode.Abbreviated:
                labels = self.__ABBREVIATED_HEADER_LABELS
            for i, field in enumerate(labels):
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
        
class xlsxData:
    __EXTENSION : str = 'xlsx'
    __OPEN_MODE : str = 'w'
    
    class Row(enum.Enum):
        Header = 0
        Data = 1
        
    class Col(enum.Enum):
        Index = 0
        X = 1
        Y = 2
        Z = 3
        Magnitude = 4
        Shot = 5
    
    class DataType(enum.Enum):
        Gyro = 0
        Accel = 1
        HiG = 2
        
    def __init__(self, name: str, folderPath : str, type : DataType = DataType.Gyro):
        self.name: str = str(name)
        self.folderPath : str = folderPath
        self.type : self.DataType = type
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(os.path.join(self.folderPath, '{0}.{1}'.format(self.name, self.__EXTENSION)))
        self.ws: typing.List[xlsxwriter.Workbook.worksheet_class] = []
        
    def __addHeader(self, ws : xlsxwriter.Workbook.worksheet_class):
        ws.write(self.Row.Header.value, self.Col.Index.value, 'Index')
        ws.write(self.Row.Header.value, self.Col.X.value, 'X')
        ws.write(self.Row.Header.value, self.Col.Y.value, 'Y')
        ws.write(self.Row.Header.value, self.Col.Z.value, 'Z')
        ws.write(self.Row.Header.value, self.Col.Magnitude.value, 'Magnitude')
        ws.write(self.Row.Header.value, self.Col.Shot.value, 'Shot')
        
    def addData(self, s : shot.data):
        MAX_VALUE = 10000
        data : typing.List[shot.vector] = s.gyro
        shotIndex = s.shot.datum.index
        if self.type is self.DataType.Accel:
            data = s.accel
        elif self.type is self.DataType.HiG:
            data = s.hiG
            shotIndex = s.hiGShot.datum.index
        ws = self.wb.add_worksheet(s.fileName)
        self.__addHeader(ws)
        for i, d in enumerate(data):
            ws.write(self.Row.Data.value + i, self.Col.Index.value, i)
            ws.write(self.Row.Data.value + i, self.Col.X.value, d.x)
            ws.write(self.Row.Data.value + i, self.Col.Y.value, d.y)
            ws.write(self.Row.Data.value + i, self.Col.Z.value, d.z)
            ws.write(self.Row.Data.value + i, self.Col.Magnitude.value, d.magnitude)
        i = len(data)
        ws.write(self.Row.Data.value + i, self.Col.Index.value, shotIndex)
        ws.write(self.Row.Data.value + i, self.Col.Shot.value, -MAX_VALUE)
        i += 1
        ws.write(self.Row.Data.value + i, self.Col.Index.value, shotIndex)
        ws.write(self.Row.Data.value + i, self.Col.Shot.value, MAX_VALUE)
        self.ws.append(ws)
        
    def finalize(self):
        self.wb.close()
    
    
class xlsxAllData:
    __EXTENSION: str = 'xlsx'
    __OPEN_MODE:  str = 'w'
    __TYPES: str = [ 'GYRO', 'ACCEL', 'HI-G' ]
    __HEADERS: str = [ '', 'Index', 'X', 'Y', 'Z', 'Shot' ]
    
    class Row(enum.Enum):
        Header = 0
        Data = 1
        
    class Col(enum.Enum):
        Type = 0
        Index = 1
        X = 2
        Y = 3
        Z = 4
        Shot = 5
        
    class Field(enum.Enum):
        Gyro = 0
        Accel = 1
        HiG = 2
        
    def __init__(self, name: str, folderPath : str):
        self.name: str = str(name)
        self.folderPath : str = folderPath
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(os.path.join(self.folderPath, '{0}.{1}'.format(self.name, self.__EXTENSION)))
        self.ws: typing.List[xlsxwriter.Workbook.worksheet_class] = []
        
    def __addHeader(self, ws : xlsxwriter.Workbook.worksheet_class):
        i: int = 0
        for j, field in enumerate(self.__TYPES):
            ws.write(self.Row.Header.value, i, self.__TYPES[j])
            for k, col in enumerate(self.__HEADERS):
                if col:
                    ws.write(self.Row.Header.value, i, self.__HEADERS[k])
                i += 1
                
    def __addChart(self, ws: xlsxwriter.Workbook.worksheet_class, row: int, col: int, fileName: str, type: str, positionRow: int, positionCol: int, plotX: bool = True, plotY: bool = True, plotZ: bool = True):
        chart = self.wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
        if plotX:
            chart.add_series({
                'name':         [fileName, self.Row.Header.value, self.Col.X.value + col],
                'categories':   [fileName, self.Row.Data.value, self.Col.Index.value + col, self.Row.Data.value + row, self.Col.Index.value + col],
                'values':       [fileName, self.Row.Data.value, self.Col.X.value + col, self.Row.Data.value + row, self.Col.X.value + col],
                'line':         {'color': '#B01010'},
            })
        if plotY:
            chart.add_series({
                'name':         [fileName, self.Row.Header.value, self.Col.Y.value + col],
                'categories':   [fileName, self.Row.Data.value, self.Col.Index.value + col, self.Row.Data.value + row, self.Col.Index.value + col],
                'values':       [fileName, self.Row.Data.value, self.Col.Y.value + col, self.Row.Data.value + row, self.Col.Y.value + col],
                'line':         {'color': '#10B010'},
            })
        if plotZ:
            chart.add_series({
                'name':         [fileName, self.Row.Header.value, self.Col.Z.value + col],
                'categories':   [fileName, self.Row.Data.value, self.Col.Index.value + col, self.Row.Data.value + row, self.Col.Index.value + col],
                'values':       [fileName, self.Row.Data.value, self.Col.Z.value + col, self.Row.Data.value + row, self.Col.Z.value + col],
                'line':         {'color': '#1010B0'},
            })
        chart.add_series({
            'name':         [fileName, self.Row.Header.value, self.Col.Shot.value + col],
            'categories':   [fileName, self.Row.Data.value, self.Col.Index.value + col, self.Row.Data.value + row, self.Col.Index.value + col],
            'values':       [fileName, self.Row.Data.value, self.Col.Shot.value + col, self.Row.Data.value + row, self.Col.Shot.value + col],
            'line':         {'color': '#7F7F7F', 'dash_type': 'dash'},
        })
        chart.set_title({'name': type})
        chart.set_x_axis({'name': self.__HEADERS[self.Col.Index.value]})
        ws.insert_chart(positionRow, positionCol, chart, {'x_scale': 3, 'y_scale': 3})
        
    def addData(self, s : shot.data):
        FACTOR = 1.4
        OFFSET = 20
        ws = self.wb.add_worksheet(s.fileName)
        self.__addHeader(ws)
        Data: typing.List[typing.List[shot.vector]] = [s.gyro, s.accel, s.hiG]
        ShotIndices: typing.List[int] = [s.shot.datum.index, s.shot.datum.index, s.hiGShot.datum.index]
        col: int = 0
        for j, data in enumerate(Data):
            row: int = 0
            shotIndex = ShotIndices[j]
            minIndex: int = shotIndex - OFFSET
            maxIndex: int = shotIndex + OFFSET + 1
            if minIndex < 0:
                minIndex = 0
            if maxIndex > len(data):
                maxIndex = len(data)
            maxVal: float = 0.1
            minVal: float = -0.1
            for k, d in enumerate(data[minIndex:maxIndex]):
                ws.write(self.Row.Data.value + row, col + self.Col.Index.value, k + minIndex)
                x = d.xUnit
                y = d.yUnit
                z = d.zUnit
                ws.write(self.Row.Data.value + row, col + self.Col.X.value, x)
                ws.write(self.Row.Data.value + row, col + self.Col.Y.value, y)
                ws.write(self.Row.Data.value + row, col + self.Col.Z.value, z)
                maxVal = max(x, y, z, maxVal)
                minVal = min(x, y, z, minVal)
                row += 1
            ws.write(self.Row.Data.value + row, col + self.Col.Index.value, shotIndex)
            ws.write(self.Row.Data.value + row, col + self.Col.Shot.value, minVal * FACTOR)
            row += 1
            ws.write(self.Row.Data.value + row, col + self.Col.Index.value, shotIndex)
            ws.write(self.Row.Data.value + row, col + self.Col.Shot.value, maxVal * FACTOR)
            self.__addChart(ws, row, col, s.fileName, self.__TYPES[j], self.Row.Data.value, self.Col.Index.value + col)
            col += len(self.Col)
        self.__addChart(ws, len(s.gyro), 0, s.fileName, 'Gyro-Y', self.Row.Data.value + 43, self.Col.Index.value + col - len(self.Col), False, True, False)
            
        self.ws.append(ws)
        
    def finalize(self):
        self.wb.close()