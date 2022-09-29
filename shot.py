# === IMPORTS ==================================================================

import enum
import math
import operator
import os
import string
import typing


# === ENUM =====================================================================

class Handedness(enum.Enum):
    Right = 0
    Left = 1
NUM_HANDEDNESS = len(Handedness)

class ShotConfidence(enum.Enum):
    NoShot = 0
    VeryLow = 1
    Low = 2
    Medium = 3
    High = 4
    VeryHigh = 5
NUM_SHOT_CONFIDENCE = len(ShotConfidence)

SHOT_SEPARATION: int = 30


# === GLOBAL CONSTANTS =========================================================

LSB_TO_G_DIVISOR = 4096

TYPE_IMU_GRYO = 1
TYPE_IMU_ACCEL = 2
TYPE_HI_G_ACCEL = 3
TYPE_CALIBRATION = 4
TYPE_SETTINGS = 5
TYPE_HI_G_ACCEL_COMP = 7


# === HELPER FUNCTIONS =========================================================

def findThreeAxisMagnitude(x : float, y : float, z : float) -> float:
    return math.sqrt((x * x) + (y * y) + (z * z))

def convertLsbToG(a : int) -> float:
    return float(a / LSB_TO_G_DIVISOR)

def convertLsbToG(a : float) -> float:
    return float(a / LSB_TO_G_DIVISOR)


# === CLASSES ==================================================================

class vector:
    def __init__(self, x : float, y : float, z: float):
        self.x: float = x
        self.y: float = y
        self.z: float = z
        self.magnitude: float = findThreeAxisMagnitude(self.x, self.y, self.z)
        self.xG: float = convertLsbToG(self.x)
        self.yG: float = convertLsbToG(self.y)
        self.zG: float = convertLsbToG(self.z)
        self.magnitudeG: float = convertLsbToG(self.magnitude)
        self.absX: float = abs(x)
        self.absY: float = abs(y)
        self.absZ: float = abs(z)
        
    def getVectorSum(self) -> int:
        return abs(self.x) + abs(self.y) + abs(self.z)

    def gyroEntryString(self) -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_IMU_GRYO), int(self.x), int(self.y), int(self.z))
        
    def accelEntryString(self) -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_IMU_ACCEL), int(self.x), int(self.y), int(self.z))
    
    def hiGEntryString(self) -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_HI_G_ACCEL), int(self.x), int(self.y), int(self.z))
        
    def calibEntryString(self) -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_CALIBRATION), float(self.x), float(self.y), float(self.z))
        
    def rightHandEntryString() -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_SETTINGS), int(Handedness.Right.value), int(0), int(0))
    
    def leftHandEntryString() -> str:
        return '{0}, {1}, {2}, {3}'.format(int(TYPE_SETTINGS), int(Handedness.Left.value), int(0), int(0))
    
    
class vectorDatum:
    def __init__(self, v: vector, index: int = 0):
        self.v: vector = v
        self.index: int = index
        
        
class shotDatum:
    def __init__(self, datum: vectorDatum, confidence: ShotConfidence = ShotConfidence.NoShot):
        self.datum: vectorDatum = datum
        self.confidence: ShotConfidence = confidence
        
        
class data:
    class LineIndex(enum.Enum):
        Type = 0
        X = 1
        Y = 2
        Z = 3
    NUM_LINE_INDICES = len(LineIndex)

    def __init__(self, fileName: str):
        self.fileName: str = fileName
        self.filePath: str = ''
        self.name: str = ''
        self.gyro: typing.List[vector] = []
        self.accel: typing.List[vector] = []
        self.hiG: typing.List[vector] = []
        self.calibration: vector = vector(0, 0, 0)
        self.handedness: Handedness = Handedness.Right
        self.maxGyro: vectorDatum
        self.maxAccel: vectorDatum
        self.maxHiG: vectorDatum
        self.maxAccelX: vectorDatum
        self.maxAccelY: vectorDatum
        self.maxAccelZ: vectorDatum
        self.shot: shotDatum
        self.altShot: shotDatum
        if self.fileName:
            self.__process()
            self.__analyze()
            self.__processShot()
        
    def __process(self):
        if self.fileName:
            self.name = self.fileName.replace('.csv', '')
            self.filePath = os.path.join(os.getcwd(), self.fileName)
            with open(self.filePath, 'r') as file:
                readLines = file.readlines()
            file.close()
            self.numIMU = 0
            processedCalibration : bool = False
            for line in readLines:
                line = line.strip()
                entries = line.split(',')
                type = int(entries[self.LineIndex.Type.value])
                v = vector(float(entries[self.LineIndex.X.value]),
                           float(entries[self.LineIndex.Y.value]),
                           float(entries[self.LineIndex.Z.value]))
                
                if (type == TYPE_IMU_GRYO):
                    self.gyro.append(v)
                elif (type == TYPE_IMU_ACCEL):
                    self.accel.append(v)
                elif (type == TYPE_HI_G_ACCEL):
                    self.hiG.append(v)
                elif (type == TYPE_CALIBRATION):
                    self.calibration = v
                    processedCalibration = True
                elif (type == TYPE_SETTINGS):
                    n = int(entries[self.LineIndex.X.value])
                    self.handedness = Handedness.Right
                    if (n == Handedness.Left.value):
                        self.handedness = Handedness.Left
                elif (type == TYPE_HI_G_ACCEL_COMP):
                    x : int = int(entries[self.LineIndex.X.value])
                    y : int = int(entries[self.LineIndex.Y.value])
                    z : int = int(entries[self.LineIndex.Z.value])
                    HI_SHIFT = 8
                    LO_SHIFT = 0
                    MASK = 0xff
                    v0 : vector = vector(
                        (x >> LO_SHIFT) & MASK,
                        (x >> HI_SHIFT) & MASK,
                        (y >> LO_SHIFT) & MASK)
                    v1 : vector = vector(
                        (y >> HI_SHIFT) & MASK,
                        (z >> LO_SHIFT) & MASK,
                        (z >> HI_SHIFT) & MASK)
                    self.hiG.append(v0)
                    self.hiG.append(v1)
            if not processedCalibration:
                if self.handedness is Handedness.Left:
                    self.calibration = vector(0.280273, -0.979248, 0.011719)
                else:
                    self.calibration = vector(-0.280273, -0.979248, -0.011719)
        
    def __analyze(self):
        MAGNITUDE = 'magnitude'
        v: vector = max(self.gyro, key = operator.attrgetter(MAGNITUDE))
        i: int = self.gyro.index(v)
        self.maxGyro = vectorDatum(v, i)
        
        v: vector = max(self.accel, key = operator.attrgetter(MAGNITUDE))
        i: int = self.accel.index(v)
        self.maxAccel = vectorDatum(v, i)
        
        v: vector = max(self.hiG, key = operator.attrgetter(MAGNITUDE))
        i: int = self.hiG.index(v)
        self.maxHiG = vectorDatum(v, i)
        
        v: vector = max(self.accel, key = operator.attrgetter('absX'))
        i: int = self.accel.index(v)
        self.maxAccelX = vectorDatum(v, i)
        
        v: vector = max(self.accel, key = operator.attrgetter('absY'))
        i: int = self.accel.index(v)
        self.maxAccelY = vectorDatum(v, i)
        
        v: vector = max(self.accel, key = operator.attrgetter('absZ'))
        i: int = self.accel.index(v)
        self.maxAccelZ = vectorDatum(v, i)
        
    def __processShot(self):
        self.shot = self.__findShotNew()
        self.altShot = self.__findShotNew(self.shot.datum.index + SHOT_SEPARATION)
        if self.shot.confidence != ShotConfidence.NoShot and self.altShot.confidence != ShotConfidence.NoShot:
            if self.altShot.datum.v.magnitude > self.shot.datum.v.magnitude:
                self.shot.confidence = ShotConfidence.VeryLow
        
    def __findShot(self, offset: int = 0) -> shotDatum:
        __SHOT_THRESHOLD: int = 10000
        __SCORE_LUT: typing.List[ShotConfidence] = (
            ShotConfidence.NoShot,
            ShotConfidence.Medium,
            ShotConfidence.High,
            ShotConfidence.VeryHigh,
        )
        shotConfidence = ShotConfidence.NoShot
        shotIndex: int = 0
        
        for i, v in enumerate(self.accel[offset:]):
            index: int = offset + i
            score: int = 0
            if v.absX > __SHOT_THRESHOLD:
                score += 1
            if v.absY > __SHOT_THRESHOLD:
                score += 1
            if v.absZ > __SHOT_THRESHOLD:
                score += 1
            if score >= len(__SCORE_LUT):
                score = len(__SCORE_LUT) - 1
            if score > 0:
                shotIndex = index
                shotConfidence = __SCORE_LUT[score]
                break
            
        shot : shotDatum = shotDatum(vectorDatum(self.accel[shotIndex], shotIndex), shotConfidence)
        return shot
            
    def __findShotNew(self, offset: int = 0) -> shotDatum:
        DETECT_THRESHOLD : int = 22000
        DETECT_PAIR_ADD_THRESHOLD : int = 29000
        DETECT_PAIR_STRONG_THRESHOLD : int = 17000
        shotConfidence : ShotConfidence = ShotConfidence.NoShot
        shotIndex : int = 0
        prevMagnitude : int = 0
        if offset > 0 and offset < len(self.accel):
            prevMagnitude = self.accel[offset - 1].magnitude
        for i, v in enumerate(self.accel[offset:]):
            index: int = offset + i
            if v.magnitude >= 40000:
                shotConfidence = ShotConfidence.VeryHigh
                shotIndex = index
                break
            elif v.magnitude >= 35000:
                shotConfidence = ShotConfidence.High
                shotIndex = index
                break
            elif v.magnitude >= 30000:
                shotConfidence = ShotConfidence.Medium
                shotIndex = index
                break
            elif v.magnitude >= DETECT_THRESHOLD:
                shotConfidence = ShotConfidence.Low
                shotIndex = index
                break
            elif (prevMagnitude >= DETECT_PAIR_STRONG_THRESHOLD or v.magnitude >= DETECT_PAIR_STRONG_THRESHOLD) and (v.magnitude + prevMagnitude) >= DETECT_PAIR_ADD_THRESHOLD:
                shotConfidence = ShotConfidence.VeryLow
                shotIndex = index
                break
            prevMagnitude = v.magnitude
        shot : shotDatum = shotDatum(vectorDatum(self.accel[shotIndex], shotIndex), shotConfidence)
        return shot
    
    def __writeEntryToFile():
        print('do something')
