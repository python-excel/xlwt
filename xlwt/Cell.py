# -*- coding: windows-1252 -*-

from struct import unpack, pack
import BIFFRecords

class StrCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__sst_idx"]

    def __init__(self, parent, idx, xf_idx, sst_idx):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__sst_idx = sst_idx

    def get_biff_data(self):
        return BIFFRecords.LabelSSTRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, self.__sst_idx).get()

class BlankCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx"]

    def __init__(self, parent, idx, xf_idx):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx

    def get_biff_data(self):
        return BIFFRecords.BlankRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx).get()

class MulBlankCell(object):
    __slots__ = ["__parent", "__col1", "__col2", "__xf_idx"]

    def __init__(self, parent, col1, col2, xf_idx):
        self.__parent = parent
        self.__col1 = col1
        self.__col2 = col2
        self.__xf_idx = xf_idx

    def get_biff_data(self):
        return BIFFRecords.MulBlankRecord(self.__parent.get_index(),
            self.__col1, self.__col2, self.__xf_idx).get()

class NumberCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__number"]

    def __init__(self, parent, idx, xf_idx, number):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__number = float(number)

    def get_biff_data(self):
        rk_encoded = 0
        num = self.__number

        # The four possible kinds of RK encoding are *not* mutually exclusive.
        # The 30-bit integer variety picks up the most.
        # In the code below, the four varieties are checked in descending order
        # of bangs per buck, or not at all.
        # SJM 2007-10-01

        if -0x20000000 <= num < 0x20000000: # fits in 30-bit *signed* int
            inum = int(num)
            if inum == num: # survives round-trip
                # print "30-bit integer RK", inum, hex(inum)
                rk_encoded = 2 | (inum << 2)
                return BIFFRecords.RKRecord(self.__parent.get_index(),
                    self.__idx, self.__xf_idx, rk_encoded).get()        

        temp = num * 100
        
        if -0x20000000 <= temp < 0x20000000:
            # That was step 1: the coded value will fit in
            # a 30-bit signed integer.
            itemp = int(round(temp, 0))
            # That was step 2: "itemp" is the best candidate coded value.
            # Now for step 3: simulate the decoding,
            # to check for round-trip correctness.
            if itemp / 100.0 == num:
                # print "30-bit integer RK*100", itemp, hex(itemp)
                rk_encoded = 3 | (itemp << 2)
                return BIFFRecords.RKRecord(self.__parent.get_index(),
                    self.__idx, self.__xf_idx, rk_encoded).get()

        if 0: # Cost of extra pack+unpack not justified by tiny yield.
            packed = pack('<d', num)
            w01, w23 = unpack('<2i', packed)
            if not w01 and not(w23 & 3):
                # 34 lsb are 0
                # print "float RK", w23, hex(w23)
                return BIFFRecords.RKRecord(self.__parent.get_index(),
                    self.__idx, self.__xf_idx, w23).get()

            packed100 = pack('<d', temp)
            w01, w23 = unpack('<2i', packed100)
            if not w01 and not(w23 & 3):
                # 34 lsb are 0
                # print "float RK*100", w23, hex(w23)
                return BIFFRecords.RKRecord(self.__parent.get_index(),
                    self.__idx, self.__xf_idx, w23 | 1).get()

        #print "Number" 
        #print
        return BIFFRecords.NumberRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, num).get()

class BooleanCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__number"]

    def __init__(self, parent, idx, xf_idx, number):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__number = number

    def get_biff_data(self):
        return BIFFRecords.BoolErrRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, self.__number, 0).get()

error_code_map = {
    0x00:  0, # Intersection of two cell ranges is empty
    0x07:  7, # Division by zero
    0x0F: 15, # Wrong type of operand
    0x17: 23, # Illegal or deleted cell reference
    0x1D: 29, # Wrong function or range name
    0x24: 36, # Value range overflow
    0x2A: 42, # Argument or function not available
    '#NULL!' :  0, # Intersection of two cell ranges is empty
    '#DIV/0!':  7, # Division by zero
    '#VALUE!': 36, # Wrong type of operand
    '#REF!'  : 23, # Illegal or deleted cell reference
    '#NAME?' : 29, # Wrong function or range name
    '#NUM!'  : 36, # Value range overflow
    '#N/A!'  : 42, # Argument or function not available
}

class ErrorCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__number"]

    def __init__(self, parent, idx, xf_idx, error_string_or_code):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        try:
            self.__number = error_code_map[error_string_or_code]
        except KeyError:
            raise Exception('Illegal error value (%r)' % error_string_or_code)

    def get_biff_data(self):
        return BIFFRecords.BoolErrRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, self.__number, 1).get()

class FormulaCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__frmla", "__calc_flags"]

    def __init__(self, parent, idx, xf_idx, frmla, calc_flags=0):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__frmla = frmla
        self.__calc_flags = calc_flags

    def get_biff_data(self):
        return BIFFRecords.FormulaRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, self.__frmla.rpn(), self.__calc_flags).get()
