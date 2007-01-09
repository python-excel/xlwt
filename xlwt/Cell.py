#!/usr/bin/env python
# -*- coding: windows-1252 -*-

#  Copyright (C) 2005 Roman V. Kiseliov
#  All rights reserved.
# 
#  Redistribution and use in source and binary forms, with or without
#  modification, are permitted provided that the following conditions
#  are met:
# 
#  1. Redistributions of source code must retain the above copyright
#     notice, this list of conditions and the following disclaimer.
# 
#  2. Redistributions in binary form must reproduce the above copyright
#     notice, this list of conditions and the following disclaimer in
#     the documentation and/or other materials provided with the
#     distribution.
# 
#  3. All advertising materials mentioning features or use of this
#     software must display the following acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <roman@kiseliov.ru>."
# 
#  4. Redistributions of any form whatsoever must retain the following
#     acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <roman@kiseliov.ru>."
# 
#  THIS SOFTWARE IS PROVIDED BY Roman V. Kiseliov ``AS IS'' AND ANY
#  EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
#  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
#  PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL Roman V. Kiseliov OR
#  ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
#  SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
#  NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
#  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
#  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
#  STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
#  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
#  OF THE POSSIBILITY OF SUCH DAMAGE.

__rev_id__ = """$Id$"""

# SJM 2007-01-10 Fixed RK encoding bug
# SJM 2007-01-10 Having method names in __slots__ is a WOFTAM. Removed.
# SJM 2007-01-10 Unused & unuseable MulNumber class removed.

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

class FormulaCell(object):
    __slots__ = ["__parent", "__idx", "__xf_idx", "__frmla"]

    def __init__(self, parent, idx, xf_idx, frmla):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__frmla = frmla

    def get_biff_data(self):
        return BIFFRecords.FormulaRecord(self.__parent.get_index(),
            self.__idx, self.__xf_idx, self.__frmla.rpn()).get()
