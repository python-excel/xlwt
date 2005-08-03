#!/usr/bin/env python
# -*- coding: windows-1251 -*-

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


import struct
import BIFFRecords


class StrCell(object):
    __slots__ = ["__init__", "get_biff_data",
                "__parent", "__idx", "__xf_idx", "__sst_idx"]

    def __init__(self, parent, idx, xf_idx, sst_idx):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__sst_idx = sst_idx

    def get_biff_data(self):
        return BIFFRecords.LabelSSTRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, self.__sst_idx).get()


class BlankCell(object):
    __slots__ = ["__init__", "get_biff_data",
                "__parent", "__idx", "__xf_idx"]

    def __init__(self, parent, idx, xf_idx):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx

    def get_biff_data(self):
        return BIFFRecords.BlankRecord(self.__parent.get_index(), self.__idx, self.__xf_idx).get()


class MulBlankCell(object):
    __slots__ = ["__init__", "get_biff_data",
                "__parent", "__col1", "__col2", "__xf_idx"]

    def __init__(self, parent, col1, col2, xf_idx):
        self.__parent = parent
        self.__col1 = col1
        self.__col2 = col2
        self.__xf_idx = xf_idx

    def get_biff_data(self):
        return BIFFRecords.MulBlankRecord(self.__parent.get_index(), self.__col1, self.__col2, self.__xf_idx).get()


class NumberCell(object):
    __slots__ = ["__init__", "get_biff_data",
                "__parent", "__idx", "__xf_idx", "__number"]


    def __init__(self, parent, idx, xf_idx, number):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__number = float(number)


    def get_biff_data(self):
        rk_encoded = 0

        packed = struct.pack('<d', self.__number)

        #print self.__number
        w0, w1, w2, w3 = struct.unpack('<4H', packed)
        if w0 == 0 and w1 == 0 and w2 & 0xFFFC == w2:
            # 34 lsb are 0
            #print "float RK"
            rk_encoded = (w3 << 16) | w2
            return BIFFRecords.RKRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

        if abs(self.__number) < 0x40000000 and int(self.__number) == self.__number:
            #print "30-bit integer RK"
            rk_encoded = 2 | (int(self.__number) << 2)
            return BIFFRecords.RKRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

        temp = self.__number*100
        packed100 = struct.pack('<d', temp)
        w0, w1, w2, w3 = struct.unpack('<4H', packed100)
        if w0 == 0 and w1 == 0 and w2 & 0xFFFC == w2:
            # 34 lsb are 0
            #print "float RK*100"
            rk_encoded = 1 | (w3 << 16) | w2
            return BIFFRecords.RKRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

        if abs(temp) < 0x40000000 and int(temp) == temp:
            #print "30-bit integer RK*100"
            rk_encoded = 3 | (int(temp) << 2)
            return BIFFRecords.RKRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

        #print "Number" 
        #print
        return BIFFRecords.NumberRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, self.__number).get()


class MulNumberCell(object):
    __slots__ = ["__init__", "get_biff_data"]

    def __init__(self, parent, idx, xf_idx, sst_idx):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__sst_idx = sst_idx


    def get_biff_data(self):
        raise Exception


class FormulaCell(object):
    __slots__ = ["__init__", "get_biff_data",
                "__parent", "__idx", "__xf_idx", "__frmla"]

    def __init__(self, parent, idx, xf_idx, frmla):
        self.__parent = parent
        self.__idx = idx
        self.__xf_idx = xf_idx
        self.__frmla = frmla


    def get_biff_data(self):
        return BIFFRecords.FormulaRecord(self.__parent.get_index(), self.__idx, self.__xf_idx, self.__frmla.rpn()).get()



