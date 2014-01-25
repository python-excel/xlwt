# -*- coding: windows-1252 -*-

import ExcelFormulaParser, ExcelFormulaLexer
import struct
from antlr import ANTLRException


class Formula(object):
    __slots__ = ["__init__",  "__s", "__parser", "__sheet_refs", "__xcall_refs"]


    def __init__(self, s):
        try:
            self.__s = s
            lexer = ExcelFormulaLexer.Lexer(s)
            self.__parser = ExcelFormulaParser.Parser(lexer)
            self.__parser.formula()
            self.__sheet_refs = self.__parser.sheet_references
            self.__xcall_refs = self.__parser.xcall_references
        except ANTLRException, e:
            # print e
            raise ExcelFormulaParser.FormulaParseException, "can't parse formula " + s

    def get_references(self):
        return self.__sheet_refs, self.__xcall_refs

    def patch_references(self, patches):
        for offset, idx in patches:
            self.__parser.rpn = self.__parser.rpn[:offset] + struct.pack('<H', idx) + self.__parser.rpn[offset+2:]

    def text(self):
        return self.__s

    def rpn(self):
        '''
        Offset    Size    Contents
        0         2       Size of the following formula data (sz)
        2         sz      Formula data (RPN token array)
        [2+sz]    var.    (optional) Additional data for specific tokens

        '''
        return struct.pack("<H", len(self.__parser.rpn)) + self.__parser.rpn

class Formula3D(object):

    def __init__(self,range1,range2=None):

        assert type(range1) in (tuple,list)
        assert range2==None or type(range2) in (tuple,list)

        self.__sheet_refs=[]
        self.__xcall_refs=[]

        self.__sheet_refs.append([range1[0],range1[1],3])
        self.__rpn=struct.pack('<BHHHHH',0x3B,0x01,*range1[2:])

        if (range2!=None):

            self.__sheet_refs.append([range2[0],range2[1],len(self.__rpn)+3])
            self.__rpn+=struct.pack('<BHHHHHB',
                    0x3B,
                    0x00,
                    range2[2],range2[3],range2[4],range2[5],
                    0x10 # tList
                )

        self.__rpn=struct.pack('<H',len(self.__rpn))+self.__rpn

    def get_references(self):
        return self.__sheet_refs, self.__xcall_refs

    def patch_references(self, patches):
        for offset, idx in patches:
            self.__rpn = self.__rpn[:offset] + struct.pack('<H', idx) + self.__rpn[offset+2:]

    def patch_references(self, patches):
        for offset, idx in patches:
            self.__rpn = self.__rpn[:offset] + struct.pack('<H', idx) + self.__rpn[offset+2:]

    def text(self):
        return self.__s

    def rpn(self):
        return self.__rpn


