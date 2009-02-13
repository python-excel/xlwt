# -*- coding: windows-1251 -*-

import ExcelFormulaParser, ExcelFormulaLexer
import struct
from antlr import ANTLRException


class Formula(object):
    __slots__ = ["__init__", "text", "rpn", "__s", "__parser", "__sheet_refs", "__rpn_offsets"]


    def __init__(self, s):
        try:
            self.__s = s
            lexer = ExcelFormulaLexer.Lexer(s)
            self.__parser = ExcelFormulaParser.Parser(lexer)
            self.__parser.formula()
            self.__rpn_offsets = self.__parser.rpn_offsets
            self.__sheet_refs = self.__parser.sheet_references
        except ANTLRException, e:
            # print e
            raise ExcelFormulaParser.FormulaParseException, "can't parse formula " + s

    def get_references(self):
        return self.__sheet_refs

    def update_references(self, ref_indexes):
        i = 0
        for idx in ref_indexes:
            offset = self.__rpn_offsets[i]
            self.__parser.rpn = self.__parser.rpn[:offset] + struct.pack('<H', idx) + self.__parser.rpn[offset+2:]
            i = i + 1

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

