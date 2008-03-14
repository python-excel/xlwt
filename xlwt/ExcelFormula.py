# -*- coding: windows-1251 -*-

import ExcelFormulaParser, ExcelFormulaLexer
import struct
from antlr import ANTLRException


class Formula(object):
    __slots__ = ["__init__", "text", "rpn", "__s", "__parser"]


    def __init__(self, s):
        try:
            self.__s = s
            lexer = ExcelFormulaLexer.Lexer(s)
            self.__parser = ExcelFormulaParser.Parser(lexer)
            self.__parser.formula()
        except ANTLRException:
            raise ExcelFormulaParser.FormulaParseException, "can't parse formula " + s

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

