# -*- coding: windows-1252 -*-

import sys
from antlr import EOF, CommonToken as Tok, TokenStream, TokenStreamException
import struct
import ExcelFormulaParser
from re import compile as recompile, match, LOCALE, UNICODE, IGNORECASE, VERBOSE


int_const_pattern = recompile(r"\d+")
flt_const_pattern = recompile(r"""
    (?:
        (?: \d* \. \d+ ) # .1 .12 .123 etc 9.1 etc 98.1 etc
        |
        (?: \d+ \. ) # 1. 12. 123. etc
    )
    # followed by optional exponent part
    (?: [Ee] [+-]? \d+ ) ?
    """, VERBOSE)
str_const_pattern = recompile(r'["][^"]*["]')
#range2d_pattern   = recompile(r"\$?[A-I]?[A-Z]\$?\d+:\$?[A-I]?[A-Z]\$?\d+")
ref2d_r1c1_pattern = recompile(r"[Rr]0*[1-9][0-9]*[Cc]0*[1-9][0-9]*")
ref2d_pattern     = recompile(r"\$?[A-I]?[A-Z]\$?0*[1-9][0-9]*", IGNORECASE)
true_pattern      = recompile(r"TRUE", IGNORECASE)
false_pattern     = recompile(r"FALSE", IGNORECASE)
name_pattern      = recompile(r"\w[\.\w]*", LOCALE+IGNORECASE)
ne_pattern        = recompile(r"<>")
ge_pattern        = recompile(r">=")
le_pattern        = recompile(r"<=")


pattern_type_tuples = (
    (flt_const_pattern, ExcelFormulaParser.NUM_CONST),
    (int_const_pattern, ExcelFormulaParser.INT_CONST),
    (str_const_pattern, ExcelFormulaParser.STR_CONST),
#    (range2d_pattern  , ExcelFormulaParser.RANGE2D),
    (ref2d_r1c1_pattern, ExcelFormulaParser.REF2D_R1C1),
    (ref2d_pattern    , ExcelFormulaParser.REF2D),
    (true_pattern     , ExcelFormulaParser.TRUE_CONST),
    (false_pattern    , ExcelFormulaParser.FALSE_CONST),
    (name_pattern     , ExcelFormulaParser.NAME),
    (ne_pattern,        ExcelFormulaParser.NE),
    (ge_pattern,        ExcelFormulaParser.GE),
    (le_pattern,        ExcelFormulaParser.LE),
)

single_char_lookup = {
    '=': ExcelFormulaParser.EQ,
    '<': ExcelFormulaParser.LT,
    '>': ExcelFormulaParser.GT,
    '+': ExcelFormulaParser.ADD,
    '-': ExcelFormulaParser.SUB,
    '*': ExcelFormulaParser.MUL,
    '/': ExcelFormulaParser.DIV,
    ':': ExcelFormulaParser.COLON,
    ';': ExcelFormulaParser.SEMICOLON,
    ',': ExcelFormulaParser.COMMA,
    '(': ExcelFormulaParser.LP,
    ')': ExcelFormulaParser.RP,
    '&': ExcelFormulaParser.CONCAT,
    '%': ExcelFormulaParser.PERCENT,
    '^': ExcelFormulaParser.POWER,
    }
    
class Lexer(TokenStream):
    def __init__(self, text):
        self._text = text[:]
        self._pos = 0
        self._line = 0

    def isEOF(self):
        return len(self._text) <= self._pos

    def curr_ch(self):
        return self._text[self._pos]

    def next_ch(self, n = 1):
        self._pos += n

    def is_whitespace(self):
        return self.curr_ch() in " \t\n\r\f\v"

    def match_pattern(self, pattern, toktype):
        m = pattern.match(self._text[self._pos:])
        if m:
            start_pos = self._pos + m.start(0)
            end_pos = self._pos + m.end(0)
            tt = self._text[start_pos:end_pos]
            self._pos = end_pos
            return Tok(type = toktype, text = tt, col = start_pos + 1)
        else:
            return None

    def nextToken(self):
        # skip whitespace
        while not self.isEOF() and self.is_whitespace():
            self.next_ch()
        if self.isEOF():
            return Tok(type = EOF)
        # first, try to match token with 2 or more chars
        for ptt in pattern_type_tuples:
            t = self.match_pattern(*ptt);
            if t:
                return t
        # second, we want 1-char tokens
        te = self.curr_ch()
        try:
            ty = single_char_lookup[te]
        except KeyError:
            raise TokenStreamException(
                "Unexpected char %r in column %u." % (self.curr_ch(), self._pos))
        self.next_ch()
        return Tok(type=ty, text=te, col=self._pos)
