### $ANTLR 2.7.5 (20050128): "excel-formula.g" -> "ExcelFormulaParser.py"$
### import antlr and other modules ..
import sys
import antlr

version = sys.version.split()[0]
if version < '2.2.1':
    False = 0
if version < '2.3':
    True = not False
### header action >>> 
__rev_id__ = """$Id$"""

import struct
import Utils
from UnicodeUtils import upack1
from ExcelMagic import *
### header action <<< 
### preamble action>>>

### preamble action <<<

### import antlr.Token 
from antlr import Token
### >>>The Known Token Types <<<
SKIP                = antlr.SKIP
INVALID_TYPE        = antlr.INVALID_TYPE
EOF_TYPE            = antlr.EOF_TYPE
EOF                 = antlr.EOF
NULL_TREE_LOOKAHEAD = antlr.NULL_TREE_LOOKAHEAD
MIN_USER_TYPE       = antlr.MIN_USER_TYPE
STR_CONST = 4
NUM_CONST = 5
NAME = 6
ASSIGN = 7
ADD = 8
SUB = 9
MUL = 10
DIV = 11
LP = 12
RP = 13
LB = 14
RB = 15
COLON = 16
COMMA = 17
SEMICOLON = 18
CONCAT = 19
INT_CONST = 20
REF2D = 21

class Parser(antlr.LLkParser):
    ### user action >>>
    ### user action <<<
    
    def __init__(self, *args, **kwargs):
        antlr.LLkParser.__init__(self, *args, **kwargs)
        self.tokenNames = _tokenNames
        ### __init__ header action >>> 
        self.rpn = ""
        ### __init__ header action <<< 
        
    def formula(self):    
        
        pass
        self.expr()
    
    def expr(self):    
        
        pass
        self.term()
        while True:
            if (self.LA(1)==ADD or self.LA(1)==SUB or self.LA(1)==CONCAT):
                pass
                op=self.term_op()
                self.term()
                self.rpn += op
            else:
                break
            
    
    def term(self):    
        
        pass
        self.fact()
        while True:
            if (self.LA(1)==MUL or self.LA(1)==DIV):
                pass
                op=self.fact_op()
                self.fact()
                self.rpn += op
            else:
                break
            
    
    def term_op(self):    
        result = None
        
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [ADD]:
            pass
            self.match(ADD)
            result = struct.pack('B', ptgAdd)
        elif la1 and la1 in [SUB]:
            pass
            self.match(SUB)
            result = struct.pack('B', ptgSub)
        elif la1 and la1 in [CONCAT]:
            pass
            self.match(CONCAT)
            result = struct.pack('B', ptgConcat)
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
        return result
    
    def fact(self):    
        
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST,NUM_CONST,NAME,LP,INT_CONST,REF2D]:
            pass
            self.primary()
        elif la1 and la1 in [SUB]:
            pass
            self.match(SUB)
            self.primary()
            self.rpn += struct.pack('B', ptgUminus)
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
    
    def fact_op(self):    
        result = None
        
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [MUL]:
            pass
            self.match(MUL)
            result = struct.pack('B', ptgMul)
        elif la1 and la1 in [DIV]:
            pass
            self.match(DIV)
            result = struct.pack('B', ptgDiv)
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
        return result
    
    def primary(self):    
        
        str_tok = None
        int_tok = None
        num_tok = None
        ref2d_tok = None
        ref2d1_tok = None
        ref2d2_tok = None
        name_tok = None
        func_tok = None
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST]:
            pass
            str_tok = self.LT(1)
            self.match(STR_CONST)
            self.rpn += struct.pack("B", ptgStr) + upack1(str_tok.text[1:-1])
        elif la1 and la1 in [INT_CONST]:
            pass
            int_tok = self.LT(1)
            self.match(INT_CONST)
            self.rpn += struct.pack("<BH", ptgInt, int(int_tok.text))
        elif la1 and la1 in [NUM_CONST]:
            pass
            num_tok = self.LT(1)
            self.match(NUM_CONST)
            self.rpn += struct.pack("<Bd", ptgNum, float(num_tok.text))
        elif la1 and la1 in [LP]:
            pass
            self.match(LP)
            self.expr()
            self.match(RP)
            self.rpn += struct.pack("B", ptgParen)
        else:
            if (self.LA(1)==REF2D) and (_tokenSet_0.member(self.LA(2))):
                pass
                ref2d_tok = self.LT(1)
                self.match(REF2D)
                r, c = Utils.cell_to_packed_rowcol(ref2d_tok.text)
                self.rpn += struct.pack("<B2H", ptgRefV, r, c)
            elif (self.LA(1)==REF2D) and (self.LA(2)==COLON):
                pass
                ref2d1_tok = self.LT(1)
                self.match(REF2D)
                self.match(COLON)
                ref2d2_tok = self.LT(1)
                self.match(REF2D)
                r1, c1 = Utils.cell_to_packed_rowcol(ref2d1_tok.text)
                r2, c2 = Utils.cell_to_packed_rowcol(ref2d2_tok.text)
                self.rpn += struct.pack("<B4H", ptgAreaV, r1, r2, c1, c2)
            elif (self.LA(1)==NAME) and (_tokenSet_0.member(self.LA(2))):
                pass
                name_tok = self.LT(1)
                self.match(NAME)
                self.rpn += ""
            elif (self.LA(1)==NAME) and (self.LA(2)==LP):
                pass
                func_tok = self.LT(1)
                self.match(NAME)
                self.match(LP)
                arg_count=self.expr_list()
                self.match(RP)
                if func_tok.text.upper() in std_func_by_name:
                   (opcode,
                   min_argc,
                   max_argc, 
                   func_type,
                   func_arg, 
                   volatile_func) = std_func_by_name[func_tok.text.upper()]
                   if arg_count > max_argc or arg_count < min_argc:
                       raise Exception, "%d parameters for function: %s" % (arg_count, func_tok.text)
                   if min_argc == max_argc:
                       if func_type == "V":
                           func_ptg = ptgFuncV
                       elif func_type == "R":
                           func_ptg = ptgFuncR
                       elif func_type == "A":
                           func_ptg = ptgFuncA
                       else:
                           raise Exception, "wrong function type"
                       self.rpn += struct.pack("<BH", func_ptg, opcode) 
                   else:
                       if func_type == "V":
                           func_ptg = ptgFuncVarV
                       elif func_type == "R":
                           func_ptg = ptgFuncVarR
                       elif func_type == "A":
                           func_ptg = ptgFuncVarA
                       else:
                           raise Exception, "wrong function type"
                       self.rpn += struct.pack("<2BH", func_ptg, arg_count, opcode) 
                else:
                   raise Exception, "unknown function: %s" % func_tok.text
            else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
    
    def expr_list(self):    
        arg_cnt = None
        
        arg_cnt = 0
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST,NUM_CONST,NAME,SUB,LP,INT_CONST,REF2D]:
            pass
            self.expr()
            arg_cnt += 1
            while True:
                if (self.LA(1)==SEMICOLON):
                    pass
                    self.match(SEMICOLON)
                    self.expr()
                    arg_cnt += 1
                else:
                    break
                
        elif la1 and la1 in [RP]:
            pass
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
        return arg_cnt
    

_tokenNames = [
    "<0>", 
    "EOF", 
    "<2>", 
    "NULL_TREE_LOOKAHEAD", 
    "STR_CONST", 
    "NUM_CONST", 
    "NAME", 
    "ASSIGN", 
    "ADD", 
    "SUB", 
    "MUL", 
    "DIV", 
    "LP", 
    "RP", 
    "LB", 
    "RB", 
    "COLON", 
    "COMMA", 
    "SEMICOLON", 
    "CONCAT", 
    "INT_CONST", 
    "REF2D"
]
    

### generate bit set
def mk_tokenSet_0(): 
    ### var1
    data = [ 798466L, 0L]
    return data
_tokenSet_0 = antlr.BitSet(mk_tokenSet_0())
    
