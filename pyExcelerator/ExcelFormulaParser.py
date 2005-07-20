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
from ExcelFormula import *
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
INT_CONST = 19
RC_CELL = 20
EQUAL = 21
NOT_EQUAL = 22
LT = 23
LE = 24
GE = 25
GT = 26
WHITE_SPACE = 27
DIGIT = 28
INT = 29
LOWER = 30
UPPER = 31
LETTER = 32

class Parser(antlr.LLkParser):
    ### user action >>>
    ### user action <<<
    
    def __init__(self, *args, **kwargs):
        antlr.LLkParser.__init__(self, *args, **kwargs)
        self.tokenNames = _tokenNames
        ### __init__ header action >>> 
        self.postfix_stack = []
        ### __init__ header action <<< 
        
    def formula(self):    
        
        pass
        self.expr()
    
    def expr(self):    
        
        pass
        self.term()
        while True:
            if (self.LA(1)==ADD or self.LA(1)==SUB):
                pass
                op=self.term_op()
                self.term()
                self.postfix_stack.extend([op])
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
                self.postfix_stack.extend([op])
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
            result = (ptgAdd, "ptgAdd")
        elif la1 and la1 in [SUB]:
            pass
            self.match(SUB)
            result = (ptgSub, "ptgSub")
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
        return result
    
    def fact(self):    
        
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST,NUM_CONST,NAME,LP,INT_CONST,RC_CELL]:
            pass
            self.primary()
        elif la1 and la1 in [SUB]:
            pass
            self.match(SUB)
            self.primary()
            self.postfix_stack.extend([(ptgUminus, "ptgUminus")])
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
            result = (ptgMul, "ptgMul")
        elif la1 and la1 in [DIV]:
            pass
            self.match(DIV)
            result = (ptgDiv, "ptgDiv")
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
        return result
    
    def primary(self):    
        
        str_tok = None
        int_tok = None
        num_tok = None
        rc_cell = None
        rc0 = None
        rc1 = None
        name_tok = None
        func_tok = None
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST]:
            pass
            str_tok = self.LT(1)
            self.match(STR_CONST)
            self.postfix_stack.extend([(ptgStr, str_tok.text)])
        elif la1 and la1 in [INT_CONST]:
            pass
            int_tok = self.LT(1)
            self.match(INT_CONST)
            self.postfix_stack.extend([(ptgInt, int(int_tok.text))])
        elif la1 and la1 in [NUM_CONST]:
            pass
            num_tok = self.LT(1)
            self.match(NUM_CONST)
            self.postfix_stack.extend([(ptgNum, float(num_tok.text))])
        elif la1 and la1 in [LP]:
            pass
            self.match(LP)
            self.expr()
            self.match(RP)
        else:
            if (self.LA(1)==RC_CELL) and (_tokenSet_0.member(self.LA(2))):
                pass
                rc_cell = self.LT(1)
                self.match(RC_CELL)
                self.postfix_stack.extend([(ptgRefV, rc_cell.text)])
            elif (self.LA(1)==RC_CELL) and (self.LA(2)==COLON):
                pass
                rc0 = self.LT(1)
                self.match(RC_CELL)
                self.match(COLON)
                rc1 = self.LT(1)
                self.match(RC_CELL)
                self.postfix_stack.extend([(ptgRange, "%s : %s" % (rc0.text, rc1.text))])
            elif (self.LA(1)==NAME) and (_tokenSet_0.member(self.LA(2))):
                pass
                name_tok = self.LT(1)
                self.match(NAME)
                self.postfix_stack.extend([(ptgName, name_tok.text)])
            elif (self.LA(1)==NAME) and (self.LA(2)==LP):
                pass
                func_tok = self.LT(1)
                self.match(NAME)
                self.match(LP)
                self.expr_list()
                self.match(RP)
                self.postfix_stack.extend([(ptgFunc, func_tok.text)])
            else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
    
    def expr_list(self):    
        
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in [STR_CONST,NUM_CONST,NAME,SUB,LP,INT_CONST,RC_CELL]:
            pass
            self.expr()
            while True:
                if (self.LA(1)==SEMICOLON):
                    pass
                    self.match(SEMICOLON)
                    self.expr()
                else:
                    break
                
        elif la1 and la1 in [RP]:
            pass
        else:
                raise antlr.NoViableAltException(self.LT(1), self.getFilename())
            
    

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
    "INT_CONST", 
    "RC_CELL", 
    "EQUAL", 
    "NOT_EQUAL", 
    "LT", 
    "LE", 
    "GE", 
    "GT", 
    "WHITE_SPACE", 
    "DIGIT", 
    "INT", 
    "LOWER", 
    "UPPER", 
    "LETTER"
]
    

### generate bit set
def mk_tokenSet_0(): 
    ### var1
    data = [ 274178L, 0L]
    return data
_tokenSet_0 = antlr.BitSet(mk_tokenSet_0())
    
