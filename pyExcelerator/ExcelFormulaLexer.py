### $ANTLR 2.7.5 (20050128): "excel-formula.g" -> "ExcelFormulaLexer.py"$
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
### header action <<< 
### preamble action >>> 

### preamble action <<< 
### >>>The Literals<<<
literals = {}


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
DIGITS = 29
LOWER = 30
UPPER = 31
LETTER = 32

class Lexer(antlr.CharScanner) :
    ### user action >>>
    ### user action <<<
    def __init__(self, *argv, **kwargs) :
        antlr.CharScanner.__init__(self, *argv, **kwargs)
        self.caseSensitiveLiterals = True
        self.setCaseSensitive(True)
        self.literals = literals
    
    def nextToken(self):
        while True:
            try: ### try again ..
                while True:
                    _token = None
                    _ttype = INVALID_TYPE
                    self.resetText()
                    try: ## for char stream error handling
                        try: ##for lexical error handling
                            la1 = self.LA(1)
                            if False:
                                pass
                            elif la1 and la1 in u'+':
                                pass
                                self.mADD(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'-':
                                pass
                                self.mSUB(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'*':
                                pass
                                self.mMUL(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'/':
                                pass
                                self.mDIV(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'=':
                                pass
                                self.mEQUAL(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'(':
                                pass
                                self.mLP(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u')':
                                pass
                                self.mRP(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'[':
                                pass
                                self.mLB(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u']':
                                pass
                                self.mRB(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u':':
                                pass
                                self.mCOLON(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u',':
                                pass
                                self.mCOMMA(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u';':
                                pass
                                self.mSEMICOLON(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'\t\n\u000c\r ':
                                pass
                                self.mWHITE_SPACE(True)
                                theRetToken = self._returnToken
                            elif la1 and la1 in u'0123456789':
                                pass
                                self.mINT_CONST(True)
                                theRetToken = self._returnToken
                            else:
                                if (self.LA(1)==u'<') and (self.LA(2)==u'>'):
                                    pass
                                    self.mNOT_EQUAL(True)
                                    theRetToken = self._returnToken
                                elif (self.LA(1)==u'<') and (self.LA(2)==u'='):
                                    pass
                                    self.mLE(True)
                                    theRetToken = self._returnToken
                                elif (self.LA(1)==u'>') and (self.LA(2)==u'='):
                                    pass
                                    self.mGE(True)
                                    theRetToken = self._returnToken
                                elif (self.LA(1)==u'R' or self.LA(1)==u'r') and (_tokenSet_0.member(self.LA(2))):
                                    pass
                                    self.mRC_CELL(True)
                                    theRetToken = self._returnToken
                                elif (self.LA(1)==u'<') and (True):
                                    pass
                                    self.mLT(True)
                                    theRetToken = self._returnToken
                                elif (self.LA(1)==u'>') and (True):
                                    pass
                                    self.mGT(True)
                                    theRetToken = self._returnToken
                                elif (_tokenSet_1.member(self.LA(1))) and (True):
                                    pass
                                    self.mNAME(True)
                                    theRetToken = self._returnToken
                                else:
                                    self.default(self.LA(1))
                                
                            if not self._returnToken:
                                raise antlr.TryAgain ### found SKIP token
                            ### option { testLiterals=true } 
                            self.testForLiteral(self._returnToken)
                            ### return token to caller
                            return self._returnToken
                        ### handle lexical errors ....
                        except antlr.RecognitionException, e:
                            raise antlr.TokenStreamRecognitionException(e)
                    ### handle char stream errors ...
                    except antlr.CharStreamException,cse:
                        if isinstance(cse, antlr.CharStreamIOException):
                            raise antlr.TokenStreamIOException(cse.io)
                        else:
                            raise antlr.TokenStreamException(str(cse))
            except antlr.TryAgain:
                pass
        
    def mADD(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = ADD
        _saveIndex = 0
        pass
        self.match('+')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mSUB(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = SUB
        _saveIndex = 0
        pass
        self.match('-')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mMUL(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = MUL
        _saveIndex = 0
        pass
        self.match('*')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mDIV(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = DIV
        _saveIndex = 0
        pass
        self.match('/')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mEQUAL(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = EQUAL
        _saveIndex = 0
        pass
        self.match('=')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mNOT_EQUAL(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = NOT_EQUAL
        _saveIndex = 0
        pass
        self.match("<>")
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLT(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LT
        _saveIndex = 0
        pass
        self.match('<')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLE(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LE
        _saveIndex = 0
        pass
        self.match("<=")
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mGE(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = GE
        _saveIndex = 0
        pass
        self.match(">=")
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mGT(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = GT
        _saveIndex = 0
        pass
        self.match('>')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLP(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LP
        _saveIndex = 0
        pass
        self.match('(')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mRP(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = RP
        _saveIndex = 0
        pass
        self.match(')')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLB(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LB
        _saveIndex = 0
        pass
        self.match('[')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mRB(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = RB
        _saveIndex = 0
        pass
        self.match(']')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mCOLON(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = COLON
        _saveIndex = 0
        pass
        self.match(':')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mCOMMA(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = COMMA
        _saveIndex = 0
        pass
        self.match(',')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mSEMICOLON(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = SEMICOLON
        _saveIndex = 0
        pass
        self.match(';')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mWHITE_SPACE(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = WHITE_SPACE
        _saveIndex = 0
        pass
        _cnt34= 0
        while True:
            la1 = self.LA(1)
            if False:
                pass
            elif la1 and la1 in u' ':
                pass
                self.match(' ')
            elif la1 and la1 in u'\t':
                pass
                self.match('\t')
            elif la1 and la1 in u'\u000c':
                pass
                self.match('\f')
            elif la1 and la1 in u'\r':
                pass
                self.match('\r')
            elif la1 and la1 in u'\n':
                pass
                self.match('\n')
            else:
                    break
                
            _cnt34 += 1
        if _cnt34 < 1:
            self.raise_NoViableAlt(self.LA(1))
        _ttype = SKIP
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mDIGIT(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = DIGIT
        _saveIndex = 0
        pass
        self.matchRange(u'0', u'9')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mDIGITS(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = DIGITS
        _saveIndex = 0
        pass
        _cnt38= 0
        while True:
            if ((self.LA(1) >= u'0' and self.LA(1) <= u'9')):
                pass
                self.mDIGIT(False)
            else:
                break
            
            _cnt38 += 1
        if _cnt38 < 1:
            self.raise_NoViableAlt(self.LA(1))
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mINT_CONST(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = INT_CONST
        _saveIndex = 0
        pass
        self.mDIGITS(False)
        if (self.LA(1)==u'.'):
            pass
            self.match(".")
            self.mDIGITS(False)
            if (self.LA(1)==u'E' or self.LA(1)==u'e'):
                pass
                la1 = self.LA(1)
                if False:
                    pass
                elif la1 and la1 in u'e':
                    pass
                    self.match('e')
                elif la1 and la1 in u'E':
                    pass
                    self.match('E')
                else:
                        self.raise_NoViableAlt(self.LA(1))
                    
                la1 = self.LA(1)
                if False:
                    pass
                elif la1 and la1 in u'+':
                    pass
                    self.match('+')
                elif la1 and la1 in u'-':
                    pass
                    self.match('-')
                elif la1 and la1 in u'0123456789':
                    pass
                else:
                        self.raise_NoViableAlt(self.LA(1))
                    
                self.mDIGITS(False)
            else: ## <m4>
                    pass
                
            _ttype = NUM_CONST
        else: ## <m4>
                pass
            
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLOWER(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LOWER
        _saveIndex = 0
        pass
        self.matchRange(u'a', u'z')
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mUPPER(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = UPPER
        _saveIndex = 0
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in u'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            pass
            self.matchRange(u'A', u'Z')
        elif la1 and la1 in u'_':
            pass
            self.match('_')
        else:
                self.raise_NoViableAlt(self.LA(1))
            
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mLETTER(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = LETTER
        _saveIndex = 0
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in u'ABCDEFGHIJKLMNOPQRSTUVWXYZ_':
            pass
            self.mUPPER(False)
        elif la1 and la1 in u'abcdefghijklmnopqrstuvwxyz':
            pass
            self.mLOWER(False)
        else:
                self.raise_NoViableAlt(self.LA(1))
            
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mRC_CELL(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = RC_CELL
        _saveIndex = 0
        pass
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in u'r':
            pass
            self.match('r')
        elif la1 and la1 in u'R':
            pass
            self.match('R')
        else:
                self.raise_NoViableAlt(self.LA(1))
            
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in u'C[c':
            pass
            la1 = self.LA(1)
            if False:
                pass
            elif la1 and la1 in u'[':
                pass
                self.mLB(False)
                la1 = self.LA(1)
                if False:
                    pass
                elif la1 and la1 in u'-':
                    pass
                    self.mSUB(False)
                elif la1 and la1 in u'0123456789':
                    pass
                else:
                        self.raise_NoViableAlt(self.LA(1))
                    
                self.mDIGITS(False)
                self.mRB(False)
            elif la1 and la1 in u'Cc':
                pass
            else:
                    self.raise_NoViableAlt(self.LA(1))
                
        elif la1 and la1 in u'0123456789':
            pass
            self.mDIGITS(False)
        else:
                self.raise_NoViableAlt(self.LA(1))
            
        la1 = self.LA(1)
        if False:
            pass
        elif la1 and la1 in u'c':
            pass
            self.match('c')
        elif la1 and la1 in u'C':
            pass
            self.match('C')
        else:
                self.raise_NoViableAlt(self.LA(1))
            
        if ((self.LA(1) >= u'0' and self.LA(1) <= u'9')):
            pass
            self.mDIGITS(False)
        else: ## <m4>
                pass
                if (self.LA(1)==u'['):
                    pass
                    self.mLB(False)
                    la1 = self.LA(1)
                    if False:
                        pass
                    elif la1 and la1 in u'-':
                        pass
                        self.mSUB(False)
                    elif la1 and la1 in u'0123456789':
                        pass
                    else:
                            self.raise_NoViableAlt(self.LA(1))
                        
                    self.mDIGITS(False)
                    self.mRB(False)
                else: ## <m4>
                        pass
                    
            
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    def mNAME(self, _createToken):    
        _ttype = 0
        _token = None
        _begin = self.text.length()
        _ttype = NAME
        _saveIndex = 0
        pass
        self.mLETTER(False)
        while True:
            la1 = self.LA(1)
            if False:
                pass
            elif la1 and la1 in u'ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz':
                pass
                self.mLETTER(False)
            elif la1 and la1 in u'0123456789':
                pass
                self.mDIGIT(False)
            else:
                    break
                
        self.set_return_token(_createToken, _token, _ttype, _begin)
    
    

### generate bit set
def mk_tokenSet_0(): 
    ### var1
    data = [ 287948901175001088L, 34493956104L, 0L, 0L]
    return data
_tokenSet_0 = antlr.BitSet(mk_tokenSet_0())

### generate bit set
def mk_tokenSet_1(): 
    ### var1
    data = [ 0L, 576460745995190270L, 0L, 0L]
    return data
_tokenSet_1 = antlr.BitSet(mk_tokenSet_1())
    
### __main__ header action >>> 
if __name__ == '__main__' :
    import sys
    import antlr
    import ExcelFormulaLexer
    
    ### create lexer - shall read from stdin
    try:
        for token in ExcelFormulaLexer.Lexer():
            print token
            
    except antlr.TokenStreamException, e:
        print "error: exception caught while lexing: ", e
### __main__ header action <<< 
