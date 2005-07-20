/*
 *  $Id$
 */

header {
    __rev_id__ = """$Id$"""
    from ExcelFormula import *
}

header "ExcelFormulaParser.__init__" {
    self.postfix_stack = []
}

options {
    language  = "Python";
}

class ExcelFormulaParser extends Parser;
options {
    k = 2;
    defaultErrorHandler = false;
    buildAST = false;
}


tokens {
    STR_CONST;
    NUM_CONST;

    NAME;
    
    ASSIGN;
    
    ADD;
    SUB;
    MUL;
    DIV;

    LP;
    RP;

    LB;
    RB;

    COLON;
    COMMA;
    SEMICOLON;
}

formula 
    : expr
    ;

expr
    : term (op = term_op term { self.postfix_stack.extend([op]) })*
    ;

term_op returns [result]
    : ADD { result = (ptgAdd, "ptgAdd") }
    | SUB { result = (ptgSub, "ptgSub") }
    ;

term
    : fact (op = fact_op fact { self.postfix_stack.extend([op]) })*
    ;

fact_op returns [result]
    : MUL { result = (ptgMul, "ptgMul") }
    | DIV { result = (ptgDiv, "ptgDiv") }
    ;

fact
    : primary
    | SUB primary { self.postfix_stack.extend([(ptgUminus, "ptgUminus")])}
    ;

primary
    : str_tok:STR_CONST                     { self.postfix_stack.extend([(ptgStr, str_tok.text)]) }
    | int_tok:INT_CONST                     { self.postfix_stack.extend([(ptgInt, int(int_tok.text))]) }
    | num_tok:NUM_CONST                     { self.postfix_stack.extend([(ptgNum, float(num_tok.text))]) }
    | rc_cell:RC_CELL                       { self.postfix_stack.extend([(ptgRefV, rc_cell.text)]) }
    | rc0:RC_CELL COLON rc1:RC_CELL         { self.postfix_stack.extend([(ptgRange, "%s : %s" % (rc0.text, rc1.text))]) }
    | name_tok:NAME                         { self.postfix_stack.extend([(ptgName, name_tok.text)]) }
    | func_tok:NAME LP expr_list RP { self.postfix_stack.extend([(ptgFunc, func_tok.text)]) }
    | LP expr RP
    ;

expr_list
    : expr (SEMICOLON expr)*
    |
    ; 



class ExcelFormulaLexer extends Lexer;
options {
    k = 2;
}

ADD         : '+' ;
SUB         : '-' ;
MUL         : '*' ;
DIV         : '/' ;
EQUAL       : '=' ;
NOT_EQUAL   : "<>";
LT          : '<' ;
LE          : "<=";
GE          : ">=";
GT          : '>' ;
LP          : '(' ;
RP          : ')' ;
LB          : '[' ;
RB          : ']' ;
COLON       : ':' ;
COMMA       : ',' ;
SEMICOLON   : ';' ;

WHITE_SPACE
    : (' ' | '\t' | '\f' | '\r' | '\n')+ { $skip }
    ;

protected
DIGIT
    : '0'..'9'
    ;

protected
INT
    : (DIGIT)+
    ;

NUM_CONST        
    : INT "." INT
    ;

INT_CONST
    : INT
    ;

protected
LOWER
    : 'a'..'z'
    ;

protected
UPPER
    : 'A'..'Z'
    | '_'
    ;

protected
LETTER
    : UPPER
    | LOWER
    ;

RC_CELL
    : ('r' | 'R') ((LB (SUB)? INT RB)? | INT) ('c' | 'C') ((LB (SUB)? INT RB)? | INT)
    ;

NAME
    : LETTER (LETTER | DIGIT)*
    ;

