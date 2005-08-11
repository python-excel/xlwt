/*
 *  $Id$
 */

header {
    __rev_id__ = """$Id$"""

    import struct
    import Utils
    from UnicodeUtils import upack1
    from ExcelMagic import *
}

header "ExcelFormulaParser.__init__" {
    self.rpn = ""
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
    : term (op = term_op term { self.rpn += op } )*
    ;

term_op returns [result]
    : ADD { result = struct.pack('B', ptgAdd) }
    | SUB { result = struct.pack('B', ptgSub) }
    | CONCAT { result = struct.pack('B', ptgConcat) }
    ;

term
    : fact (op = fact_op fact { self.rpn += op } )*
    ;

fact_op returns [result]
    : MUL { result = struct.pack('B', ptgMul) }
    | DIV { result = struct.pack('B', ptgDiv) }
    ;

fact
    : primary
    | SUB primary { self.rpn += struct.pack('B', ptgUminus) }
    ;

primary
    : str_tok:STR_CONST                     
        { 
            self.rpn += struct.pack("B", ptgStr) + upack1(str_tok.text[1:-1])
        }
    | int_tok:INT_CONST                     
        { 
            self.rpn += struct.pack("<BH", ptgInt, int(int_tok.text)) 
        }
    | num_tok:NUM_CONST                     
        { 
            self.rpn += struct.pack("<Bd", ptgNum, float(num_tok.text)) 
        }
    | ref2d_tok:REF2D                       
        { 
            r, c = Utils.cell_to_packed_rowcol(ref2d_tok.text)
            self.rpn += struct.pack("<B2H", ptgRefV, r, c) 
        }
    | ref2d1_tok:REF2D COLON ref2d2_tok:REF2D
        { 
            r1, c1 = Utils.cell_to_packed_rowcol(ref2d1_tok.text)
            r2, c2 = Utils.cell_to_packed_rowcol(ref2d2_tok.text)
            self.rpn += struct.pack("<B4H", ptgAreaV, r1, r2, c1, c2)
        }
    | name_tok:NAME                         
        { 
            self.rpn += "" 
        }
    | func_tok:NAME LP arg_count = expr_list RP         
        {   
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
        }
    | LP expr RP                            
        { 
            self.rpn += struct.pack("B", ptgParen) 
        }
    ;

expr_list returns [arg_cnt]
{ arg_cnt = 0 }
    : expr { arg_cnt += 1 } (SEMICOLON expr { arg_cnt += 1 })*
    | 
    ; 
