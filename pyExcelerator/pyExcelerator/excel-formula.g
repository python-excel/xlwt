/*
 *  $Id$
 */

header {
    __rev_id__ = """$Id$"""

    import struct
}

header "ExcelFormulaParser.__init__" {
    self.rpn = ""
}

options {
    language  = "Python";
}
{
# Parse things

ptgExp          = 0x01
ptgTbl          = 0x02
ptgAdd          = 0x03
ptgSub          = 0x04
ptgMul          = 0x05
ptgDiv          = 0x06
ptgPower        = 0x07
ptgConcat       = 0x08
ptgLT           = 0x09
ptgLE           = 0x0a
ptgEQ           = 0x0b
ptgGE           = 0x0c
ptgGT           = 0x0d
ptgNE           = 0x0e
ptgIsect        = 0x0f
ptgUnion        = 0x10
ptgRange        = 0x11
ptgUplus        = 0x12
ptgUminus       = 0x13
ptgPercent      = 0x14
ptgParen        = 0x15
ptgMissArg      = 0x16
ptgStr          = 0x17
ptgExtend       = 0x18
ptgAttr         = 0x19
ptgSheet        = 0x1a
ptgEndSheet     = 0x1b
ptgErr          = 0x1c
ptgBool         = 0x1d
ptgInt          = 0x1e
ptgNum          = 0x1f

ptgArray        = 0x20
ptgFunc         = 0x21
ptgFuncVar      = 0x22
ptgName         = 0x23
ptgRef          = 0x24
ptgArea         = 0x25
ptgMemArea      = 0x26
ptgMemErr       = 0x27
ptgMemNoMem     = 0x28
ptgMemFunc      = 0x29
ptgRefErr       = 0x2a
ptgAreaErr      = 0x2b
ptgRefN         = 0x2c
ptgAreaN        = 0x2d
ptgMemAreaN     = 0x2e
ptgMemNoMemN    = 0x2f
ptgNameX        = 0x39
ptgRef3d        = 0x3a
ptgArea3d       = 0x3b
ptgRefErr3d     = 0x3c
ptgAreaErr3d    = 0x3d

ptgArrayV       = 0x40
ptgFuncV        = 0x41
ptgFuncVarV     = 0x42
ptgNameV        = 0x43
ptgRefV         = 0x44
ptgAreaV        = 0x45
ptgMemAreaV     = 0x46
ptgMemErrV      = 0x47
ptgMemNoMemV    = 0x48
ptgMemFuncV     = 0x49
ptgRefErrV      = 0x4a
ptgAreaErrV     = 0x4b
ptgRefNV        = 0x4c
ptgAreaNV       = 0x4d
ptgMemAreaNV    = 0x4e
ptgMemNoMemNV   = 0x4f
ptgFuncCEV      = 0x58
ptgNameXV       = 0x59
ptgRef3dV       = 0x5a
ptgArea3dV      = 0x5b
ptgRefErr3dV    = 0x5c
ptgAreaErr3dV   = 0x5d

ptgArrayA       = 0x60
ptgFuncA        = 0x61
ptgFuncVarA     = 0x62
ptgNameA        = 0x63
ptgRefA         = 0x64
ptgAreaA        = 0x65
ptgMemAreaA     = 0x66
ptgMemErrA      = 0x67
ptgMemNoMemA    = 0x68
ptgMemFuncA     = 0x69
ptgRefErrA      = 0x6a
ptgAreaErrA     = 0x6b
ptgRefNA        = 0x6c
ptgAreaNA       = 0x6d
ptgMemAreaNA    = 0x6e
ptgMemNoMemNA   = 0x6f
ptgFuncCEA      = 0x78
ptgNameXA       = 0x79
ptgRef3dA       = 0x7a
ptgArea3dA      = 0x7b
ptgRefErr3dA    = 0x7c
ptgAreaErr3dA   = 0x7d


PtgNames = {
    ptgExp         : "ptgExp",
    ptgTbl         : "ptgTbl",
    ptgAdd         : "ptgAdd",
    ptgSub         : "ptgSub",
    ptgMul         : "ptgMul",
    ptgDiv         : "ptgDiv",
    ptgPower       : "ptgPower",
    ptgConcat      : "ptgConcat",
    ptgLT          : "ptgLT",
    ptgLE          : "ptgLE",
    ptgEQ          : "ptgEQ",
    ptgGE          : "ptgGE",
    ptgGT          : "ptgGT",
    ptgNE          : "ptgNE",
    ptgIsect       : "ptgIsect",
    ptgUnion       : "ptgUnion",
    ptgRange       : "ptgRange",
    ptgUplus       : "ptgUplus",
    ptgUminus      : "ptgUminus",
    ptgPercent     : "ptgPercent",
    ptgParen       : "ptgParen",
    ptgMissArg     : "ptgMissArg",
    ptgStr         : "ptgStr",
    ptgExtend      : "ptgExtend",
    ptgAttr        : "ptgAttr",
    ptgSheet       : "ptgSheet",
    ptgEndSheet    : "ptgEndSheet",
    ptgErr         : "ptgErr",
    ptgBool        : "ptgBool",
    ptgInt         : "ptgInt",
    ptgNum         : "ptgNum",
    ptgArray       : "ptgArray",
    ptgFunc        : "ptgFunc",
    ptgFuncVar     : "ptgFuncVar",
    ptgName        : "ptgName",
    ptgRef         : "ptgRef",
    ptgArea        : "ptgArea",
    ptgMemArea     : "ptgMemArea",
    ptgMemErr      : "ptgMemErr",
    ptgMemNoMem    : "ptgMemNoMem",
    ptgMemFunc     : "ptgMemFunc",
    ptgRefErr      : "ptgRefErr",
    ptgAreaErr     : "ptgAreaErr",
    ptgRefN        : "ptgRefN",
    ptgAreaN       : "ptgAreaN",
    ptgMemAreaN    : "ptgMemAreaN",
    ptgMemNoMemN   : "ptgMemNoMemN",
    ptgNameX       : "ptgNameX",
    ptgRef3d       : "ptgRef3d",
    ptgArea3d      : "ptgArea3d",
    ptgRefErr3d    : "ptgRefErr3d",
    ptgAreaErr3d   : "ptgAreaErr3d",
    ptgArrayV      : "ptgArrayV",
    ptgFuncV       : "ptgFuncV",
    ptgFuncVarV    : "ptgFuncVarV",
    ptgNameV       : "ptgNameV",
    ptgRefV        : "ptgRefV",
    ptgAreaV       : "ptgAreaV",
    ptgMemAreaV    : "ptgMemAreaV",
    ptgMemErrV     : "ptgMemErrV",
    ptgMemNoMemV   : "ptgMemNoMemV",
    ptgMemFuncV    : "ptgMemFuncV",
    ptgRefErrV     : "ptgRefErrV",
    ptgAreaErrV    : "ptgAreaErrV",
    ptgRefNV       : "ptgRefNV",
    ptgAreaNV      : "ptgAreaNV",
    ptgMemAreaNV   : "ptgMemAreaNV",
    ptgMemNoMemNV  : "ptgMemNoMemNV",
    ptgFuncCEV     : "ptgFuncCEV",
    ptgNameXV      : "ptgNameXV",
    ptgRef3dV      : "ptgRef3dV",
    ptgArea3dV     : "ptgArea3dV",
    ptgRefErr3dV   : "ptgRefErr3dV",
    ptgAreaErr3dV  : "ptgAreaErr3dV",
    ptgArrayA      : "ptgArrayA",
    ptgFuncA       : "ptgFuncA",
    ptgFuncVarA    : "ptgFuncVarA",
    ptgNameA       : "ptgNameA",
    ptgRefA        : "ptgRefA",
    ptgAreaA       : "ptgAreaA",
    ptgMemAreaA    : "ptgMemAreaA",
    ptgMemErrA     : "ptgMemErrA",
    ptgMemNoMemA   : "ptgMemNoMemA",
    ptgMemFuncA    : "ptgMemFuncA",
    ptgRefErrA     : "ptgRefErrA",
    ptgAreaErrA    : "ptgAreaErrA",
    ptgRefNA       : "ptgRefNA",
    ptgAreaNA      : "ptgAreaNA",
    ptgMemAreaNA   : "ptgMemAreaNA",
    ptgMemNoMemNA  : "ptgMemNoMemNA",
    ptgFuncCEA     : "ptgFuncCEA",
    ptgNameXA      : "ptgNameXA",
    ptgRef3dA      : "ptgRef3dA",
    ptgArea3dA     : "ptgArea3dA",
    ptgRefErr3dA   : "ptgRefErr3dA",
    ptgAreaErr3dA  : "ptgAreaErr3dA"
}


# Functions

ocCount         =   0
ocIf            =   1
ocIsNV          =   2
ocIsError       =   3
ocSum           =   4
ocAverage       =   5
ocMin           =   6
ocMax           =   7
ocRow           =   8
ocColumn        =   9
ocNoValue       =  10
ocNBW           =  11
ocStDev         =  12
ocCurrency      =  13
ocFixed         =  14
ocSin           =  15
ocCos           =  16
ocTan           =  17
ocArcTan        =  18
ocPi            =  19
ocSqrt          =  20
ocExp           =  21
ocLn            =  22
ocLog10         =  23
ocAbs           =  24
ocInt           =  25
ocPlusMinus     =  26
ocRound         =  27
ocLookup        =  28
ocIndex         =  29
ocRept          =  30
ocMid           =  31
ocLen           =  32
ocValue         =  33
ocTrue          =  34
ocFalse         =  35
ocAnd           =  36
ocOr            =  37
ocNot           =  38
ocMod           =  39
ocDBCount       =  40
ocDBSum         =  41
ocDBAverage     =  42
ocDBMin         =  43
ocDBMax         =  44
ocDBStdDev      =  45
ocVar           =  46
ocDBVar         =  47
ocText          =  48
ocRGP           =  49
ocTrend         =  50
ocRKP           =  51
ocGrowth        =  52
ocBW            =  56
ocZW            =  57
ocZZR           =  58
ocRMZ           =  59
ocZins          =  60
ocMIRR          =  61
ocIKV           =  62
ocRandom        =  63
ocMatch         =  64
ocGetDate       =  65
ocGetTime       =  66
ocGetDay        =  67
ocGetMonth      =  68
ocGetYear       =  69
ocGetDayOfWeek  =  70
ocGetHour       =  71
ocGetMin        =  72
ocGetSec        =  73
ocGetActTime    =  74
ocAreas         =  75
ocRows          =  76
ocColumns       =  77
ocOffset        =  78
ocSearch        =  82
ocMatTrans      =  83
ocType          =  86
ocArcTan2       =  97
ocArcSin        =  98
ocArcCos        =  99
ocChose         = 100
ocHLookup       = 101
ocVLookup       = 102
ocIsRef         = 105
ocLog           = 109
ocChar          = 111
ocLower         = 112
ocUpper         = 113
ocPropper       = 114
ocLeft          = 115
ocRight         = 116
ocExact         = 117
ocTrim          = 118
ocReplace       = 119
ocSubstitute    = 120
ocCode          = 121
ocFind          = 124
ocCell          = 125
ocIsErr         = 126
ocIsString      = 127
ocIsValue       = 128
ocIsEmpty       = 129
ocT             = 130
ocN             = 131
ocGetDateValue  = 140
ocGetTimeValue  = 141
ocLIA           = 142
ocDIA           = 143
ocGDA           = 144
ocIndirect      = 148
ocClean         = 162
ocMatDet        = 163
ocMatInv        = 164
ocMatMult       = 165
ocZinsZ         = 167
ocKapz          = 168
ocCount2        = 169
ocProduct       = 183
ocFact          = 184
ocDBProduct     = 189
ocIsNonString   = 190
ocStDevP        = 193
ocVarP          = 194
ocDBStdDevP     = 195
ocDBVarP        = 196
ocTrunc         = 197
ocIsLogical     = 198
ocDBCount2      = 199
ocRoundUp       = 212
ocRoundDown     = 213
ocRank          = 216
ocAdress        = 219
ocGetDiffDate360= 220
ocGetActDate    = 221
ocVBD           = 222
ocMedian        = 227
ocSumProduct    = 228
ocSinHyp        = 229
ocCosHyp        = 230
ocTanHyp        = 231
ocArcSinHyp     = 232
ocArcCosHyp     = 233
ocArcTanHyp     = 234
ocDBGet         = 235
ocGDA2          = 247
ocFrequency     = 252
ocErrorType     = 261
ocAveDev        = 269
ocBetaDist      = 270
ocGammaLn       = 271
ocBetaInv       = 272
ocBinomDist     = 273
ocChiDist       = 274
ocChiInv        = 275
ocKombin        = 276
ocConfidence    = 277
ocKritBinom     = 278
ocEven          = 279
ocExpDist       = 280
ocFDist         = 281
ocFInv          = 282
ocFisher        = 283
ocFisherInv     = 284
ocFloor         = 285
ocGammaDist     = 286
ocGammaInv      = 287
ocCeil          = 288
ocHypGeomDist   = 289
ocLogNormDist   = 290
ocLogInv        = 291
ocNegBinomVert  = 292
ocNormDist      = 293
ocStdNormDist   = 294
ocNormInv       = 295
ocSNormInv      = 296
ocStandard      = 297
ocOdd           = 298
ocVariationen   = 299
ocPoissonDist   = 300
ocTDist         = 301
ocWeibull       = 302
ocSumXMY2       = 303
ocSumX2MY2      = 304
ocSumX2DY2      = 305
ocChiTest       = 306
ocCorrel        = 307
ocCovar         = 308
ocForecast      = 309
ocFTest         = 310
ocIntercept     = 311
ocPearson       = 312
ocRSQ           = 313
ocSTEYX         = 314
ocSlope         = 315
ocTTest         = 316
ocProb          = 317
ocDevSq         = 318
ocGeoMean       = 319
ocHarMean       = 320
ocSumSQ         = 321
ocKurt          = 322
ocSchiefe       = 323
ocZTest         = 324
ocLarge         = 325
ocSmall         = 326
ocQuartile      = 327
ocPercentile    = 328
ocPercentrank   = 329
ocModalValue    = 330
ocTrimMean      = 331
ocTInv          = 332
ocConcat        = 336
ocPower         = 337
ocRad           = 342
ocDeg           = 343
ocSubTotal      = 344
ocSumIf         = 345
ocCountIf       = 346
ocCountEmptyCells=347
ocRoman         = 354

# from BIFF8
ocExternal      = 255
ocMacro         = 255
ocISPMT         = 350
ocAverageA      = 361
ocMaxA          = 362
ocMinA          = 363
ocStDevPA       = 364
ocVarPA         = 365
ocStDevA        = 366
ocVarA          = 367

nMaxParamCount   = 5
nVarParam        = -1
nTypeErr         = 0 
nTypeRef         = 1
nTypeVal         = 2
nTypeArr         = 3
nTypeIgn         = 4 
nTypeIns         = 5 


#                  Opcode                | FuncType | Volatile  | ParamCount | VarParamTypeCount | PType1,PType2 ...
std_functions = {
    "COUNT"           : [ocCount          , nTypeVal, False,  -1,   1, nTypeRef],
    "IF"              : [ocIf             , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "ISNV"            : [ocIsNV           , nTypeVal, False,   1,   1, nTypeVal],
    "ISERROR"         : [ocIsError        , nTypeVal, False,   1,   1, nTypeVal],
    "SUM"             : [ocSum            , nTypeVal, False,  -1,   1, nTypeRef],
    "AVERAGE"         : [ocAverage        , nTypeVal, False,  -1,   1, nTypeRef],
    "MIN"             : [ocMin            , nTypeVal, False,  -1,   1, nTypeRef],
    "MAX"             : [ocMax            , nTypeVal, False,  -1,   1, nTypeRef],
    "ROW"             : [ocRow            , nTypeVal, False,  -1,   1, nTypeRef],
    "COLUMN"          : [ocColumn         , nTypeVal, False,  -1,   1, nTypeRef],
    "NOVALUE"         : [ocNoValue        , nTypeVal, False,   0,   1, 0],
    "NBW"             : [ocNBW            , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "STDEV"           : [ocStDev          , nTypeVal, False,  -1,   1, nTypeRef],
    "CURRENCY"        : [ocCurrency       , nTypeVal, False,  -1,   1, nTypeVal],
    "FIXED"           : [ocFixed          , nTypeVal, False,  -1,   1, nTypeVal],
    "SIN"             : [ocSin            , nTypeVal, False,   1,   1, nTypeVal],
    "COS"             : [ocCos            , nTypeVal, False,   1,   1, nTypeVal],
    "TAN"             : [ocTan            , nTypeVal, False,   1,   1, nTypeVal],
    "ARCTAN"          : [ocArcTan         , nTypeVal, False,   1,   1, nTypeVal],
    "PI"              : [ocPi             , nTypeVal, False,   0,   1, 0],
    "SQRT"            : [ocSqrt           , nTypeVal, False,   1,   1, nTypeVal],
    "EXP"             : [ocExp            , nTypeVal, False,   1,   1, nTypeVal],
    "LN"              : [ocLn             , nTypeVal, False,   1,   1, nTypeVal],
    "LOG10"           : [ocLog10          , nTypeVal, False,   1,   1, nTypeVal],
    "ABS"             : [ocAbs            , nTypeVal, False,   1,   1, nTypeVal],
    "INT"             : [ocInt            , nTypeVal, False,   1,   1, nTypeVal],
    "PLUSMINUS"       : [ocPlusMinus      , nTypeVal, False,   1,   1, nTypeVal],
    "ROUND"           : [ocRound          , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "LOOKUP"          : [ocLookup         , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "INDEX"           : [ocIndex          , nTypeRef,  True,  -1,   4, nTypeRef, nTypeVal, nTypeVal, nTypeVal],
    "REPT"            : [ocRept           , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "MID"             : [ocMid            , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "LEN"             : [ocLen            , nTypeVal, False,   1,   1, nTypeVal],
    "VALUE"           : [ocValue          , nTypeVal, False,   1,   1, nTypeVal],
    "TRUE"            : [ocTrue           , nTypeVal, False,   0,   1, 0],
    "FALSE"           : [ocFalse          , nTypeVal, False,   0,   1, 0],
    "AND"             : [ocAnd            , nTypeVal, False,  -1,   1, nTypeRef],
    "OR"              : [ocOr             , nTypeVal, False,  -1,   1, nTypeRef],
    "NOT"             : [ocNot            , nTypeVal, False,   1,   1, nTypeVal],
    "MOD"             : [ocMod            , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "DBCOUNT"         : [ocDBCount        , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBSUM"           : [ocDBSum          , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBAVERAGE"       : [ocDBAverage      , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBMIN"           : [ocDBMin          , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBMAX"           : [ocDBMax          , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBSTDDEV"        : [ocDBStdDev       , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "VAR"             : [ocVar            , nTypeVal, False,  -1,   1, nTypeRef],
    "DBVAR"           : [ocDBVar          , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "TEXT"            : [ocText           , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "RGP"             : [ocRGP            , nTypeVal, False,  -1,   1, nTypeRef],
    "TREND"           : [ocTrend          , nTypeVal, False,  -1,   1, nTypeRef],
    "RKP"             : [ocRKP            , nTypeVal, False,  -1,   1, nTypeRef],
    "GROWTH"          : [ocGrowth         , nTypeVal, False,  -1,   1, nTypeRef],
    "BW"              : [ocBW             , nTypeVal, False,  -1,   1, nTypeVal],
    "ZW"              : [ocZW             , nTypeVal, False,  -1,   1, nTypeVal],
    "ZZR"             : [ocZZR            , nTypeVal, False,  -1,   1, nTypeVal],
    "RMZ"             : [ocRMZ            , nTypeVal, False,  -1,   1, nTypeVal],
    "ZINS"            : [ocZins           , nTypeVal, False,  -1,   1, nTypeVal],
    "MIRR"            : [ocMIRR           , nTypeVal, False,   3,   1, nTypeRef, nTypeVal, nTypeVal],
    "IKV"             : [ocIKV            , nTypeVal, False,  -1,   2, nTypeRef, nTypeVal],
    "RANDOM"          : [ocRandom         , nTypeVal,  True,   0,   1, 0],
    "MATCH"           : [ocMatch          , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "GETDATE"         : [ocGetDate        , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "GETTIME"         : [ocGetTime        , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "GETDAY"          : [ocGetDay         , nTypeVal, False,   1,   1, nTypeVal],
    "GETMONTH"        : [ocGetMonth       , nTypeVal, False,   1,   1, nTypeVal],
    "GETYEAR"         : [ocGetYear        , nTypeVal, False,   1,   1, nTypeVal],
    "GETDAYOFWEEK"    : [ocGetDayOfWeek   , nTypeVal, False,  -1,   1, nTypeVal],
    "GETHOUR"         : [ocGetHour        , nTypeVal, False,   1,   1, nTypeVal],
    "GETMIN"          : [ocGetMin         , nTypeVal, False,   1,   1, nTypeVal],
    "GETSEC"          : [ocGetSec         , nTypeVal, False,   1,   1, nTypeVal],
    "GETACTTIME"      : [ocGetActTime     , nTypeVal,  True,   0,   1, 0],
    "AREAS"           : [ocAreas          , nTypeVal, False,   1,   1, nTypeRef],
    "ROWS"            : [ocRows           , nTypeVal, False,   1,   1, nTypeRef],
    "COLUMNS"         : [ocColumns        , nTypeVal, False,   1,   1, nTypeRef],
    "OFFSET"          : [ocOffset         , nTypeRef, False,  -1,   2, nTypeRef, nTypeVal],
    "SEARCH"          : [ocSearch         , nTypeVal, False,  -1,   1, nTypeVal],
    "MATTRANS"        : [ocMatTrans       , nTypeVal, False,   1,   1, nTypeArr],
    "TYPE"            : [ocType           , nTypeVal, False,   1,   1, nTypeVal],
    "ARCTAN2"         : [ocArcTan2        , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "ARCSIN"          : [ocArcSin         , nTypeVal, False,   1,   1, nTypeVal],
    "ARCCOS"          : [ocArcCos         , nTypeVal, False,   1,   1, nTypeVal],
    "CHOSE"           : [ocChose          , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "HLOOKUP"         : [ocHLookup        , nTypeVal, False,  -1,   4, nTypeVal, nTypeRef, nTypeRef, nTypeVal],
    "VLOOKUP"         : [ocVLookup        , nTypeVal, False,  -1,   4, nTypeVal, nTypeRef, nTypeRef, nTypeVal],
    "ISREF"           : [ocIsRef          , nTypeVal, False,   1,   1, nTypeRef],
    "LOG"             : [ocLog            , nTypeVal, False,  -1,   1, nTypeVal],
    "CHAR"            : [ocChar           , nTypeVal, False,   1,   1, nTypeVal],
    "LOWER"           : [ocLower          , nTypeVal, False,   1,   1, nTypeVal],
    "UPPER"           : [ocUpper          , nTypeVal, False,   1,   1, nTypeVal],
    "PROPPER"         : [ocPropper        , nTypeVal, False,   1,   1, nTypeVal],
    "LEFT"            : [ocLeft           , nTypeVal, False,  -1,   1, nTypeVal],
    "RIGHT"           : [ocRight          , nTypeVal, False,  -1,   1, nTypeVal],
    "EXACT"           : [ocExact          , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "TRIM"            : [ocTrim           , nTypeVal, False,   1,   1, nTypeVal],
    "REPLACE"         : [ocReplace        , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "SUBSTITUTE"      : [ocSubstitute     , nTypeVal, False,  -1,   1, nTypeVal],
    "CODE"            : [ocCode           , nTypeVal, False,   1,   1, nTypeVal],
    "FIND"            : [ocFind           , nTypeVal, False,  -1,   1, nTypeVal],
    "CELL"            : [ocCell           , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "ISERR"           : [ocIsErr          , nTypeVal, False,   1,   1, nTypeVal],
    "ISSTRING"        : [ocIsString       , nTypeVal, False,   1,   1, nTypeVal],
    "ISVALUE"         : [ocIsValue        , nTypeVal, False,   1,   1, nTypeVal],
    "ISEMPTY"         : [ocIsEmpty        , nTypeVal, False,   1,   1, nTypeVal],
    "T"               : [ocT              , nTypeVal, False,   1,   1, nTypeRef],
    "N"               : [ocN              , nTypeVal, False,   1,   1, nTypeRef],
    "GETDATEVALUE"    : [ocGetDateValue   , nTypeVal, False,   1,   1, nTypeVal],
    "GETTIMEVALUE"    : [ocGetTimeValue   , nTypeVal, False,   1,   1, nTypeVal],
    "LIA"             : [ocLIA            , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "DIA"             : [ocDIA            , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "GDA"             : [ocGDA            , nTypeVal, False,  -1,   1, nTypeVal],
    "INDIRECT"        : [ocIndirect       , nTypeRef,  True,  -1,   1, nTypeVal],
    "CLEAN"           : [ocClean          , nTypeVal, False,   1,   1, nTypeVal],
    "MATDET"          : [ocMatDet         , nTypeVal, False,   1,   1, nTypeArr],
    "MATINV"          : [ocMatInv         , nTypeVal, False,   1,   1, nTypeArr],
    "MATMULT"         : [ocMatMult        , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "ZINSZ"           : [ocZinsZ          , nTypeVal, False,  -1,   1, nTypeVal],
    "KAPZ"            : [ocKapz           , nTypeVal, False,  -1,   1, nTypeVal],
    "COUNT2"          : [ocCount2         , nTypeVal, False,  -1,   1, nTypeRef],
    "PRODUCT"         : [ocProduct        , nTypeVal, False,  -1,   1, nTypeRef],
    "FACT"            : [ocFact           , nTypeVal, False,   1,   1, nTypeVal],
    "DBPRODUCT"       : [ocDBProduct      , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "ISNONSTRING"     : [ocIsNonString    , nTypeVal, False,   1,   1, nTypeVal],
    "STDEVP"          : [ocStDevP         , nTypeVal, False,  -1,   1, nTypeRef],
    "VARP"            : [ocVarP           , nTypeVal, False,  -1,   1, nTypeRef],
    "DBSTDDEVP"       : [ocDBStdDevP      , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "DBVARP"          : [ocDBVarP         , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "TRUNC"           : [ocTrunc          , nTypeVal, False,  -1,   1, nTypeVal],
    "ISLOGICAL"       : [ocIsLogical      , nTypeVal, False,   1,   1, nTypeVal],
    "DBCOUNT2"        : [ocDBCount2       , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "ROUNDUP"         : [ocRoundUp        , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "ROUNDDOWN"       : [ocRoundDown      , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "RANK"            : [ocRank           , nTypeVal, False,  -1,   3, nTypeVal, nTypeRef, nTypeVal],
    "ADRESS"          : [ocAdress         , nTypeVal, False,  -1,   5, nTypeVal, nTypeVal, nTypeVal, nTypeIns, nTypeVal],
    "GETDIFFDATE360"  : [ocGetDiffDate360 , nTypeVal, False,  -1,   1, nTypeVal],
    "GETACTDATE"      : [ocGetActDate     , nTypeVal,  True,   0,   1, 0],
    "VBD"             : [ocVBD            , nTypeVal, False,  -1,   1, nTypeVal],
    "MEDIAN"          : [ocMedian         , nTypeVal, False,  -1,   1, nTypeRef],
    "SUMPRODUCT"      : [ocSumProduct     , nTypeVal, False,  -1,   1, nTypeArr],
    "SINHYP"          : [ocSinHyp         , nTypeVal, False,   1,   1, nTypeVal],
    "COSHYP"          : [ocCosHyp         , nTypeVal, False,   1,   1, nTypeVal],
    "TANHYP"          : [ocTanHyp         , nTypeVal, False,   1,   1, nTypeVal],
    "ARCSINHYP"       : [ocArcSinHyp      , nTypeVal, False,   1,   1, nTypeVal],
    "ARCCOSHYP"       : [ocArcCosHyp      , nTypeVal, False,   1,   1, nTypeVal],
    "ARCTANHYP"       : [ocArcTanHyp      , nTypeVal, False,   1,   1, nTypeVal],
    "DBGET"           : [ocDBGet          , nTypeVal, False,   3,   1, nTypeRef, nTypeRef, nTypeRef],
    "GDA2"            : [ocGDA2           , nTypeVal, False,  -1,   1, nTypeVal],
    "FREQUENCY"       : [ocFrequency      , nTypeVal, False,   2,   1, nTypeRef, nTypeRef],
    "ERRORTYPE"       : [ocErrorType      , nTypeVal, False,   1,   1, nTypeVal],
    "AVEDEV"          : [ocAveDev         , nTypeVal, False,  -1,   1, nTypeRef],
    "BETADIST"        : [ocBetaDist       , nTypeVal, False,  -1,   1, nTypeVal],
    "GAMMALN"         : [ocGammaLn        , nTypeVal, False,   1,   1, nTypeVal],
    "BETAINV"         : [ocBetaInv        , nTypeVal, False,  -1,   1, nTypeVal],
    "BINOMDIST"       : [ocBinomDist      , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "CHIDIST"         : [ocChiDist        , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "CHIINV"          : [ocChiInv         , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "KOMBIN"          : [ocKombin         , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "CONFIDENCE"      : [ocConfidence     , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "KRITBINOM"       : [ocKritBinom      , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "EVEN"            : [ocEven           , nTypeVal, False,   1,   1, nTypeVal],
    "EXPDIST"         : [ocExpDist        , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "FDIST"           : [ocFDist          , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "FINV"            : [ocFInv           , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "FISHER"          : [ocFisher         , nTypeVal, False,   1,   1, nTypeVal],
    "FISHERINV"       : [ocFisherInv      , nTypeVal, False,   1,   1, nTypeVal],
    "FLOOR"           : [ocFloor          , nTypeVal, False,   2,   1, nTypeVal, nTypeVal, nTypeIgn],
    "GAMMADIST"       : [ocGammaDist      , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "GAMMAINV"        : [ocGammaInv       , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "CEIL"            : [ocCeil           , nTypeVal, False,   2,   1, nTypeVal, nTypeVal, nTypeIgn],
    "HYPGEOMDIST"     : [ocHypGeomDist    , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "LOGNORMDIST"     : [ocLogNormDist    , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "LOGINV"          : [ocLogInv         , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "NEGBINOMVERT"    : [ocNegBinomVert   , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "NORMDIST"        : [ocNormDist       , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "STDNORMDIST"     : [ocStdNormDist    , nTypeVal, False,   1,   1, nTypeVal],
    "NORMINV"         : [ocNormInv        , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "SNORMINV"        : [ocSNormInv       , nTypeVal, False,   1,   1, nTypeVal],
    "STANDARD"        : [ocStandard       , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "ODD"             : [ocOdd            , nTypeVal, False,   1,   1, nTypeVal],
    "VARIATIONEN"     : [ocVariationen    , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "POISSONDIST"     : [ocPoissonDist    , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "TDIST"           : [ocTDist          , nTypeVal, False,   3,   1, nTypeVal, nTypeVal, nTypeVal],
    "WEIBULL"         : [ocWeibull        , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "SUMXMY2"         : [ocSumXMY2        , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "SUMX2MY2"        : [ocSumX2MY2       , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "SUMX2DY2"        : [ocSumX2DY2       , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "CHITEST"         : [ocChiTest        , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "CORREL"          : [ocCorrel         , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "COVAR"           : [ocCovar          , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "FORECAST"        : [ocForecast       , nTypeVal, False,   3,   1, nTypeVal, nTypeArr, nTypeArr],
    "FTEST"           : [ocFTest          , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "INTERCEPT"       : [ocIntercept      , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "PEARSON"         : [ocPearson        , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "RSQ"             : [ocRSQ            , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "STEYX"           : [ocSTEYX          , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "SLOPE"           : [ocSlope          , nTypeVal, False,   2,   1, nTypeArr, nTypeArr],
    "TTEST"           : [ocTTest          , nTypeVal, False,   4,   1, nTypeArr, nTypeArr, nTypeVal, nTypeVal],
    "PROB"            : [ocProb           , nTypeVal, False,  -1,   3, nTypeArr, nTypeArr, nTypeVal],
    "DEVSQ"           : [ocDevSq          , nTypeVal, False,  -1,   1, nTypeRef],
    "GEOMEAN"         : [ocGeoMean        , nTypeVal, False,  -1,   1, nTypeRef],
    "HARMEAN"         : [ocHarMean        , nTypeVal, False,  -1,   1, nTypeRef],
    "SUMSQ"           : [ocSumSQ          , nTypeVal, False,  -1,   1, nTypeRef],
    "KURT"            : [ocKurt           , nTypeVal, False,  -1,   1, nTypeRef],
    "SCHIEFE"         : [ocSchiefe        , nTypeVal, False,  -1,   1, nTypeRef],
    "ZTEST"           : [ocZTest          , nTypeVal, False,  -1,   2, nTypeRef, nTypeVal],
    "LARGE"           : [ocLarge          , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "SMALL"           : [ocSmall          , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "QUARTILE"        : [ocQuartile       , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "PERCENTILE"      : [ocPercentile     , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "PERCENTRANK"     : [ocPercentrank    , nTypeVal, False,  -1,   2, nTypeRef, nTypeVal],
    "MODALVALUE"      : [ocModalValue     , nTypeVal, False,  -1,   1, nTypeArr],
    "TRIMMEAN"        : [ocTrimMean       , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "TINV"            : [ocTInv           , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "CONCAT"          : [ocConcat         , nTypeVal, False,  -1,   1, nTypeVal],
    "POWER"           : [ocPower          , nTypeVal, False,   2,   1, nTypeVal, nTypeVal],
    "RAD"             : [ocRad            , nTypeVal, False,   1,   1, nTypeVal],
    "DEG"             : [ocDeg            , nTypeVal, False,   1,   1, nTypeVal],
    "SUBTOTAL"        : [ocSubTotal       , nTypeVal, False,  -1,   2, nTypeVal, nTypeRef],
    "SUMIF"           : [ocSumIf          , nTypeVal, False,  -1,   3, nTypeRef, nTypeVal, nTypeRef],
    "COUNTIF"         : [ocCountIf        , nTypeVal, False,   2,   1, nTypeRef, nTypeVal],
    "COUNTEMPTYCELLS" : [ocCountEmptyCells, nTypeVal, False,   1,   1, nTypeRef],
    "ROMAN"           : [ocRoman          , nTypeVal, False,  -1,   2, nTypeVal, nTypeVal],
    "EXTERNAL"        : [ocExternal       , nTypeVal, False,  -1,   2, nTypeIns, nTypeRef],
    "MACRO"           : [ocMacro          , nTypeVal, False,  -1,   1, nTypeRef],
    "ISPMT"           : [ocISPMT          , nTypeVal, False,   4,   1, nTypeVal, nTypeVal, nTypeVal, nTypeVal],
    "AVERAGEA"        : [ocAverageA       , nTypeVal, False,  -1,   1, nTypeRef],
    "MAXA"            : [ocMaxA           , nTypeVal, False,  -1,   1, nTypeRef],
    "MINA"            : [ocMinA           , nTypeVal, False,  -1,   1, nTypeRef],
    "STDEVPA"         : [ocStDevPA        , nTypeVal, False,  -1,   1, nTypeRef],
    "VARPA"           : [ocVarPA          , nTypeVal, False,  -1,   1, nTypeRef],
    "STDEVA"          : [ocStDevA         , nTypeVal, False,  -1,   1, nTypeRef],
    "VARA"            : [ocVarA           , nTypeVal, False,  -1,   1, nTypeRef]
}

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
    : str_tok:STR_CONST                     { self.rpn += "" }
    | int_tok:INT_CONST                     { self.rpn += struct.pack("<BH", ptgInt, int(int_tok.text)) }
    | num_tok:NUM_CONST                     { self.rpn += struct.pack("<Bd", ptgNum, float(num_tok.text)) }
    | rc_cell:RC_CELL                       { }
    | rc0:RC_CELL COLON rc1:RC_CELL         { }
    | name_tok:NAME                         { }
    | func_tok:NAME LP expr_list RP         { }
    | LP expr RP                            { self.rpn += struct.pack("B", ptgParen) }
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
DIGITS
    : (DIGIT)+
    ;

INT_CONST        
    : DIGITS ("." DIGITS (('e'|'E') ('+'|'-')? DIGITS)? { _ttype = NUM_CONST} )?
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
    : ('r' | 'R') ((LB (SUB)? DIGITS RB)? | DIGITS) ('c' | 'C') ((LB (SUB)? DIGITS RB)? | DIGITS)
    ;

NAME
    : LETTER (LETTER | DIGIT)*
    ;

