# -*- coding: windows-1251 -*-
"""
lots of Excel Magic Numbers
"""

# Boundaries BIFF8+

MAX_ROW = 65536
MAX_COL = 256


biff_records = {
    0x0000: "DIMENSIONS",
    0x0001: "BLANK",
    0x0002: "INTEGER",
    0x0003: "NUMBER",
    0x0004: "LABEL",
    0x0005: "BOOLERR",
    0x0006: "FORMULA",
    0x0007: "STRING",
    0x0008: "ROW",
    0x0009: "BOF",
    0x000A: "EOF",
    0x000B: "INDEX",
    0x000C: "CALCCOUNT",
    0x000D: "CALCMODE",
    0x000E: "PRECISION",
    0x000F: "REFMODE",
    0x0010: "DELTA",
    0x0011: "ITERATION",
    0x0012: "PROTECT",
    0x0013: "PASSWORD",
    0x0014: "HEADER",
    0x0015: "FOOTER",
    0x0016: "EXTERNCOUNT",
    0x0017: "EXTERNSHEET",
    0x0018: "NAME",
    0x0019: "WINDOWPROTECT",
    0x001A: "VERTICALPAGEBREAKS",
    0x001B: "HORIZONTALPAGEBREAKS",
    0x001C: "NOTE",
    0x001D: "SELECTION",
    0x001E: "FORMAT",
    0x001F: "FORMATCOUNT",
    0x0020: "COLUMNDEFAULT",
    0x0021: "ARRAY",
    0x0022: "1904",
    0x0023: "EXTERNNAME",
    0x0024: "COLWIDTH",
    0x0025: "DEFAULTROWHEIGHT",
    0x0026: "LEFTMARGIN",
    0x0027: "RIGHTMARGIN",
    0x0028: "TOPMARGIN",
    0x0029: "BOTTOMMARGIN",
    0x002A: "PRINTHEADERS",
    0x002B: "PRINTGRIDLINES",
    0x002F: "FILEPASS",
    0x0031: "FONT",
    0x0036: "TABLE",
    0x003C: "CONTINUE",
    0x003D: "WINDOW1",
    0x003E: "WINDOW2",
    0x0040: "BACKUP",
    0x0041: "PANE",
    0x0042: "CODEPAGE",
    0x0043: "XF",
    0x0044: "IXFE",
    0x0045: "EFONT",
    0x004D: "PLS",
    0x0050: "DCON",
    0x0051: "DCONREF",
    0x0053: "DCONNAME",
    0x0055: "DEFCOLWIDTH",
    0x0056: "BUILTINFMTCNT",
    0x0059: "XCT",
    0x005A: "CRN",
    0x005B: "FILESHARING",
    0x005C: "WRITEACCESS",
    0x005D: "OBJ",
    0x005E: "UNCALCED",
    0x005F: "SAFERECALC",
    0x0060: "TEMPLATE",
    0x0063: "OBJPROTECT",
    0x007D: "COLINFO",
    0x007E: "RK",
    0x007F: "IMDATA",
    0x0080: "GUTS",
    0x0081: "WSBOOL",
    0x0082: "GRIDSET",
    0x0083: "HCENTER",
    0x0084: "VCENTER",
    0x0085: "BOUNDSHEET",
    0x0086: "WRITEPROT",
    0x0087: "ADDIN",
    0x0088: "EDG",
    0x0089: "PUB",
    0x008C: "COUNTRY",
    0x008D: "HIDEOBJ",
    0x008E: "BUNDLESOFFSET",
    0x008F: "BUNDLEHEADER",
    0x0090: "SORT",
    0x0091: "SUB",
    0x0092: "PALETTE",
    0x0093: "STYLE",
    0x0094: "LHRECORD",
    0x0095: "LHNGRAPH",
    0x0096: "SOUND",
    0x0098: "LPR",
    0x0099: "STANDARDWIDTH",
    0x009A: "FNGROUPNAME",
    0x009B: "FILTERMODE",
    0x009C: "FNGROUPCOUNT",
    0x009D: "AUTOFILTERINFO",
    0x009E: "AUTOFILTER",
    0x00A0: "SCL",
    0x00A1: "SETUP",
    0x00A9: "COORDLIST",
    0x00AB: "GCW",
    0x00AE: "SCENMAN",
    0x00AF: "SCENARIO",
    0x00B0: "SXVIEW",
    0x00B1: "SXVD",
    0x00B2: "SXVI",
    0x00B4: "SXIVD",
    0x00B5: "SXLI",
    0x00B6: "SXPI",
    0x00B8: "DOCROUTE",
    0x00B9: "RECIPNAME",
    0x00BC: "SHRFMLA",
    0x00BD: "MULRK",
    0x00BE: "MULBLANK",
    0x00C1: "MMS",
    0x00C2: "ADDMENU",
    0x00C3: "DELMENU",
    0x00C5: "SXDI",
    0x00C6: "SXDB",
    0x00C7: "SXFIELD",
    0x00C8: "SXINDEXLIST",
    0x00C9: "SXDOUBLE",
    0x00CD: "SXSTRING",
    0x00CE: "SXDATETIME",
    0x00D0: "SXTBL",
    0x00D1: "SXTBRGITEM",
    0x00D2: "SXTBPG",
    0x00D3: "OBPROJ",
    0x00D5: "SXIDSTM",
    0x00D6: "RSTRING",
    0x00D7: "DBCELL",
    0x00DA: "BOOKBOOL",
    0x00DC: "SXEXT|PARAMQRY",
    0x00DD: "SCENPROTECT",
    0x00DE: "OLESIZE",
    0x00DF: "UDDESC",
    0x00E0: "XF",
    0x00E1: "INTERFACEHDR",
    0x00E2: "INTERFACEEND",
    0x00E3: "SXVS",
    0x00E5: "MERGEDCELLS",
    0x00E9: "BITMAP",
    0x00EB: "MSODRAWINGGROUP",
    0x00EC: "MSODRAWING",
    0x00ED: "MSODRAWINGSELECTION",
    0x00F0: "SXRULE",
    0x00F1: "SXEX",
    0x00F2: "SXFILT",
    0x00F6: "SXNAME",
    0x00F7: "SXSELECT",
    0x00F8: "SXPAIR",
    0x00F9: "SXFMLA",
    0x00FB: "SXFORMAT",
    0x00FC: "SST",
    0x00FD: "LABELSST",
    0x00FF: "EXTSST",
    0x0100: "SXVDEX",
    0x0103: "SXFORMULA",
    0x0122: "SXDBEX",
    0x0137: "CHTRINSERT",
    0x0138: "CHTRINFO",
    0x013B: "CHTRCELLCONTENT",
    0x013D: "TABID",
    0x0140: "CHTRMOVERANGE",
    0x014D: "CHTRINSERTTAB",
    0x015F: "LABELRANGES",
    0x0160: "USESELFS",
    0x0161: "DSF",
    0x0162: "XL5MODIFY",
    0x0196: "CHTRHEADER",
    0x01A9: "USERBVIEW",
    0x01AA: "USERSVIEWBEGIN",
    0x01AB: "USERSVIEWEND",
    0x01AD: "QSI",
    0x01AE: "SUPBOOK",
    0x01AF: "PROT4REV",
    0x01B0: "CONDFMT",
    0x01B1: "CF",
    0x01B2: "DVAL",
    0x01B5: "DCONBIN",
    0x01B6: "TXO",
    0x01B7: "REFRESHALL",
    0x01B8: "HLINK",
    0x01BA: "CODENAME",
    0x01BB: "SXFDBTYPE",
    0x01BC: "PROT4REVPASS",
    0x01BE: "DV",
    0x01C0: "XL9FILE",
    0x01C1: "RECALCID",
    0x0200: "DIMENSIONS",
    0x0201: "BLANK",
    0x0203: "NUMBER",
    0x0204: "LABEL",
    0x0205: "BOOLERR",
    0x0206: "FORMULA",
    0x0207: "STRING",
    0x0208: "ROW",
    0x0209: "BOF",
    0x020B: "INDEX",
    0x0218: "NAME",
    0x0221: "ARRAY",
    0x0223: "EXTERNNAME",
    0x0225: "DEFAULTROWHEIGHT",
    0x0231: "FONT",
    0x0236: "TABLE",
    0x023E: "WINDOW2",
    0x0243: "XF",
    0x027E: "RK",
    0x0293: "STYLE",
    0x0406: "FORMULA",
    0x0409: "BOF",
    0x041E: "FORMAT",
    0x0443: "XF",
    0x04BC: "SHRFMLA",
    0x0800: "SCREENTIP",
    0x0803: "WEBQRYSETTINGS",
    0x0804: "WEBQRYTABLES",
    0x0809: "BOF",
    0x0862: "SHEETLAYOUT",
    0x0867: "SHEETPROTECTION",
    0x1001: "UNITS",
    0x1002: "ChartChart",
    0x1003: "ChartSeries",
    0x1006: "ChartDataformat",
    0x1007: "ChartLineformat",
    0x1009: "ChartMarkerformat",
    0x100A: "ChartAreaformat",
    0x100B: "ChartPieformat",
    0x100C: "ChartAttachedlabel",
    0x100D: "ChartSeriestext",
    0x1014: "ChartChartformat",
    0x1015: "ChartLegend",
    0x1016: "ChartSerieslist",
    0x1017: "ChartBar",
    0x1018: "ChartLine",
    0x1019: "ChartPie",
    0x101A: "ChartArea",
    0x101B: "ChartScatter",
    0x101C: "ChartChartline",
    0x101D: "ChartAxis",
    0x101E: "ChartTick",
    0x101F: "ChartValuerange",
    0x1020: "ChartCatserrange",
    0x1021: "ChartAxislineformat",
    0x1022: "ChartFormatlink",
    0x1024: "ChartDefaulttext",
    0x1025: "ChartText",
    0x1026: "ChartFontx",
    0x1027: "ChartObjectLink",
    0x1032: "ChartFrame",
    0x1033: "BEGIN",
    0x1034: "END",
    0x1035: "ChartPlotarea",
    0x103A: "Chart3D",
    0x103C: "ChartPicf",
    0x103D: "ChartDropbar",
    0x103E: "ChartRadar",
    0x103F: "ChartSurface",
    0x1040: "ChartRadararea",
    0x1041: "ChartAxisparent",
    0x1043: "ChartLegendxn",
    0x1044: "ChartShtprops",
    0x1045: "ChartSertocrt",
    0x1046: "ChartAxesused",
    0x1048: "ChartSbaseref",
    0x104A: "ChartSerparent",
    0x104B: "ChartSerauxtrend",
    0x104E: "ChartIfmt",
    0x104F: "ChartPos",
    0x1050: "ChartAlruns",
    0x1051: "ChartAI",
    0x105B: "ChartSerauxerrbar",
    0x105D: "ChartSerfmt",
    0x105F: "Chart3DDataFormat",
    0x1060: "ChartFbi",
    0x1061: "ChartBoppop",
    0x1062: "ChartAxcext",
    0x1063: "ChartDat",
    0x1064: "ChartPlotgrowth",
    0x1065: "ChartSiindex",
    0x1066: "ChartGelframe",
    0x1067: "ChartBoppcustom",
    0xFFFF: ""
}


all_funcs_by_name = {
    # Includes Analysis ToolPak aka ATP aka add-in aka xcall functions,
    # distinguished by -ve opcode.
    # name: (opcode, min # args, max # args, func return type, func arg types)
    # + in func arg types means more of the same.
    'ABS'         : ( 24, 1,  1, 'V', 'V'),
    'ACCRINT'     : ( -1, 6,  7, 'V', 'VVVVVVV'),
    'ACCRINTM'    : ( -1, 3,  5, 'V', 'VVVVV'),
    'ACOS'        : ( 99, 1,  1, 'V', 'V'),
    'ACOSH'       : (233, 1,  1, 'V', 'V'),
    'ADDRESS'     : (219, 2,  5, 'V', 'VVVVV'),
    'AMORDEGRC'   : ( -1, 7,  7, 'V', 'VVVVVVV'),
    'AMORLINC'    : ( -1, 7,  7, 'V', 'VVVVVVV'),
    'AND'         : ( 36, 1, 30, 'V', 'R+'),
    'AREAS'       : ( 75, 1,  1, 'V', 'R'),
    'ASC'         : (214, 1,  1, 'V', 'V'),
    'ASIN'        : ( 98, 1,  1, 'V', 'V'),
    'ASINH'       : (232, 1,  1, 'V', 'V'),
    'ATAN'        : ( 18, 1,  1, 'V', 'V'),
    'ATAN2'       : ( 97, 2,  2, 'V', 'VV'),
    'ATANH'       : (234, 1,  1, 'V', 'V'),
    'AVEDEV'      : (269, 1, 30, 'V', 'R+'),
    'AVERAGE'     : (  5, 1, 30, 'V', 'R+'),
    'AVERAGEA'    : (361, 1, 30, 'V', 'R+'),
    'BAHTTEXT'    : (368, 1,  1, 'V', 'V'),
    'BESSELI'     : ( -1, 2,  2, 'V', 'VV'),
    'BESSELJ'     : ( -1, 2,  2, 'V', 'VV'),
    'BESSELK'     : ( -1, 2,  2, 'V', 'VV'),
    'BESSELY'     : ( -1, 2,  2, 'V', 'VV'),
    'BETADIST'    : (270, 3,  5, 'V', 'VVVVV'),
    'BETAINV'     : (272, 3,  5, 'V', 'VVVVV'),
    'BIN2DEC'     : ( -1, 1,  1, 'V', 'V'),
    'BIN2HEX'     : ( -1, 1,  2, 'V', 'VV'),
    'BIN2OCT'     : ( -1, 1,  2, 'V', 'VV'),
    'BINOMDIST'   : (273, 4,  4, 'V', 'VVVV'),
    'CEILING'     : (288, 2,  2, 'V', 'VV'),
    'CELL'        : (125, 1,  2, 'V', 'VR'),
    'CHAR'        : (111, 1,  1, 'V', 'V'),
    'CHIDIST'     : (274, 2,  2, 'V', 'VV'),
    'CHIINV'      : (275, 2,  2, 'V', 'VV'),
    'CHITEST'     : (306, 2,  2, 'V', 'AA'),
    'CHOOSE'      : (100, 2, 30, 'R', 'VR+'),
    'CLEAN'       : (162, 1,  1, 'V', 'V'),
    'CODE'        : (121, 1,  1, 'V', 'V'),
    'COLUMN'      : (  9, 0,  1, 'V', 'R'),
    'COLUMNS'     : ( 77, 1,  1, 'V', 'R'),
    'COMBIN'      : (276, 2,  2, 'V', 'VV'),
    'COMPLEX'     : ( -1, 2,  3, 'V', 'VVV'),
    'CONCATENATE' : (336, 1, 30, 'V', 'V+'),
    'CONFIDENCE'  : (277, 3,  3, 'V', 'VVV'),
    'CONVERT'     : ( -1, 3,  3, 'V', 'VVV'),
    'CORREL'      : (307, 2,  2, 'V', 'AA'),
    'COS'         : ( 16, 1,  1, 'V', 'V'),
    'COSH'        : (230, 1,  1, 'V', 'V'),
    'COUNT'       : (  0, 1, 30, 'V', 'R+'),
    'COUNTA'      : (169, 1, 30, 'V', 'R+'),
    'COUNTBLANK'  : (347, 1,  1, 'V', 'R'),
    'COUNTIF'     : (346, 2,  2, 'V', 'RV'),
    'COUPDAYBS'   : ( -1, 3,  5, 'V', 'VVVVV'),
    'COUPDAYS'    : ( -1, 3,  5, 'V', 'VVVVV'),
    'COUPDAYSNC'  : ( -1, 3,  5, 'V', 'VVVVV'),
    'COUPNCD'     : ( -1, 3,  5, 'V', 'VVVVV'),
    'COUPNUM'     : ( -1, 3,  5, 'V', 'VVVVV'),
    'COUPPCD'     : ( -1, 3,  5, 'V', 'VVVVV'),
    'COVAR'       : (308, 2,  2, 'V', 'AA'),
    'CRITBINOM'   : (278, 3,  3, 'V', 'VVV'),
    'CUMIPMT'     : ( -1, 6,  6, 'V', 'VVVVVV'),
    'CUMPRINC'    : ( -1, 6,  6, 'V', 'VVVVVV'),
    'DATE'        : ( 65, 3,  3, 'V', 'VVV'),
    'DATEDIF'     : (351, 3,  3, 'V', 'VVV'),
    'DATEVALUE'   : (140, 1,  1, 'V', 'V'),
    'DAVERAGE'    : ( 42, 3,  3, 'V', 'RRR'),
    'DAY'         : ( 67, 1,  1, 'V', 'V'),
    'DAYS360'     : (220, 2,  3, 'V', 'VVV'),
    'DB'          : (247, 4,  5, 'V', 'VVVVV'),
    'DBCS'        : (215, 1,  1, 'V', 'V'),
    'DCOUNT'      : ( 40, 3,  3, 'V', 'RRR'),
    'DCOUNTA'     : (199, 3,  3, 'V', 'RRR'),
    'DDB'         : (144, 4,  5, 'V', 'VVVVV'),
    'DEC2BIN'     : ( -1, 1,  2, 'V', 'VV'),
    'DEC2HEX'     : ( -1, 1,  2, 'V', 'VV'),
    'DEC2OCT'     : ( -1, 1,  2, 'V', 'VV'),
    'DEGREES'     : (343, 1,  1, 'V', 'V'),
    'DELTA'       : ( -1, 1,  2, 'V', 'VV'),
    'DEVSQ'       : (318, 1, 30, 'V', 'R+'),
    'DGET'        : (235, 3,  3, 'V', 'RRR'),
    'DISC'        : ( -1, 4,  5, 'V', 'VVVVV'),
    'DMAX'        : ( 44, 3,  3, 'V', 'RRR'),
    'DMIN'        : ( 43, 3,  3, 'V', 'RRR'),
    'DOLLAR'      : ( 13, 1,  2, 'V', 'VV'),
    'DOLLARDE'    : ( -1, 2,  2, 'V', 'VV'),
    'DOLLARFR'    : ( -1, 2,  2, 'V', 'VV'),
    'DPRODUCT'    : (189, 3,  3, 'V', 'RRR'),
    'DSTDEV'      : ( 45, 3,  3, 'V', 'RRR'),
    'DSTDEVP'     : (195, 3,  3, 'V', 'RRR'),
    'DSUM'        : ( 41, 3,  3, 'V', 'RRR'),
    'DURATION'    : ( -1, 5,  6, 'V', 'VVVVVV'),
    'DVAR'        : ( 47, 3,  3, 'V', 'RRR'),
    'DVARP'       : (196, 3,  3, 'V', 'RRR'),
    'EDATE'       : ( -1, 2,  2, 'V', 'VV'),
    'EFFECT'      : ( -1, 2,  2, 'V', 'VV'),
    'EOMONTH'     : ( -1, 1,  2, 'V', 'VV'),
    'ERF'         : ( -1, 1,  2, 'V', 'VV'),
    'ERFC'        : ( -1, 1,  1, 'V', 'V'),
    'ERROR.TYPE'  : (261, 1,  1, 'V', 'V'),
    'EVEN'        : (279, 1,  1, 'V', 'V'),
    'EXACT'       : (117, 2,  2, 'V', 'VV'),
    'EXP'         : ( 21, 1,  1, 'V', 'V'),
    'EXPONDIST'   : (280, 3,  3, 'V', 'VVV'),
    'FACT'        : (184, 1,  1, 'V', 'V'),
    'FACTDOUBLE'  : ( -1, 1,  1, 'V', 'V'),
    'FALSE'       : ( 35, 0,  0, 'V', '-'),
    'FDIST'       : (281, 3,  3, 'V', 'VVV'),
    'FIND'        : (124, 2,  3, 'V', 'VVV'),
    'FINDB'       : (205, 2,  3, 'V', 'VVV'),
    'FINV'        : (282, 3,  3, 'V', 'VVV'),
    'FISHER'      : (283, 1,  1, 'V', 'V'),
    'FISHERINV'   : (284, 1,  1, 'V', 'V'),
    'FIXED'       : ( 14, 2,  3, 'V', 'VVV'),
    'FLOOR'       : (285, 2,  2, 'V', 'VV'),
    'FORECAST'    : (309, 3,  3, 'V', 'VAA'),
    'FREQUENCY'   : (252, 2,  2, 'A', 'RR'),
    'FTEST'       : (310, 2,  2, 'V', 'AA'),
    'FV'          : ( 57, 3,  5, 'V', 'VVVVV'),
    'FVSCHEDULE'  : ( -1, 2,  2, 'V', 'VA'),
    'GAMMADIST'   : (286, 4,  4, 'V', 'VVVV'),
    'GAMMAINV'    : (287, 3,  3, 'V', 'VVV'),
    'GAMMALN'     : (271, 1,  1, 'V', 'V'),
    'GCD'         : ( -1, 1, 29, 'V', 'V+'),
    'GEOMEAN'     : (319, 1, 30, 'V', 'R+'),
    'GESTEP'      : ( -1, 1,  2, 'V', 'VV'),
    'GETPIVOTDATA': (358, 2, 30, 'A', 'VAV+'),
    'GROWTH'      : ( 52, 1,  4, 'A', 'RRRV'),
    'HARMEAN'     : (320, 1, 30, 'V', 'R+'),
    'HEX2BIN'     : ( -1, 1,  2, 'V', 'VV'),
    'HEX2DEC'     : ( -1, 1,  1, 'V', 'V'),
    'HEX2OCT'     : ( -1, 1,  2, 'V', 'VV'),
    'HLOOKUP'     : (101, 3,  4, 'V', 'VRRV'),
    'HOUR'        : ( 71, 1,  1, 'V', 'V'),
    'HYPERLINK'   : (359, 1,  2, 'V', 'VV'),
    'HYPGEOMDIST' : (289, 4,  4, 'V', 'VVVV'),
    'IF'          : (  1, 2,  3, 'R', 'VRR'),
    'IMABS'       : ( -1, 1,  1, 'V', 'V'),
    'IMAGINARY'   : ( -1, 1,  1, 'V', 'V'),
    'IMARGUMENT'  : ( -1, 1,  1, 'V', 'V'),
    'IMCONJUGATE' : ( -1, 1,  1, 'V', 'V'),
    'IMCOS'       : ( -1, 1,  1, 'V', 'V'),
    'IMDIV'       : ( -1, 2,  2, 'V', 'VV'),
    'IMEXP'       : ( -1, 1,  1, 'V', 'V'),
    'IMLN'        : ( -1, 1,  1, 'V', 'V'),
    'IMLOG10'     : ( -1, 1,  1, 'V', 'V'),
    'IMLOG2'      : ( -1, 1,  1, 'V', 'V'),
    'IMPOWER'     : ( -1, 2,  2, 'V', 'VV'),
    'IMPRODUCT'   : ( -1, 2,  2, 'V', 'VV'),
    'IMREAL'      : ( -1, 1,  1, 'V', 'V'),
    'IMSIN'       : ( -1, 1,  1, 'V', 'V'),
    'IMSQRT'      : ( -1, 1,  1, 'V', 'V'),
    'IMSUB'       : ( -1, 2,  2, 'V', 'VV'),
    'IMSUM'       : ( -1, 1, 29, 'V', 'V+'),
    'INDEX'       : ( 29, 2,  4, 'R', 'RVVV'),
    'INDIRECT'    : (148, 1,  2, 'R', 'VV'),
    'INFO'        : (244, 1,  1, 'V', 'V'),
    'INT'         : ( 25, 1,  1, 'V', 'V'),
    'INTERCEPT'   : (311, 2,  2, 'V', 'AA'),
    'INTRATE'     : ( -1, 4,  5, 'V', 'VVVVV'),
    'IPMT'        : (167, 4,  6, 'V', 'VVVVVV'),
    'IRR'         : ( 62, 1,  2, 'V', 'RV'),
    'ISBLANK'     : (129, 1,  1, 'V', 'V'),
    'ISERR'       : (126, 1,  1, 'V', 'V'),
    'ISERROR'     : (  3, 1,  1, 'V', 'V'),
    'ISEVEN'      : ( -1, 1,  1, 'V', 'V'),
    'ISLOGICAL'   : (198, 1,  1, 'V', 'V'),
    'ISNA'        : (  2, 1,  1, 'V', 'V'),
    'ISNONTEXT'   : (190, 1,  1, 'V', 'V'),
    'ISNUMBER'    : (128, 1,  1, 'V', 'V'),
    'ISODD'       : ( -1, 1,  1, 'V', 'V'),
    'ISPMT'       : (350, 4,  4, 'V', 'VVVV'),
    'ISREF'       : (105, 1,  1, 'V', 'R'),
    'ISTEXT'      : (127, 1,  1, 'V', 'V'),
    'KURT'        : (322, 1, 30, 'V', 'R+'),
    'LARGE'       : (325, 2,  2, 'V', 'RV'),
    'LCM'         : ( -1, 1, 29, 'V', 'V+'),
    'LEFT'        : (115, 1,  2, 'V', 'VV'),
    'LEFTB'       : (208, 1,  2, 'V', 'VV'),
    'LEN'         : ( 32, 1,  1, 'V', 'V'),
    'LENB'        : (211, 1,  1, 'V', 'V'),
    'LINEST'      : ( 49, 1,  4, 'A', 'RRVV'),
    'LN'          : ( 22, 1,  1, 'V', 'V'),
    'LOG'         : (109, 1,  2, 'V', 'VV'),
    'LOG10'       : ( 23, 1,  1, 'V', 'V'),
    'LOGEST'      : ( 51, 1,  4, 'A', 'RRVV'),
    'LOGINV'      : (291, 3,  3, 'V', 'VVV'),
    'LOGNORMDIST' : (290, 3,  3, 'V', 'VVV'),
    'LOOKUP'      : ( 28, 2,  3, 'V', 'VRR'),
    'LOWER'       : (112, 1,  1, 'V', 'V'),
    'MATCH'       : ( 64, 2,  3, 'V', 'VRR'),
    'MAX'         : (  7, 1, 30, 'V', 'R+'),
    'MAXA'        : (362, 1, 30, 'V', 'R+'),
    'MDETERM'     : (163, 1,  1, 'V', 'A'),
    'MDURATION'   : ( -1, 5,  6, 'V', 'VVVVVV'),
    'MEDIAN'      : (227, 1, 30, 'V', 'R+'),
    'MID'         : ( 31, 3,  3, 'V', 'VVV'),
    'MIDB'        : (210, 3,  3, 'V', 'VVV'),
    'MIN'         : (  6, 1, 30, 'V', 'R+'),
    'MINA'        : (363, 1, 30, 'V', 'R+'),
    'MINUTE'      : ( 72, 1,  1, 'V', 'V'),
    'MINVERSE'    : (164, 1,  1, 'A', 'A'),
    'MIRR'        : ( 61, 3,  3, 'V', 'RVV'),
    'MMULT'       : (165, 2,  2, 'A', 'AA'),
    'MOD'         : ( 39, 2,  2, 'V', 'VV'),
    'MODE'        : (330, 1, 30, 'V', 'A+'),
    'MONTH'       : ( 68, 1,  1, 'V', 'V'),
    'MROUND'      : ( -1, 2,  2, 'V', 'VV'),
    'MULTINOMIAL' : ( -1, 1, 29, 'V', 'V+'),
    'N'           : (131, 1,  1, 'V', 'R'),
    'NA'          : ( 10, 0,  0, 'V', '-'),
    'NEGBINOMDIST': (292, 3,  3, 'V', 'VVV'),
    'NETWORKDAYS' : ( -1, 2,  3, 'V', 'VVR'),
    'NOMINAL'     : ( -1, 2,  2, 'V', 'VV'),
    'NORMDIST'    : (293, 4,  4, 'V', 'VVVV'),
    'NORMINV'     : (295, 3,  3, 'V', 'VVV'),
    'NORMSDIST'   : (294, 1,  1, 'V', 'V'),
    'NORMSINV'    : (296, 1,  1, 'V', 'V'),
    'NOT'         : ( 38, 1,  1, 'V', 'V'),
    'NOW'         : ( 74, 0,  0, 'V', '-'),
    'NPER'        : ( 58, 3,  5, 'V', 'VVVVV'),
    'NPV'         : ( 11, 2, 30, 'V', 'VR+'),
    'OCT2BIN'     : ( -1, 1,  2, 'V', 'VV'),
    'OCT2DEC'     : ( -1, 1,  1, 'V', 'V'),
    'OCT2HEX'     : ( -1, 1,  2, 'V', 'VV'),
    'ODD'         : (298, 1,  1, 'V', 'V'),
    'ODDFPRICE'   : ( -1, 9,  9, 'V', 'VVVVVVVVV'),
    'ODDFYIELD'   : ( -1, 9,  9, 'V', 'VVVVVVVVV'),
    'ODDLPRICE'   : ( -1, 8,  8, 'V', 'VVVVVVVV'),
    'ODDLYIELD'   : ( -1, 8,  8, 'V', 'VVVVVVVV'),
    'OFFSET'      : ( 78, 3,  5, 'R', 'RVVVV'),
    'OR'          : ( 37, 1, 30, 'V', 'R+'),
    'PEARSON'     : (312, 2,  2, 'V', 'AA'),
    'PERCENTILE'  : (328, 2,  2, 'V', 'RV'),
    'PERCENTRANK' : (329, 2,  3, 'V', 'RVV'),
    'PERMUT'      : (299, 2,  2, 'V', 'VV'),
    'PHONETIC'    : (360, 1,  1, 'V', 'R'),
    'PI'          : ( 19, 0,  0, 'V', '-'),
    'PMT'         : ( 59, 3,  5, 'V', 'VVVVV'),
    'POISSON'     : (300, 3,  3, 'V', 'VVV'),
    'POWER'       : (337, 2,  2, 'V', 'VV'),
    'PPMT'        : (168, 4,  6, 'V', 'VVVVVV'),
    'PRICE'       : ( -1, 6,  7, 'V', 'VVVVVVV'),
    'PRICEDISC'   : ( -1, 4,  5, 'V', 'VVVVV'),
    'PRICEMAT'    : ( -1, 5,  6, 'V', 'VVVVVV'),
    'PROB'        : (317, 3,  4, 'V', 'AAVV'),
    'PRODUCT'     : (183, 1, 30, 'V', 'R+'),
    'PROPER'      : (114, 1,  1, 'V', 'V'),
    'PV'          : ( 56, 3,  5, 'V', 'VVVVV'),
    'QUARTILE'    : (327, 2,  2, 'V', 'RV'),
    'QUOTIENT'    : ( -1, 2,  2, 'V', 'VV'),
    'RADIANS'     : (342, 1,  1, 'V', 'V'),
    'RAND'        : ( 63, 0,  0, 'V', '-'),
    'RANDBETWEEN' : ( -1, 2,  2, 'V', 'VV'),
    'RANK'        : (216, 2,  3, 'V', 'VRV'),
    'RATE'        : ( 60, 3,  6, 'V', 'VVVVVV'),
    'RECEIVED'    : ( -1, 4,  5, 'V', 'VVVVV'),
    'REPLACE'     : (119, 4,  4, 'V', 'VVVV'),
    'REPLACEB'    : (207, 4,  4, 'V', 'VVVV'),
    'REPT'        : ( 30, 2,  2, 'V', 'VV'),
    'RIGHT'       : (116, 1,  2, 'V', 'VV'),
    'RIGHTB'      : (209, 1,  2, 'V', 'VV'),
    'ROMAN'       : (354, 1,  2, 'V', 'VV'),
    'ROUND'       : ( 27, 2,  2, 'V', 'VV'),
    'ROUNDDOWN'   : (213, 2,  2, 'V', 'VV'),
    'ROUNDUP'     : (212, 2,  2, 'V', 'VV'),
    'ROW'         : (  8, 0,  1, 'V', 'R'),
    'ROWS'        : ( 76, 1,  1, 'V', 'R'),
    'RSQ'         : (313, 2,  2, 'V', 'AA'),
    'RTD'         : (379, 3, 30, 'A', 'VVV+'),
    'SEARCH'      : ( 82, 2,  3, 'V', 'VVV'),
    'SEARCHB'     : (206, 2,  3, 'V', 'VVV'),
    'SECOND'      : ( 73, 1,  1, 'V', 'V'),
    'SERIESSUM'   : ( -1, 4,  4, 'V', 'VVVA'),
    'SIGN'        : ( 26, 1,  1, 'V', 'V'),
    'SIN'         : ( 15, 1,  1, 'V', 'V'),
    'SINH'        : (229, 1,  1, 'V', 'V'),
    'SKEW'        : (323, 1, 30, 'V', 'R+'),
    'SLN'         : (142, 3,  3, 'V', 'VVV'),
    'SLOPE'       : (315, 2,  2, 'V', 'AA'),
    'SMALL'       : (326, 2,  2, 'V', 'RV'),
    'SQRT'        : ( 20, 1,  1, 'V', 'V'),
    'SQRTPI'      : ( -1, 1,  1, 'V', 'V'),
    'STANDARDIZE' : (297, 3,  3, 'V', 'VVV'),
    'STDEV'       : ( 12, 1, 30, 'V', 'R+'),
    'STDEVA'      : (366, 1, 30, 'V', 'R+'),
    'STDEVP'      : (193, 1, 30, 'V', 'R+'),
    'STDEVPA'     : (364, 1, 30, 'V', 'R+'),
    'STEYX'       : (314, 2,  2, 'V', 'AA'),
    'SUBSTITUTE'  : (120, 3,  4, 'V', 'VVVV'),
    'SUBTOTAL'    : (344, 2, 30, 'V', 'VR+'),
    'SUM'         : (  4, 1, 30, 'V', 'R+'),
    'SUMIF'       : (345, 2,  3, 'V', 'RVR'),
    'SUMPRODUCT'  : (228, 1, 30, 'V', 'A+'),
    'SUMSQ'       : (321, 1, 30, 'V', 'R+'),
    'SUMX2MY2'    : (304, 2,  2, 'V', 'AA'),
    'SUMX2PY2'    : (305, 2,  2, 'V', 'AA'),
    'SUMXMY2'     : (303, 2,  2, 'V', 'AA'),
    'SYD'         : (143, 4,  4, 'V', 'VVVV'),
    'T'           : (130, 1,  1, 'V', 'R'),
    'TAN'         : ( 17, 1,  1, 'V', 'V'),
    'TANH'        : (231, 1,  1, 'V', 'V'),
    'TBILLEQ'     : ( -1, 3,  3, 'V', 'VVV'),
    'TBILLPRICE'  : ( -1, 3,  3, 'V', 'VVV'),
    'TBILLYIELD'  : ( -1, 3,  3, 'V', 'VVV'),
    'TDIST'       : (301, 3,  3, 'V', 'VVV'),
    'TEXT'        : ( 48, 2,  2, 'V', 'VV'),
    'TIME'        : ( 66, 3,  3, 'V', 'VVV'),
    'TIMEVALUE'   : (141, 1,  1, 'V', 'V'),
    'TINV'        : (332, 2,  2, 'V', 'VV'),
    'TODAY'       : (221, 0,  0, 'V', '-'),
    'TRANSPOSE'   : ( 83, 1,  1, 'A', 'A'),
    'TREND'       : ( 50, 1,  4, 'A', 'RRRV'),
    'TRIM'        : (118, 1,  1, 'V', 'V'),
    'TRIMMEAN'    : (331, 2,  2, 'V', 'RV'),
    'TRUE'        : ( 34, 0,  0, 'V', '-'),
    'TRUNC'       : (197, 1,  2, 'V', 'VV'),
    'TTEST'       : (316, 4,  4, 'V', 'AAVV'),
    'TYPE'        : ( 86, 1,  1, 'V', 'V'),
    'UPPER'       : (113, 1,  1, 'V', 'V'),
    'USDOLLAR'    : (204, 1,  2, 'V', 'VV'),
    'VALUE'       : ( 33, 1,  1, 'V', 'V'),
    'VAR'         : ( 46, 1, 30, 'V', 'R+'),
    'VARA'        : (367, 1, 30, 'V', 'R+'),
    'VARP'        : (194, 1, 30, 'V', 'R+'),
    'VARPA'       : (365, 1, 30, 'V', 'R+'),
    'VDB'         : (222, 5,  7, 'V', 'VVVVVVV'),
    'VLOOKUP'     : (102, 3,  4, 'V', 'VRRV'),
    'WEEKDAY'     : ( 70, 1,  2, 'V', 'VV'),
    'WEEKNUM'     : ( -1, 1,  2, 'V', 'VV'),
    'WEIBULL'     : (302, 4,  4, 'V', 'VVVV'),
    'WORKDAY'     : ( -1, 2,  3, 'V', 'VVR'),
    'XIRR'        : ( -1, 2,  3, 'V', 'AAV'),
    'XNPV'        : ( -1, 3,  3, 'V', 'VAA'),
    'YEAR'        : ( 69, 1,  1, 'V', 'V'),
    'YEARFRAC'    : ( -1, 2,  3, 'V', 'VVV'),
    'YIELD'       : ( -1, 6,  7, 'V', 'VVVVVVV'),
    'YIELDDISC'   : ( -1, 4,  5, 'V', 'VVVVV'),
    'YIELDMAT'    : ( -1, 5,  6, 'V', 'VVVVVV'),
    'ZTEST'       : (324, 2,  3, 'V', 'RVV'),
    }

####### This table is not used ###########
# std_func_by_num = {
#     0x000: (       "COUNT",  0, 30,   "V",         "R ...", False), # 1
#     0x001: (          "IF",  2,  3,   "R",         "V R R", False), # 2
#     0x002: (        "ISNA",  1,  1,   "V",             "V", False), # 3
#     0x003: (     "ISERROR",  1,  1,   "V",             "V", False), # 4
#     0x004: (         "SUM",  0, 30,   "V",         "R ...", False), # 5
#     0x005: (     "AVERAGE",  1, 30,   "V",         "R ...", False), # 6
#     0x006: (         "MIN",  1, 30,   "V",         "R ...", False), # 7
#     0x007: (         "MAX",  1, 30,   "V",         "R ...", False), # 8
#     0x008: (         "ROW",  0,  1,   "V",             "R", False), # 9
#     0x009: (      "COLUMN",  0,  1,   "V",             "R", False), # 10
#     0x00a: (          "NA",  0,  0,   "V",             "-", False), # 11
#     0x00b: (         "NPV",  2, 30,   "V",       "V R ...", False), # 12
#     0x00c: (       "STDEV",  1, 30,   "V",         "R ...", False), # 13
#     0x00d: (      "DOLLAR",  1,  2,   "V",           "V V", False), # 14
#     0x00e: (       "FIXED",  2,  3,   "V",         "V V V", False), # 15
#     0x00f: (         "SIN",  1,  1,   "V",             "V", False), # 16
#     0x010: (         "COS",  1,  1,   "V",             "V", False), # 17
#     0x011: (         "TAN",  1,  1,   "V",             "V", False), # 18
#     0x012: (      "ARCTAN",  1,  1,   "V",             "V", False), # 19
#     0x013: (          "PI",  0,  0,   "V",             "-", False), # 20
#     0x014: (        "SQRT",  1,  1,   "V",             "V", False), # 21
#     0x015: (         "EXP",  1,  1,   "V",             "V", False), # 22
#     0x016: (          "LN",  1,  1,   "V",             "V", False), # 23
#     0x017: (       "LOG10",  1,  1,   "V",             "V", False), # 24
#     0x018: (         "ABS",  1,  1,   "V",             "V", False), # 25
#     0x019: (         "INT",  1,  1,   "V",             "V", False), # 26
#     0x01a: (        "SIGN",  1,  1,   "V",             "V", False), # 27
#     0x01b: (       "ROUND",  2,  2,   "V",           "V V", False), # 28
#     0x01c: (      "LOOKUP",  2,  3,   "V",         "V R R", False), # 29
#     0x01d: (       "INDEX",  2,  4,   "R",       "R V V V", False), # 30
#     0x01e: (        "REPT",  2,  2,   "V",           "V V", False), # 31
#     0x01f: (         "MID",  3,  3,   "V",         "V V V", False), # 32
#     0x020: (         "LEN",  1,  1,   "V",             "V", False), # 33
#     0x021: (       "VALUE",  1,  1,   "V",             "V", False), # 34
#     0x022: (        "TRUE",  0,  0,   "V",             "-", False), # 35
#     0x023: (       "FALSE",  0,  0,   "V",             "-", False), # 36
#     0x024: (         "AND",  1, 30,   "V",         "R ...", False), # 37
#     0x025: (          "OR",  1, 30,   "V",         "R ...", False), # 38
#     0x026: (         "NOT",  1,  1,   "V",             "V", False), # 39
#     0x027: (         "MOD",  2,  2,   "V",           "V V", False), # 40
#     0x028: (      "DCOUNT",  3,  3,   "V",         "R R R", False), # 41
#     0x029: (        "DSUM",  3,  3,   "V",         "R R R", False), # 42
#     0x02a: (    "DAVERAGE",  3,  3,   "V",         "R R R", False), # 43
#     0x02b: (        "DMIN",  3,  3,   "V",         "R R R", False), # 44
#     0x02c: (        "DMAX",  3,  3,   "V",         "R R R", False), # 45
#     0x02d: (      "DSTDEV",  3,  3,   "V",         "R R R", False), # 46
#     0x02e: (         "VAR",  1, 30,   "V",         "R ...", False), # 47
#     0x02f: (        "DVAR",  3,  3,   "V",         "R R R", False), # 48
#     0x030: (        "TEXT",  2,  2,   "V",           "V V", False), # 49
#     0x031: (      "LINEST",  1,  4,   "A",       "R R V V", False), # 50
#     0x032: (       "TREND",  1,  4,   "A",       "R R R V", False), # 51
#     0x033: (      "LOGEST",  1,  4,   "A",       "R R V V", False), # 52
#     0x034: (      "GROWTH",  1,  4,   "A",       "R R R V", False), # 53
#     0x038: (          "PV",  3,  5,   "V",     "V V V V V", False), # 54
#     0x039: (          "FV",  3,  5,   "V",     "V V V V V", False), # 55
#     0x03a: (        "NPER",  3,  5,   "V",     "V V V V V", False), # 56
#     0x03b: (         "PMT",  3,  5,   "V",     "V V V V V", False), # 57
#     0x03c: (        "RATE",  3,  6,   "V",   "V V V V V V", False), # 58
#     0x03d: (        "MIRR",  3,  3,   "V",         "R V V", False), # 59
#     0x03e: (         "IRR",  1,  2,   "V",           "R V", False), # 60
#     0x03f: (        "RAND",  0,  0,   "V",             "-",  True), # 61
#     0x040: (       "MATCH",  2,  3,   "V",         "V R R", False), # 62
#     0x041: (        "DATE",  3,  3,   "V",         "V V V", False), # 63
#     0x042: (        "TIME",  3,  3,   "V",         "V V V", False), # 64
#     0x043: (         "DAY",  1,  1,   "V",             "V", False), # 65
#     0x044: (       "MONTH",  1,  1,   "V",             "V", False), # 66
#     0x045: (        "YEAR",  1,  1,   "V",             "V", False), # 67
#     0x046: (     "WEEKDAY",  1,  2,   "V",           "V V", False), # 68
#     0x047: (        "HOUR",  1,  1,   "V",             "V", False), # 69
#     0x048: (      "MINUTE",  1,  1,   "V",             "V", False), # 70
#     0x049: (      "SECOND",  1,  1,   "V",             "V", False), # 71
#     0x04a: (         "NOW",  0,  0,   "V",             "-",  True), # 72
#     0x04b: (       "AREAS",  1,  1,   "V",             "R", False), # 73
#     0x04c: (        "ROWS",  1,  1,   "V",             "R", False), # 74
#     0x04d: (     "COLUMNS",  1,  1,   "V",             "R", False), # 75
#     0x04e: (      "OFFSET",  3,  5,   "R",     "R V V V V",  True), # 76
#     0x052: (      "SEARCH",  2,  3,   "V",         "V V V", False), # 77
#     0x053: (   "TRANSPOSE",  1,  1,   "A",             "A", False), # 78
#     0x056: (        "TYPE",  1,  1,   "V",             "V", False), # 79
#     0x061: (       "ATAN2",  2,  2,   "V",           "V V", False), # 80
#     0x062: (        "ASIN",  1,  1,   "V",             "V", False), # 81
#     0x063: (        "ACOS",  1,  1,   "V",             "V", False), # 82
#     0x064: (      "CHOOSE",  2, 30,   "R",       "V R ...", False), # 83
#     0x065: (     "HLOOKUP",  3,  4,   "V",       "V R R V", False), # 84
#     0x066: (     "VLOOKUP",  3,  4,   "V",       "V R R V", False), # 85
#     0x069: (       "ISREF",  1,  1,   "V",             "R", False), # 86
#     0x06d: (         "LOG",  1,  2,   "V",           "V V", False), # 87
#     0x06f: (        "CHAR",  1,  1,   "V",             "V", False), # 88
#     0x070: (       "LOWER",  1,  1,   "V",             "V", False), # 89
#     0x071: (       "UPPER",  1,  1,   "V",             "V", False), # 90
#     0x072: (      "PROPER",  1,  1,   "V",             "V", False), # 91
#     0x073: (        "LEFT",  1,  2,   "V",           "V V", False), # 92
#     0x074: (       "RIGHT",  1,  2,   "V",           "V V", False), # 93
#     0x075: (       "EXACT",  2,  2,   "V",           "V V", False), # 94
#     0x076: (        "TRIM",  1,  1,   "V",             "V", False), # 95
#     0x077: (     "REPLACE",  4,  4,   "V",       "V V V V", False), # 96
#     0x078: (  "SUBSTITUTE",  3,  4,   "V",       "V V V V", False), # 97
#     0x079: (        "CODE",  1,  1,   "V",             "V", False), # 98
#     0x07c: (        "FIND",  2,  3,   "V",         "V V V", False), # 99
#     0x07d: (        "CELL",  1,  2,   "V",           "V R",  True), # 100
#     0x07e: (       "ISERR",  1,  1,   "V",             "V", False), # 101
#     0x07f: (      "ISTEXT",  1,  1,   "V",             "V", False), # 102
#     0x080: (    "ISNUMBER",  1,  1,   "V",             "V", False), # 103
#     0x081: (     "ISBLANK",  1,  1,   "V",             "V", False), # 104
#     0x082: (           "T",  1,  1,   "V",             "R", False), # 105
#     0x083: (           "N",  1,  1,   "V",             "R", False), # 106
#     0x08c: (   "DATEVALUE",  1,  1,   "V",             "V", False), # 107
#     0x08d: (   "TIMEVALUE",  1,  1,   "V",             "V", False), # 108
#     0x08e: (         "SLN",  3,  3,   "V",         "V V V", False), # 109
#     0x08f: (         "SYD",  4,  4,   "V",       "V V V V", False), # 110
#     0x090: (         "DDB",  4,  5,   "V",     "V V V V V", False), # 111
#     0x094: (    "INDIRECT",  1,  2,   "R",           "V V",  True), # 112
#     0x0a2: (       "CLEAN",  1,  1,   "V",             "V", False), # 113
#     0x0a3: (     "MDETERM",  1,  1,   "V",             "A", False), # 114
#     0x0a4: (    "MINVERSE",  1,  1,   "A",             "A", False), # 115
#     0x0a5: (       "MMULT",  2,  2,   "A",           "A A", False), # 116
#     0x0a7: (        "IPMT",  4,  6,   "V",   "V V V V V V", False), # 117
#     0x0a8: (        "PPMT",  4,  6,   "V",   "V V V V V V", False), # 118
#     0x0a9: (      "COUNTA",  0, 30,   "V",         "R ...", False), # 119
#     0x0b7: (     "PRODUCT",  0, 30,   "V",         "R ...", False), # 120
#     0x0b8: (        "FACT",  1,  1,   "V",             "V", False), # 121
#     0x0bf: (    "DPRODUCT",  3,  3,   "V",         "R R R", False), # 122
#     0x0c0: (   "ISNONTEXT",  1,  1,   "V",             "V", False), # 123
#     0x0c1: (      "STDEVP",  1, 30,   "V",         "R ...", False), # 124
#     0x0c2: (        "VARP",  1, 30,   "V",         "R ...", False), # 125
#     0x0c3: (     "DSTDEVP",  3,  3,   "V",         "R R R", False), # 126
#     0x0c4: (       "DVARP",  3,  3,   "V",         "R R R", False), # 127
#     0x0c5: (       "TRUNC",  1,  2,   "V",           "V V", False), # 128
#     0x0c6: (   "ISLOGICAL",  1,  1,   "V",             "V", False), # 129
#     0x0c7: (     "DCOUNTA",  3,  3,   "V",         "R R R", False), # 130
#     0x0cc: (    "USDOLLAR",  1,  2,   "V",           "V V", False), # 131
#     0x0cd: (       "FINDB",  2,  3,   "V",         "V V V", False), # 132
#     0x0ce: (     "SEARCHB",  2,  3,   "V",         "V V V", False), # 133
#     0x0cf: (    "REPLACEB",  4,  4,   "V",       "V V V V", False), # 134
#     0x0d0: (       "LEFTB",  1,  2,   "V",           "V V", False), # 135
#     0x0d1: (      "RIGHTB",  1,  2,   "V",           "V V", False), # 136
#     0x0d2: (        "MIDB",  3,  3,   "V",         "V V V", False), # 137
#     0x0d3: (        "LENB",  1,  1,   "V",             "V", False), # 138
#     0x0d4: (     "ROUNDUP",  2,  2,   "V",           "V V", False), # 139
#     0x0d5: (   "ROUNDDOWN",  2,  2,   "V",           "V V", False), # 140
#     0x0d6: (         "ASC",  1,  1,   "V",             "V", False), # 141
#     0x0d7: (        "DBSC",  1,  1,   "V",             "V", False), # 142
#     0x0d8: (        "RANK",  2,  3,   "V",         "V R V", False), # 143
#     0x0db: (     "ADDRESS",  2,  5,   "V",     "V V V V V", False), # 144
#     0x0dc: (     "DAYS360",  2,  3,   "V",         "V V V", False), # 145
#     0x0dd: (       "TODAY",  0,  0,   "V",             "-",  True), # 146
#     0x0de: (         "VDB",  5,  7,   "V", "V V V V V V V", False), # 147
#     0x0e3: (      "MEDIAN",  1, 30,   "V",         "R ...", False), # 148
#     0x0e4: (  "SUMPRODUCT",  1, 30,   "V",         "A ...", False), # 149
#     0x0e5: (        "SINH",  1,  1,   "V",             "V", False), # 150
#     0x0e6: (        "COSH",  1,  1,   "V",             "V", False), # 151
#     0x0e7: (        "TANH",  1,  1,   "V",             "V", False), # 152
#     0x0e8: (       "ASINH",  1,  1,   "V",             "V", False), # 153
#     0x0e9: (       "ACOSH",  1,  1,   "V",             "V", False), # 154
#     0x0ea: (       "ATANH",  1,  1,   "V",             "V", False), # 155
#     0x0eb: (        "DGET",  3,  3,   "V",         "R R R", False), # 156
#     0x0f4: (        "INFO",  1,  1,   "V",             "V", False), # 157
#     0x0f7: (          "DB",  4,  5,   "V",     "V V V V V", False), # 158
#     0x0fc: (   "FREQUENCY",  2,  2,   "A",           "R R", False), # 159
#     0x105: (  "ERROR.TYPE",  1,  1,   "V",             "V", False), # 160
#     0x10d: (      "AVEDEV",  1, 30,   "V",         "R ...", False), # 161
#     0x10e: (    "BETADIST",  3,  5,   "V",     "V V V V V", False), # 162
#     0x10f: (     "GAMMALN",  1,  1,   "V",             "V", False), # 163
#     0x110: (     "BETAINV",  3,  5,   "V",     "V V V V V", False), # 164
#     0x111: (   "BINOMDIST",  4,  4,   "V",       "V V V V", False), # 165
#     0x112: (     "CHIDIST",  2,  2,   "V",           "V V", False), # 166
#     0x113: (      "CHIINV",  2,  2,   "V",           "V V", False), # 167
#     0x114: (      "COMBIN",  2,  2,   "V",           "V V", False), # 168
#     0x115: (  "CONFIDENCE",  3,  3,   "V",         "V V V", False), # 169
#     0x116: (   "CRITBINOM",  3,  3,   "V",         "V V V", False), # 170
#     0x117: (        "EVEN",  1,  1,   "V",             "V", False), # 171
#     0x118: (   "EXPONDIST",  3,  3,   "V",         "V V V", False), # 172
#     0x119: (       "FDIST",  3,  3,   "V",         "V V V", False), # 173
#     0x11a: (        "FINV",  3,  3,   "V",         "V V V", False), # 174
#     0x11b: (      "FISHER",  1,  1,   "V",             "V", False), # 175
#     0x11c: (   "FISHERINV",  1,  1,   "V",             "V", False), # 176
#     0x11d: (       "FLOOR",  2,  2,   "V",           "V V", False), # 177
#     0x11e: (   "GAMMADIST",  4,  4,   "V",       "V V V V", False), # 178
#     0x11f: (    "GAMMAINV",  3,  3,   "V",         "V V V", False), # 179
#     0x120: (     "CEILING",  2,  2,   "V",           "V V", False), # 180
#     0x121: ( "HYPGEOMVERT",  4,  4,   "V",       "V V V V", False), # 181
#     0x122: ( "LOGNORMDIST",  3,  3,   "V",         "V V V", False), # 182
#     0x123: (      "LOGINV",  3,  3,   "V",         "V V V", False), # 183
#     0x124: ("NEGBINOMDIST",  3,  3,   "V",         "V V V", False), # 184
#     0x125: (    "NORMDIST",  4,  4,   "V",       "V V V V", False), # 185
#     0x126: (   "NORMSDIST",  1,  1,   "V",             "V", False), # 186
#     0x127: (     "NORMINV",  3,  3,   "V",         "V V V", False), # 187
#     0x128: (   "MNORMSINV",  1,  1,   "V",             "V", False), # 188
#     0x129: ( "STANDARDIZE",  3,  3,   "V",         "V V V", False), # 189
#     0x12a: (         "ODD",  1,  1,   "V",             "V", False), # 190
#     0x12b: (      "PERMUT",  2,  2,   "V",           "V V", False), # 191
#     0x12c: (     "POISSON",  3,  3,   "V",         "V V V", False), # 192
#     0x12d: (       "TDIST",  3,  3,   "V",         "V V V", False), # 193
#     0x12e: (     "WEIBULL",  4,  4,   "V",       "V V V V", False), # 194
#     0x12f: (     "SUMXMY2",  2,  2,   "V",           "A A", False), # 195
#     0x130: (    "SUMX2MY2",  2,  2,   "V",           "A A", False), # 196
#     0x131: (    "SUMX2PY2",  2,  2,   "V",           "A A", False), # 197
#     0x132: (     "CHITEST",  2,  2,   "V",           "A A", False), # 198
#     0x133: (      "CORREL",  2,  2,   "V",           "A A", False), # 199
#     0x134: (       "COVAR",  2,  2,   "V",           "A A", False), # 200
#     0x135: (    "FORECAST",  3,  3,   "V",         "V A A", False), # 201
#     0x136: (       "FTEST",  2,  2,   "V",           "A A", False), # 202
#     0x137: (   "INTERCEPT",  2,  2,   "V",           "A A", False), # 203
#     0x138: (     "PEARSON",  2,  2,   "V",           "A A", False), # 204
#     0x139: (         "RSQ",  2,  2,   "V",           "A A", False), # 205
#     0x13a: (       "STEYX",  2,  2,   "V",           "A A", False), # 206
#     0x13b: (       "SLOPE",  2,  2,   "V",           "A A", False), # 207
#     0x13c: (       "TTEST",  4,  4,   "V",       "A A V V", False), # 208
#     0x13d: (        "PROB",  3,  4,   "V",       "A A V V", False), # 209
#     0x13e: (       "DEVSQ",  1, 30,   "V",         "R ...", False), # 210
#     0x13f: (     "GEOMEAN",  1, 30,   "V",         "R ...", False), # 211
#     0x140: (     "HARMEAN",  1, 30,   "V",         "R ...", False), # 212
#     0x141: (       "SUMSQ",  0, 30,   "V",         "R ...", False), # 213
#     0x142: (        "KURT",  1, 30,   "V",         "R ...", False), # 214
#     0x143: (        "SKEW",  1, 30,   "V",         "R ...", False), # 215
#     0x144: (       "ZTEST",  2,  3,   "V",         "R V V", False), # 216
#     0x145: (       "LARGE",  2,  2,   "V",           "R V", False), # 217
#     0x146: (       "SMALL",  2,  2,   "V",           "R V", False), # 218
#     0x147: (    "QUARTILE",  2,  2,   "V",           "R V", False), # 219
#     0x148: (  "PERCENTILE",  2,  2,   "V",           "R V", False), # 220
#     0x149: ( "PERCENTRANK",  2,  3,   "V",         "R V V", False), # 221
#     0x14a: (        "MODE",  1, 30,   "V",         "A ...", False), # 222
#     0x14b: (    "TRIMMEAN",  2,  2,   "V",           "R V", False), # 223
#     0x14c: (        "TINV",  2,  2,   "V",           "V V", False), # 224
#     0x150: ( "CONCATENATE",  0, 30,   "V",         "V ...", False), # 225
#     0x151: (       "POWER",  2,  2,   "V",           "V V", False), # 226
#     0x156: (     "RADIANS",  1,  1,   "V",             "V", False), # 227
#     0x157: (     "DEGREES",  1,  1,   "V",             "V", False), # 228
#     0x158: (    "SUBTOTAL",  2, 30,   "V",       "V R ...", False), # 229
#     0x159: (       "SUMIF",  2,  3,   "V",         "R V R", False), # 230
#     0x15a: (     "COUNTIF",  2,  2,   "V",           "R V", False), # 231
#     0x15b: (  "COUNTBLANK",  1,  1,   "V",             "R", False), # 232
#     0x15e: (       "ISPMT",  4,  4,   "V",       "V V V V", False), # 233
#     0x15f: (     "DATEDIF",  3,  3,   "V",         "V V V", False), # 234
#     0x160: (  "DATESTRING",  1,  1,   "V",             "V", False), # 235
#     0x161: ("NUMBERSTRING",  2,  2,   "V",           "V V", False), # 236
#     0x162: (       "ROMAN",  1,  2,   "V",           "V V", False), # 237
#     0x166: ("GETPIVOTDATA",  2, 30,   "A",             "-", False), # 238
#     0x167: (   "HYPERLINK",  1,  2,   "V",           "V V", False), # 239
#     0x168: (    "PHONETIC",  1,  1,   "V",             "R", False), # 240
#     0x169: (    "AVERAGEA",  1, 30,   "V",         "R ...", False), # 241
#     0x16a: (        "MAXA",  1, 30,   "V",         "R ...", False), # 242
#     0x16b: (        "MINA",  1, 30,   "V",         "R ...", False), # 243
#     0x16c: (     "STDEVPA",  1, 30,   "V",         "R ...", False), # 244
#     0x16d: (       "VARPA",  1, 30,   "V",         "R ...", False), # 245
#     0x16e: (      "STDEVA",  1, 30,   "V",         "R ...", False), # 246
#     0x16f: (        "VARA",  1, 30,   "V",         "R ...", False)  # 247
# }


# Formulas Parse things

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

ptgArrayR       = 0x20
ptgFuncR        = 0x21
ptgFuncVarR     = 0x22
ptgNameR        = 0x23
ptgRefR         = 0x24
ptgAreaR        = 0x25
ptgMemAreaR     = 0x26
ptgMemErrR      = 0x27
ptgMemNoMemR    = 0x28
ptgMemFuncR     = 0x29
ptgRefErrR      = 0x2a
ptgAreaErrR     = 0x2b
ptgRefNR        = 0x2c
ptgAreaNR       = 0x2d
ptgMemAreaNR    = 0x2e
ptgMemNoMemNR   = 0x2f
ptgNameXR       = 0x39
ptgRef3dR       = 0x3a
ptgArea3dR      = 0x3b
ptgRefErr3dR    = 0x3c
ptgAreaErr3dR   = 0x3d

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
    ptgArrayR      : "ptgArrayR",
    ptgFuncR       : "ptgFuncR",
    ptgFuncVarR    : "ptgFuncVarR",
    ptgNameR       : "ptgNameR",
    ptgRefR        : "ptgRefR",
    ptgAreaR       : "ptgAreaR",
    ptgMemAreaR    : "ptgMemAreaR",
    ptgMemErrR     : "ptgMemErrR",
    ptgMemNoMemR   : "ptgMemNoMemR",
    ptgMemFuncR    : "ptgMemFuncR",
    ptgRefErrR     : "ptgRefErrR",
    ptgAreaErrR    : "ptgAreaErrR",
    ptgRefNR       : "ptgRefNR",
    ptgAreaNR      : "ptgAreaNR",
    ptgMemAreaNR   : "ptgMemAreaNR",
    ptgMemNoMemNR  : "ptgMemNoMemNR",
    ptgNameXR      : "ptgNameXR",
    ptgRef3dR      : "ptgRef3dR",
    ptgArea3dR     : "ptgArea3dR",
    ptgRefErr3dR   : "ptgRefErr3dR",
    ptgAreaErr3dR  : "ptgAreaErr3dR",
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


error_msg_by_code = {
    0x00: u"#NULL!",  # intersection of two cell ranges is empty
    0x07: u"#DIV/0!", # division by zero
    0x0F: u"#VALUE!", # wrong type of operand
    0x17: u"#REF!",   # illegal or deleted cell reference
    0x1D: u"#NAME?",  # wrong function or range name
    0x24: u"#NUM!",   # value range overflow
    0x2A: u"#N/A!"    # argument or function not available
}
