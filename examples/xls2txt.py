#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman

__rev_id__ = """$Id$"""


from pyExcelerator import *
import sys

me, args = sys.argv[0], sys.argv[1:]

if args:
    for arg in args:
        for sheet_name, values in parse_xls(arg):
            print 'Sheet = "%s"' % sheet_name.encode('cp866', 'backslashreplace')
            print '----------------'
            for row_idx, col_idx in sorted(values.keys()):
                v = values[(row_idx, col_idx)]
                if isinstance(v, unicode):
                    v = v.encode('cp866', 'backslashreplace')
                print '(%d, %d) =' % (row_idx, col_idx), v
            print '----------------'
else:
    print 'usage: %s (inputfile)+' % me

