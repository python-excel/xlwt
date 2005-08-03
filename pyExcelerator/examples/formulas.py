#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman
__rev_id__ = """$Id$"""


from pyExcelerator import *

w = Workbook()
ws = w.add_sheet('F')

ws.write(0, 0, Formula("-(1+1)"))
ws.write(1, 0, Formula("-(1+1)/(-2-2)"))
ws.write(2, 0, Formula("-(134.8780789+1)"))
ws.write(3, 0, Formula("-(134.8780789e-10+1)"))
ws.write(4, 0, Formula("-1/(1+1)+9344"))

w.save('formulas.xls')
