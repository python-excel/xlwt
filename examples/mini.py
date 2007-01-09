#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman
__rev_id__ = """$Id$"""


from xlwt import *

w = Workbook()
ws = w.add_sheet('xlwt was here')
w.save('mini.xls')
