#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman

from xlwt import *

w = Workbook()
ws1 = w.add_sheet('cp1251')

UnicodeUtils.DEFAULT_ENCODING = 'cp1251'
ws1.write(0, 0, 'Оля')

w.save('unicode0.xls')

