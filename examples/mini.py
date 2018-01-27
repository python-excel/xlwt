#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Copyright (C) 2005 Kiseliov Roman

from xlwt import *

w = Workbook()
ws = w.add_sheet('xlwt was here')
w.save('mini.xls')
