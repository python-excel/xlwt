#!/usr/bin/env python
# -*- coding: windows-1251 -*-

#  Copyright (C) 2005 Roman V. Kiseliov
#  All rights reserved.
# 
#  Redistribution and use in source and binary forms, with or without
#  modification, are permitted provided that the following conditions
#  are met:
# 
#  1. Redistributions of source code must retain the above copyright
#     notice, this list of conditions and the following disclaimer.
# 
#  2. Redistributions in binary form must reproduce the above copyright
#     notice, this list of conditions and the following disclaimer in
#     the documentation and/or other materials provided with the
#     distribution.
# 
#  3. All advertising materials mentioning features or use of this
#     software must display the following acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <unicorn@kurskline.ru>."
# 
#  4. Redistributions of any form whatsoever must retain the following
#     acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <unicorn@kurskline.ru>."
# 
#  THIS SOFTWARE IS PROVIDED BY Roman V. Kiseliov ``AS IS'' AND ANY
#  EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
#  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
#  PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL Roman V. Kiseliov OR
#  ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
#  SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
#  NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
#  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
#  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
#  STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
#  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
#  OF THE POSSIBILITY OF SUCH DAMAGE.


__rev_id__ = """$Id$"""

import BIFFRecords
from Deco import *
from Worksheet import Worksheet
import Style

class Row(object):
    #################################################################
    ## Constructor
    #################################################################
    def __init__(self, index, parent_sheet):
        self._index = index
        self._parent = parent_sheet
        self._parent_wb = parent_sheet.get_parent()
        self._cells = {}
        self._mul_blanks = []
        self._total_str = 0
        self._xf_index = 0x0F
        self._has_default_format = 0
        self._height_in_pixels = 0x11
        
        self.height = 0x00FF
        self.has_default_height = 0x00
        self.level = 0
        self.collapse = 0
        self.hidden = 0
        self.space_above = 0
        self.space_below = 0

    def height_in_pixels(self):
        return self._height_in_pixels

    @accepts(object, Style.XFStyle)
    def set_style(self, style):
        twips = style.font.height
        points = float(twips)/20.0
        # Cell height in pixels can be calcuted by following approx. formula:
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi 
        pix = int(round(points*83.0/50.0 + 2.0/5.0))
        if pix > self._height_in_pixels:
            self._height_in_pixels = pix
        self._xf_index = self._parent_wb.add_style(style)
            
    def get_xf_index(self):
        return self._xf_index
    
    def get_cells_count(self):
        return len(self._cells)
    
    def get_min_col(self):
        result = 0
        if len(self._cells) > 0:
            result = min(self._cells)
        return result
        
    def get_max_col(self):
        result = 0
        if len(self._cells) > 0:
            result = max(self._cells)
        return result
        
    def get_str_count(self):
        return self._total_str

    @accepts(object, int, str, Style.XFStyle)
    def write(self, col, label, style):
        twips = style.font.height
        points = float(twips)/20.0
        # Cell height in pixels can be calcuted by following approx. formula:
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi 
        pix = int(round(points*83.0/50.0 + 2.0/5.0))
        if pix > self._height_in_pixels:
            self._height_in_pixels = pix
        self._cells[col] = (self._parent_wb.add_style(style), self._parent_wb.add_str(label))
        self._total_str += 1

    @accepts(object, int, int, Style.XFStyle)                        
    def write_blanks(self, c1, c2, style):
        twips = style.font.height
        points = float(twips)/20.0
        # Cell height in pixels can be calcuted by following approx. formula:
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi 
        pix = int(round(points*83.0/50.0 + 2.0/5.0))
        if pix > self._height_in_pixels:
            self._height_in_pixels = pix
        style_idx = self._parent_wb.add_style(style)
        for col in range(c1, c2):
            self._cells[col] = (style_idx, -2)
        self._mul_blanks.append((c1, c2, style_idx))
        
    def get_row_biff_data(self):
        height_options = (self.height & 0x07FFF) 
        height_options |= (self.has_default_height & 0x01) << 15

        options =  (self.level & 0x07) << 0
        options |= (self.collapse & 0x01) << 4
        options |= (self.hidden & 0x01) << 5
        options |= (0x00 & 0x01) << 6
        options |= (0x01 & 0x01) << 8
        if self._xf_index != 0x0F:
            options |= (0x01 & 0x01) << 7
        else:
            options |= (0x00 & 0x01) << 7
        options |= (self._xf_index & 0x0FFF) << 16 
        options |= (0x00 & self.space_above) << 28
        options |= (0x00 & self.space_below) << 29
        
        col_nums = self._cells.keys()
        if len(col_nums) > 0:
            min_col = min(col_nums)
            max_col = max(col_nums)
        else:
            min_col = 0
            max_col = 0
        return BIFFRecords.RowRecord(self._index, min_col, max_col, height_options, options).get()                                              
                        
    def get_cells_biff_data(self):
        result = ''
        for col_idx in self._cells:
            xf_idx, sst_idx = self._cells[col_idx]
            if self._xf_index != 0x0F:
                xf_idx = self._xf_index
            if sst_idx != -2:
                 result += BIFFRecords.LabelSSTRecord(self._index, col_idx, xf_idx, sst_idx).get()        
                
        for col1, col2, xf_idx in self._mul_blanks:
            if self._xf_index != 0x0F:
                xf_idx = self._xf_index
            result += BIFFRecords.MulBlankRecord(self._index, col1, col2, xf_idx).get()
            
        return result
        