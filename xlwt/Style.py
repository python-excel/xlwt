# -*- coding: windows-1252 -*-

# 2007-10-06 SJM Optional compression of duplicate fonts and XFs
# 2007-02-21 SJM Remove debris from earlier fix
# 2007-02-21 SJM "The first user-defined format starts at 164" (not 163)
# 2007-02-21 SJM Make it work with Python 2.3
# 2007-01-10 SJM Fix XFStyle() always returning the same object.

import Formatting
from BIFFRecords import *

FIRST_USER_DEFINED_NUM_FORMAT_IDX = 164

class XFStyle(object):
    
    def __init__(self):
        self.num_format_str  = 'General'
        self.font            = Formatting.Font() 
        self.alignment       = Formatting.Alignment()
        self.borders         = Formatting.Borders()
        self.pattern         = Formatting.Pattern()
        self.protection      = Formatting.Protection()
            
default_style = XFStyle()            
            
class StyleCollection(object):
    _std_num_fmt_list = [
            'general',
            '0',
            '0.00',
            '#,##0',
            '#,##0.00',
            '"$"#,##0_);("$"#,##',
            '"$"#,##0_);[Red]("$"#,##',
            '"$"#,##0.00_);("$"#,##',
            '"$"#,##0.00_);[Red]("$"#,##',
            '0%',
            '0.00%',
            '0.00E+00',
            '# ?/?',
            '# ??/??',
            'M/D/YY',
            'D-MMM-YY',
            'D-MMM',
            'MMM-YY',
            'h:mm AM/PM',
            'h:mm:ss AM/PM',
            'h:mm',
            'h:mm:ss',
            'M/D/YY h:mm',
            '_(#,##0_);(#,##0)',
            '_(#,##0_);[Red](#,##0)',
            '_(#,##0.00_);(#,##0.00)',
            '_(#,##0.00_);[Red](#,##0.00)',
            '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
            '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
            'mm:ss',
            '[h]:mm:ss',
            'mm:ss.0',
            '##0.0E+0',
            '@'   
    ]

    def __init__(self, style_compression=0):
        self.style_compression = style_compression
        self.stats = [0, 0, 0, 0, 0, 0]
        self._font_id2x = {}
        self._font_x2id = {}
        self._font_val2x = {}
        
        for x in (0, 1, 2, 3, 5): # The font with index 4 is omitted in all BIFF versions
            font = Formatting.Font()
            search_key = font._search_key()
            self._font_id2x[font] = x
            self._font_x2id[x] = font
            self._font_val2x[search_key] = x
        
        self._xf_id2x = {}
        self._xf_x2id = {}
        self._xf_val2x = {}
        
        self._num_formats = {}
        for fmtidx, fmtstr in zip(range(0, 23), StyleCollection._std_num_fmt_list[0:23]):
            self._num_formats[fmtstr] = fmtidx 
        for fmtidx, fmtstr in zip(range(37, 50), StyleCollection._std_num_fmt_list[23:]):
            self._num_formats[fmtstr] = fmtidx 

        self.default_style = XFStyle()
        self._default_xf = self._add_style(self.default_style)[0]
        
    def add(self, style): 
        if style == None:
            return 0x10
        return self._add_style(style)[1]
    
    def _add_style(self, style):
        num_format_str = style.num_format_str
        if num_format_str in self._num_formats:
            num_format_idx = self._num_formats[num_format_str]
        else:
            num_format_idx = (
                FIRST_USER_DEFINED_NUM_FORMAT_IDX
                + len(self._num_formats)
                - len(StyleCollection._std_num_fmt_list)
                )
            self._num_formats[num_format_str] = num_format_idx
            
        font = style.font
        if font in self._font_id2x:
            font_idx = self._font_id2x[font]
            self.stats[0] += 1
        elif self.style_compression:
            search_key = font._search_key()
            font_idx = self._font_val2x.get(search_key)
            if font_idx is not None:
                self._font_id2x[font] = font_idx
                self.stats[1] += 1
            else:
                font_idx = len(self._font_x2id) + 1 # Why plus 1? Font 4 is missing
                self._font_id2x[font] = font_idx
                self._font_val2x[search_key] = font_idx
                self._font_x2id[font_idx] = font
                self.stats[2] += 1
        else:
            font_idx = len(self._font_id2x) + 1
            self._font_id2x[font] = font_idx
            self.stats[2] += 1
        
        gof = (style.alignment, style.borders, style.pattern, style.protection)
        xf = (font_idx, num_format_idx) + gof
        if xf in self._xf_id2x:
            xf_index = self._xf_id2x[xf]
            self.stats[3] += 1
        elif self.style_compression == 2:
            xf_key = (font_idx, num_format_idx) + tuple([obj._search_key() for obj in gof])
            xf_index = self._xf_val2x.get(xf_key)
            if xf_index is not None:
                self._xf_id2x[xf] = xf_index
                self.stats[4] += 1
            else:
                xf_index = 0x10 + len(self._xf_x2id)
                self._xf_id2x[xf] = xf_index
                self._xf_val2x[xf_key] = xf_index
                self._xf_x2id[xf_index] = xf
                self.stats[5] += 1
        else:
            xf_index = 0x10 + len(self._xf_id2x)
            self._xf_id2x[xf] = xf_index
            self.stats[5] += 1

        if xf_index >= 0xFFF: 
            # 12 bits allowed, 0xFFF is a sentinel value
            raise ValueError("More than 4094 XFs (styles)")
            
        return xf, xf_index
        
    def get_biff_data(self):
        result = ''
        result += self._all_fonts()
        result += self._all_num_formats()
        result += self._all_cell_styles()
        result += self._all_styles()
        return result 
            
    def _all_fonts(self):
        result = ''
        if self.style_compression:
            alist = self._font_x2id.items()
        else: 
            alist = [(x, o) for o, x in self._font_id2x.items()]
        alist.sort()
        for font_idx, font in alist:
            result += font.get_biff_record().get()
        return result
    
    def _all_num_formats(self):
        result = ''
        alist = [
            (v, k)
            for k, v in self._num_formats.items()
            if v >= FIRST_USER_DEFINED_NUM_FORMAT_IDX
            ]
        alist.sort()
        for fmtidx, fmtstr in alist:
            result += NumberFormatRecord(fmtidx, fmtstr).get()
        return result
    
    def _all_cell_styles(self):        
        result = ''
        for i in range(0, 16):
            result += XFRecord(self._default_xf, 'style').get()
        if self.style_compression == 2:            
            alist = self._xf_x2id.items()
        else:
            alist = [(x, o) for o, x in self._xf_id2x.items()]
        alist.sort()
        for xf_idx, xf in alist:
            result += XFRecord(xf).get()
        return result
        
    def _all_styles(self):
        return StyleRecord().get()
        
if __name__ == '__main__':
    sc = StyleCollection()
    f = file('styles.bin', 'wb')
    f.write(sc.get_biff_data())
    f.close()

            
