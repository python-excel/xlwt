__author__ = 'hrwl'

from datetime import datetime
import hashlib
import os
from six import BytesIO
import unittest

from xlwt import *


class FullTestCase(unittest.TestCase):
    def setUp(self):
        self.python_bmp = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'python.bmp')

    def test_blanks(self):
        font0 = Font()
        font0.name = 'Times New Roman'
        font0.struck_out = True
        font0.bold = True

        style0 = XFStyle()
        style0.font = font0

        wb = Workbook()
        ws0 = wb.add_sheet('0')

        ws0.write(1, 1, 'Test', style0)

        for i in range(0, 0x53):
            borders = Borders()
            borders.left = i
            borders.right = i
            borders.top = i
            borders.bottom = i

            style = XFStyle()
            style.borders = borders

            ws0.write(i, 2, '', style)
            ws0.write(i, 3, hex(i), style0)

        ws0.write_merge(5, 8, 6, 10, "")

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('353fa63cf6f9f4b1369ac589564d7f5b', md5.hexdigest())

    def test_col_width(self):
        wb = Workbook()
        ws = wb.add_sheet('Hey, Dude')

        for i in range(6, 80):
            fnt = Font()
            fnt.height = i*20
            style = XFStyle()
            style.font = fnt
            ws.write(1, i, 'Test')
            ws.col(i).width = 0x0d00 + i

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('2f60c22c584d8c078427cbaafc9dd23d', md5.hexdigest())

    def test_country(self):
        wb = Workbook()

        wb.country_code = 61
        ws = wb.add_sheet('AU')

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('25322fbea1f258f93137f8b6f72cfdaa', md5.hexdigest())

    def test_dates(self):
        w = Workbook()
        ws = w.add_sheet('Hey, Dude')

        fmts = [
            'M/D/YY',
            'D-MMM-YY',
            'D-MMM',
            'MMM-YY',
            'h:mm AM/PM',
            'h:mm:ss AM/PM',
            'h:mm',
            'h:mm:ss',
            'M/D/YY h:mm',
            'mm:ss',
            '[h]:mm:ss',
            'mm:ss.0',
            ]

        i = 0
        for fmt in fmts:
            ws.write(i, 0, fmt)

            style = XFStyle()
            style.num_format_str = fmt

            ws.write(i, 4, datetime(2013, 5, 10), style)

            i += 1

        stream = BytesIO()
        w.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('82fed69d4f9ea0444d159fa30080f0a3', md5.hexdigest())

    def test_format(self):
        font0 = Font()
        font0.name = 'Times New Roman'
        font0.struck_out = True
        font0.bold = True

        style0 = XFStyle()
        style0.font = font0

        wb = Workbook()
        ws0 = wb.add_sheet('0')

        ws0.write(1, 1, 'Test', style0)

        for i in range(0, 0x53):
            fnt = Font()
            fnt.name = 'Arial'
            fnt.colour_index = i
            fnt.outline = True

            borders = Borders()
            borders.left = i

            style = XFStyle()
            style.font = fnt
            style.borders = borders

            ws0.write(i, 2, 'colour', style)
            ws0.write(i, 3, hex(i), style0)

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('da1e62ee3e7f9a09430e7a8227cb1215', md5.hexdigest())

    def test_formula(self):
        wb = Workbook()
        ws = wb.add_sheet('F')

        ws.write(0, 0, Formula("-(1+1)"))
        ws.write(1, 0, Formula("-(1+1)/(-2-2)"))
        ws.write(2, 0, Formula("-(134.8780789+1)"))
        ws.write(3, 0, Formula("-(134.8780789e-10+1)"))
        ws.write(4, 0, Formula("-1/(1+1)+9344"))

        ws.write(0, 1, Formula("-(1+1)"))
        ws.write(1, 1, Formula("-(1+1)/(-2-2)"))
        ws.write(2, 1, Formula("-(134.8780789+1)"))
        ws.write(3, 1, Formula("-(134.8780789e-10+1)"))
        ws.write(4, 1, Formula("-1/(1+1)+9344"))

        ws.write(0, 2, Formula("A1*B1"))
        ws.write(1, 2, Formula("A2*B2"))
        ws.write(2, 2, Formula("A3*B3"))
        ws.write(3, 2, Formula("A4*B4*sin(pi()/4)"))
        ws.write(4, 2, Formula("A5%*B5*pi()/1000"))

        ##############
        ## NOTE: parameters are separated by semicolon!!!
        ##############


        ws.write(5, 2, Formula("C1+C2+C3+C4+C5/(C1+C2+C3+C4/(C1+C2+C3+C4/(C1+C2+C3+C4)+C5)+C5)-20.3e-2"))
        ws.write(5, 3, Formula("C1^2"))
        ws.write(6, 2, Formula("SUM(C1;C2;;;;;C3;;;C4)"))
        ws.write(6, 3, Formula("SUM($A$1:$C$5)"))

        ws.write(7, 0, Formula('"lkjljllkllkl"'))
        ws.write(7, 1, Formula('"yuyiyiyiyi"'))
        ws.write(7, 2, Formula('A8 & B8 & A8'))
        ws.write(8, 2, Formula('now()'))

        ws.write(10, 2, Formula('TRUE'))
        ws.write(11, 2, Formula('FALSE'))
        ws.write(12, 3, Formula('IF(A1>A2;3;"hkjhjkhk")'))

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('2ce3745746c1b254f916623b92c3cec0', md5.hexdigest())

    def test_hyperlink(self):
        f = Font()
        f.height = 20*72
        f.name = 'Verdana'
        f.bold = True
        f.underline = Font.UNDERLINE_DOUBLE
        f.colour_index = 4

        h_style = XFStyle()
        h_style.font = f

        w = Workbook()
        ws = w.add_sheet('F')

        ##############
        ## NOTE: parameters are separated by semicolon!!!
        ##############

        n = "HYPERLINK"
        ws.write_merge(1, 1, 1, 10, Formula(n + '("http://www.irs.gov/pub/irs-pdf/f1000.pdf";"f1000.pdf")'), h_style)
        ws.write_merge(2, 2, 2, 25, Formula(n + '("mailto:roman.kiseliov@gmail.com?subject=pyExcelerator-feedback&Body=Hello,%20Roman!";"pyExcelerator-feedback")'), h_style)

        stream = BytesIO()
        w.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('420ff541deaad546519c8748118954d6', md5.hexdigest())

    def test_image(self):
        w = Workbook()
        ws = w.add_sheet('Image')
        ws.insert_bitmap(self.python_bmp, 2, 2)
        ws.insert_bitmap(self.python_bmp, 10, 2)

        stream = BytesIO()
        w.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('e3462afe1c0a38e1395c50e819d1b6cc', md5.hexdigest())

    def test_img_chg_col_wid(self):
        w = Workbook()
        ws = w.add_sheet('Image')

        ws.write(0, 2, "chg wid: none")
        ws.insert_bitmap(self.python_bmp, 2, 2)

        ws.write(0, 4, "chg wid: after")
        ws.insert_bitmap(self.python_bmp, 2, 4)
        ws.col(4).width = 20 * 256

        ws.write(0, 6, "chg wid: before")
        ws.col(6).width = 20 * 256
        ws.insert_bitmap(self.python_bmp, 2, 6)

        ws.write(0, 8, "chg wid: after")
        ws.insert_bitmap(self.python_bmp, 2, 8)
        ws.col(5).width = 8 * 256

        ws.write(0, 10, "chg wid: before")
        ws.col(10).width = 8 * 256
        ws.insert_bitmap(self.python_bmp, 2, 10)

        stream = BytesIO()
        w.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('765cdbe8ae8e27a81ba54d07861bdce5', md5.hexdigest())

    def test_merged(self):
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 4
        fnt.bold = True

        borders = Borders()
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6

        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER

        style = XFStyle()
        style.font = fnt
        style.borders = borders
        style.alignment = al

        wb = Workbook()
        ws0 = wb.add_sheet('sheet0')
        ws1 = wb.add_sheet('sheet1')
        ws2 = wb.add_sheet('sheet2')

        for i in range(0, 0x200, 2):
            ws0.write_merge(i, i+1, 1, 5, 'test %d' % i, style)
            ws1.write_merge(i, i, 1, 7, 'test %d' % i, style)
            ws2.write_merge(i, i+1, 1, 7 + (i%10), 'test %d' % i, style)

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('1ec1ec7bc76911f3a100b6573711b818', md5.hexdigest())

    def test_merged0(self):
        wb = Workbook()
        ws0 = wb.add_sheet('0')

        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 4
        fnt.bold = True

        borders = Borders()
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6

        style = XFStyle()
        style.font = fnt
        style.borders = borders

        ws0.write_merge(3, 3, 1, 5, 'test1', style)
        ws0.write_merge(4, 10, 1, 5, 'test2', style)
        ws0.col(1).width = 0x0d00

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('b55f49cfd1fb786bd611ed5b02b9d16c', md5.hexdigest())

    def test_merged1(self):
        wb = Workbook()
        ws0 = wb.add_sheet('0')

        fnt1 = Font()
        fnt1.name = 'Verdana'
        fnt1.bold = True
        fnt1.height = 18*0x14

        pat1 = Pattern()
        pat1.pattern = Pattern.SOLID_PATTERN
        pat1.pattern_fore_colour = 0x16

        brd1 = Borders()
        brd1.left = 0x06
        brd1.right = 0x06
        brd1.top = 0x06
        brd1.bottom = 0x06

        fnt2 = Font()
        fnt2.name = 'Verdana'
        fnt2.bold = True
        fnt2.height = 14*0x14

        brd2 = Borders()
        brd2.left = 0x01
        brd2.right = 0x01
        brd2.top = 0x01
        brd2.bottom = 0x01

        pat2 = Pattern()
        pat2.pattern = Pattern.SOLID_PATTERN
        pat2.pattern_fore_colour = 0x01F

        fnt3 = Font()
        fnt3.name = 'Verdana'
        fnt3.bold = True
        fnt3.italic = True
        fnt3.height = 12*0x14

        brd3 = Borders()
        brd3.left = 0x07
        brd3.right = 0x07
        brd3.top = 0x07
        brd3.bottom = 0x07

        fnt4 = Font()

        al1 = Alignment()
        al1.horz = Alignment.HORZ_CENTER
        al1.vert = Alignment.VERT_CENTER

        al2 = Alignment()
        al2.horz = Alignment.HORZ_RIGHT
        al2.vert = Alignment.VERT_CENTER

        al3 = Alignment()
        al3.horz = Alignment.HORZ_LEFT
        al3.vert = Alignment.VERT_CENTER

        style1 = XFStyle()
        style1.font = fnt1
        style1.alignment = al1
        style1.pattern = pat1
        style1.borders = brd1

        style2 = XFStyle()
        style2.font = fnt2
        style2.alignment = al1
        style2.pattern = pat2
        style2.borders = brd2

        style3 = XFStyle()
        style3.font = fnt3
        style3.alignment = al1
        style3.pattern = pat2
        style3.borders = brd3

        price_style = XFStyle()
        price_style.font = fnt4
        price_style.alignment = al2
        price_style.borders = brd3
        price_style.num_format_str = '_(#,##0.00_) "money"'

        ware_style = XFStyle()
        ware_style.font = fnt4
        ware_style.alignment = al3
        ware_style.borders = brd3


        ws0.merge(3, 3, 1, 5, style1)
        ws0.merge(4, 10, 1, 6, style2)
        ws0.merge(14, 16, 1, 7, style3)
        ws0.col(1).width = 0x0d00

        stream = BytesIO()
        wb.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('335d9b34aa4cf170d7416b5313a3a260', md5.hexdigest())

    def test_number_formats(self):
        w = Workbook()
        ws = w.add_sheet('Hey, Dude')

        fmts = [
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

        i = 0
        for fmt in fmts:
            ws.write(i, 0, fmt)

            style = XFStyle()
            style.num_format_str = fmt

            ws.write(i, 4, -1278.9078, style)

            i += 1

        stream = BytesIO()
        w.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('a87b0d3626f3ab3a2ece9bd03df1bf79', md5.hexdigest())

if __name__ == '__main__':

    unittest.main()
