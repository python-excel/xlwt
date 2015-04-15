import sys
import os
import unittest
import filecmp
from datetime import datetime

from utils import in_tst_dir, in_tst_output_dir

import xlwt

class TestSimple(unittest.TestCase):
    def create_simple_xls(self):
        font0 = xlwt.Font()
        font0.name = 'Times New Roman'
        font0.colour_index = 2
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        style1 = xlwt.XFStyle()
        style1.num_format_str = 'D-MMM-YY'

        wb = xlwt.Workbook()
        ws = wb.add_sheet('A Test Sheet')

        ws.write(0, 0, 'Test', style0)
        ws.write(1, 0, datetime(2010, 12, 5), style1)
        ws.write(2, 0, 1)
        ws.write(2, 1, 1)
        ws.write(2, 2, xlwt.Formula("A3+B3"))

        wb.save(in_tst_output_dir('simple.xls'))

    def test_create_simple_xls(self):
        self.create_simple_xls()
        self.assertTrue(filecmp.cmp(in_tst_dir('simple.xls'),
                                    in_tst_output_dir('simple.xls'),
                                    shallow=False))
