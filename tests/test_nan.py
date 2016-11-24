import filecmp
import unittest

import xlwt
from utils import in_tst_dir, in_tst_output_dir


class TestNan(unittest.TestCase):
    def create_nan_xls(self, filename='nan.xls'):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('A Test Sheet')

        ws.write(0, 0, 1)
        ws.write(0, 1, float('nan'))
        wb.save(in_tst_output_dir(filename))

    def test_create_nan_xls(self):
        self.create_nan_xls()
        self.assertTrue(filecmp.cmp(in_tst_dir('nan.xls'),
                                    in_tst_output_dir('nan.xls'),
                                    shallow=False))
