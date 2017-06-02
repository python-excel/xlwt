__author__ = 'hrwl'

import unittest

from xlwt.Utils import *


class UtilsTestCase(unittest.TestCase):
    def test_col_by_name(self):
        self.assertEqual(col_by_name('A'), 0)
        self.assertEqual(col_by_name('AA'), 26)

    def test_cell_to_rowcoll(self):
        x = cell_to_rowcol('AA7')
        self.assertEqual(x, (6, 26, False, False))
        y = cell_to_rowcol('$A$9')
        self.assertEqual(y, (8, 0, True, True))

    def test_cell_to_rowcoll2(self):
        row, col = cell_to_rowcol2('AA7')
        self.assertEqual(col, 26)
        self.assertEqual(row, 6)

    def test_rowcol_to_cell(self):
        cell = rowcol_to_cell(5, 27)
        self.assertEqual(cell, 'AB6')
        cell = rowcol_to_cell(5, 27, True, True)
        self.assertEqual(cell, '$AB$6')

    def test_rowcol_pair_to_cellrange(self):
        cell = rowcol_pair_to_cellrange(5, 27, 9, 30)
        self.assertEqual(cell, 'AB6:AE10')

    def test_cellrange_to_rowcol_pair(self):
        x = cellrange_to_rowcol_pair('AB6:AE10')
        self.assertEqual(x, (5, 27, 9, 30))

    def test_cell_to_packed_rowcol(self):
        x = cell_to_packed_rowcol('AA7')
        self.assertEqual(x, (6, 49178))
        y = cell_to_packed_rowcol('$AA7')
        self.assertEqual(y, (6, 32794))



if __name__ == '__main__':
    unittest.main()
