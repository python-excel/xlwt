#!/usr/bin/env python
# coding:utf-8
# Author:  Fabio --<fabio_ferreiradasilva@yahoo.com.br>
# Purpose: test functions by sheet name
# Created: 30.07.2014

import unittest
import xlwt

class TestByName(unittest.TestCase):
    def setUp(self):
        self.wb = xlwt.Workbook()
        self.wb.add_sheet('Plan1')    
        self.wb.add_sheet('Plan2')
        self.wb.add_sheet('Plan3')
        self.wb.add_sheet('Plan4')

    def test_sheet_index(self):
        'Return sheet index by sheet name'
        idx = self.wb.sheet_index('Plan3')
        self.assertEqual(2, idx)

    def test_get_by_name(self):
        'Get sheet by name'
        ws = self.wb.get_sheet_by_name('Plan2')
        self.assertEqual('Plan2', ws.name)

if __name__=='__main__':
    unittest.main()
