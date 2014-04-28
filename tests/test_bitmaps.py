#!/usr/bin/env python
#coding:utf-8
# Author:  Ross Golder --<ross@golder.org>
# Purpose: test_bitmaps
# Created: 23/04/2014
# Copyright (C) 2014, Ross Golder
# License: BSD licence

import sys
import os
import unittest

import xlwt

class TestBitmaps(unittest.TestCase):
    def test_create_bitmap_from_memory(self):
        book = xlwt.Workbook()
        sheet = book.add_sheet('bitmap from memory')
        
        with open ("tests/boats.bmp", "r") as bmpfile:
            bmpdata = bmpfile.read()
            sheet.insert_bitmap_data(bmpdata, 1, 1)
        
        self.assertTrue(True)

    def test_create_bitmap_from_file(self):
        book = xlwt.Workbook()
        sheet = book.add_sheet('bitmap from file')
        
        sheet.insert_bitmap("tests/boats.bmp", 1, 1)
        
        self.assertTrue(True)

if __name__=='__main__':
    unittest.main()
