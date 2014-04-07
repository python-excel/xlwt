#!/usr/bin/env python
#coding:utf-8
# Author:  mozman --<mozman@gmx.at>
# Purpose: test_mini
# Created: 03.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import os
import unittest
import filecmp

import xlwt

def from_tst_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

def create_example_xls(filename):
    w = xlwt.Workbook()
    ws1 = w.add_sheet(u'\N{GREEK SMALL LETTER ALPHA}\N{GREEK SMALL LETTER BETA}\N{GREEK SMALL LETTER GAMMA}')

    ws1.write(0, 0, u'\N{GREEK SMALL LETTER ALPHA}\N{GREEK SMALL LETTER BETA}\N{GREEK SMALL LETTER GAMMA}')
    ws1.write(1, 1, u'\N{GREEK SMALL LETTER DELTA}x = 1 + \N{GREEK SMALL LETTER DELTA}')

    ws1.write(2,0, u'A\u2262\u0391.')     # RFC2152 example
    ws1.write(3,0, u'Hi Mom -\u263a-!')   # RFC2152 example
    ws1.write(4,0, u'\u65E5\u672C\u8A9E') # RFC2152 example
    ws1.write(5,0, u'Item 3 is \u00a31.') # RFC2152 example
    ws1.write(8,0, u'\N{INTEGRAL}')       # RFC2152 example

    w.add_sheet(u'A\u2262\u0391.')     # RFC2152 example
    w.add_sheet(u'Hi Mom -\u263a-!')   # RFC2152 example
    one_more_ws = w.add_sheet(u'\u65E5\u672C\u8A9E') # RFC2152 example
    w.add_sheet(u'Item 3 is \u00a31.') # RFC2152 example

    one_more_ws.write(0, 0, u'\u2665\u2665')

    w.add_sheet(u'\N{GREEK SMALL LETTER ETA WITH TONOS}')
    w.save(filename)

EXAMPLE_XLS = 'unicode1.xls'

class TestUnicode1(unittest.TestCase):

    def test_example_xls(self):
        create_example_xls(EXAMPLE_XLS)
        self.assertTrue(filecmp.cmp(from_tst_dir(EXAMPLE_XLS),
                                    from_tst_dir(os.path.join('output-0.7.2', EXAMPLE_XLS)),
                                    shallow=False))
if __name__=='__main__':
    unittest.main()