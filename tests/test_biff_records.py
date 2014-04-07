#!/usr/bin/env python
#coding:utf-8
# Author:  mozman --<mozman@gmx.at>
# Purpose: test BIFF records
# Created: 09.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import unittest

from xlwt3 import biffrecords

class TestSharedStringTable(unittest.TestCase):
    def test_shared_string_table(self):
        expected_result = b'\xfc\x00\x11\x00\x01\x00\x00\x00\x01\x00\x00\x00\x03\x00\x01\x1e\x04;\x04O\x04'
        string_record = biffrecords.SharedStringTable(encoding='cp1251')
        string_record.add_str('Оля')
        self.assertEqual(expected_result, string_record.get_biff_record())

if __name__=='__main__':
    unittest.main()