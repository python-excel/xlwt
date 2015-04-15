# coding:utf-8

import unittest

from xlwt import BIFFRecords

class TestSharedStringTable(unittest.TestCase):
    def test_shared_string_table(self):
        expected_result = b'\xfc\x00\x11\x00\x01\x00\x00\x00\x01\x00\x00\x00\x03\x00\x01\x1e\x04;\x04O\x04'
        string_record = BIFFRecords.SharedStringTable(encoding='cp1251')
        string_record.add_str(u'Оля')
        self.assertEqual(expected_result, string_record.get_biff_record())
