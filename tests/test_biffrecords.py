__author__ = 'hrwl'

import unittest
import six

from xlwt import BIFFRecords


class SharedStringTableTestCase(unittest.TestCase):
    def test_shared_string_table(self):
        expected_result = six.b('\xfc\x00\x11\x00\x01\x00\x00\x00\x01\x00\x00\x00\x03\x00\x01\x1e\x04;\x04O\x04')
        string_record = BIFFRecords.SharedStringTable(encoding='cp1251')
        string_record.add_str(six.b('\xCE\xEB\xFF'))
        self.assertEqual(expected_result, string_record.get_biff_record())

        expected_result = six.b('\xfc\x00\x16\x00\x01\x00\x00\x00\x01\x00\x00\x00\x0b\x00\x00All around!')
        string_record = BIFFRecords.SharedStringTable(encoding='ascii')
        string_record.add_str(six.b('All around!'))
        self.assertEqual(expected_result, string_record.get_biff_record())


class Biff8BOFRecordTestCase(unittest.TestCase):
    def test_class(self):
        biff = BIFFRecords.Biff8BOFRecord(BIFFRecords.Biff8BOFRecord.BOOK_GLOBAL).get()
        self.assertEqual(six.b('\x09\x08\x10\x00\x00\x06\x05\x00\xbb\x0D\xcc\x07\x00\x00\x00\x00\x06\x00\x00\x00'), biff)


class WriteAccessRecordTestCase(unittest.TestCase):
    def test_class(self):
        biff = BIFFRecords.WriteAccessRecord('Hans').get()
        self.assertEqual(six.b('\\\x00p\x00Hans                                                                                                            '), biff)

if __name__ == '__main__':
    unittest.main()

