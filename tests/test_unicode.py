__author__ = 'hrwl'

import hashlib
import six
import unittest

from xlwt import Workbook


class UnicodeTestCase(unittest.TestCase):
    def test_unicode(self):
        book = Workbook(encoding='cp1251')
        sheet = book.add_sheet('cp1251-demo')
        sheet.write(0, 0, six.b('\xCE\xEB\xFF'))

        stream = six.BytesIO()
        book.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('27dd1a8f898c84c58be1b6ce3e704371', md5.hexdigest())

    def test_unicode1(self):
        book = Workbook()
        ws1 = book.add_sheet(six.u('\N{GREEK SMALL LETTER ALPHA}\N{GREEK SMALL LETTER BETA}\N{GREEK SMALL LETTER GAMMA}'))

        ws1.write(0, 0, six.u('\N{GREEK SMALL LETTER ALPHA}\N{GREEK SMALL LETTER BETA}\N{GREEK SMALL LETTER GAMMA}'))
        ws1.write(1, 1, six.u('\N{GREEK SMALL LETTER DELTA}x = 1 + \N{GREEK SMALL LETTER DELTA}'))

        ws1.write(2, 0, six.u('A\u2262\u0391.'))      # RFC2152 example
        ws1.write(3, 0, six.u('Hi Mom -\u263a-!'))    # RFC2152 example
        ws1.write(4, 0, six.u('\u65E5\u672C\u8A9E'))  # RFC2152 example
        ws1.write(5, 0, six.u('Item 3 is \u00a31.'))  # RFC2152 example
        ws1.write(8, 0, six.u('\N{INTEGRAL}'))        # RFC2152 example

        book.add_sheet(six.u('A\u2262\u0391.'))     # RFC2152 example
        book.add_sheet(six.u('Hi Mom -\u263a-!'))   # RFC2152 example
        one_more_ws = book.add_sheet(six.u('\u65E5\u672C\u8A9E'))  # RFC2152 example
        book.add_sheet(six.u('Item 3 is \u00a31.'))  # RFC2152 example

        one_more_ws.write(0, 0, six.u('\u2665\u2665'))

        book.add_sheet(six.u('\N{GREEK SMALL LETTER ETA WITH TONOS}'))

        stream = six.BytesIO()
        book.save(stream)
        md5 = hashlib.md5()
        md5.update(stream.getvalue())
        self.assertEqual('0049fc8cdd164385c45198d2a75a4155', md5.hexdigest())

if __name__ == '__main__':
    unittest.main()

