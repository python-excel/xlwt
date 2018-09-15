import unittest

import xlwt

from .utils import in_tst_output_dir


class TestWorksheet(unittest.TestCase):
    def test_worksheet(self):
        book = xlwt.Workbook()
        first_page = "Page 1"

        sheet = book.add_sheet(first_page)
        assert sheet.get_name() == first_page

        sheet.set_name('Page 2')
        assert sheet.get_name() == 'Page 2'

        sheet.set_header_str(first_page)
        assert sheet.get_header_str() == first_page

        sheet.set_footer_str(first_page)
        assert sheet.get_footer_str() == first_page

        sheet.set_left_margin(0.45)
        assert sheet.get_left_margin() == 0.45

        sheet.set_right_margin(0.45)
        assert sheet.get_right_margin() == 0.45

        sheet.set_top_margin(0.45)
        assert sheet.get_top_margin() == 0.45

        sheet.set_bottom_margin(0.45)
        assert sheet.get_bottom_margin() == 0.45

        sheet.set_header_margin(0.45)
        assert sheet.get_header_margin() == 0.45

        sheet.set_footer_margin(0.45)
        assert sheet.get_footer_margin() == 0.45

        sheet.set_portrait(False)
        assert sheet.get_portrait() is False

        sheet.set_fit_num_pages(2)
        assert sheet.get_fit_num_pages() == 2

        sheet.set_fit_num_pages(1)
        assert sheet.get_fit_num_pages() == 1

        sheet.set_fit_width_to_pages(1)
        assert sheet.get_fit_width_to_pages() == 1

        sheet.set_fit_height_to_pages(0)
        assert sheet.get_fit_height_to_pages() == 0

        book.save(in_tst_output_dir('worksheet.xls'))
