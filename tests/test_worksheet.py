
import unittest

import xlrd
from utils import in_tst_output_dir

import xlwt


class TestWorksheet(unittest.TestCase):

    result_filename = 'Result.xls'
    result_sheetname = 'result_sheetname'

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

    def test_write_column(self):
        workbook = xlwt.Workbook()
        workbook.add_sheet(self.result_sheetname)
        sheet = workbook.get_sheet(self.result_sheetname)
        testarray = [1, 2, 3, 4]
        sheet.write_column(3, ar = testarray)
        workbook.save(self.result_filename)

        resultbook = xlrd.open_workbook(self.result_filename)
        resultsheet = resultbook.sheet_by_index(0)

        for row_num in range(len(testarray)):
            for col_num in range(3):
                self.assertEqual('', resultsheet.cell(row_num, col_num).value)
            self.assertEqual(testarray[row_num], resultsheet.cell(row_num, 3).value)

    def test_write_column_error_on_overwrite(self):
        workbook = xlwt.Workbook()
        workbook.add_sheet(self.result_sheetname)
        sheet = workbook.get_sheet(self.result_sheetname)
        testarray = [1, 2, 3, 4]
        sheet.write_column(3, ar = testarray)
        try:
            with self.assertRaisesRegex(Exception, 'Attempt to overwrite cell'):
                sheet.write(0, 3, '')
        except AttributeError:
            with self.assertRaises(Exception):
                sheet.write(0, 3, '')
