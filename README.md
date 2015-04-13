This fork supports custom default style (original xlwt hardcodes Arial 10 as default font) and uses that style for all .write()'s on worksheets and rows when style is not specified. Default style applies to all empty cells (and row/column headers when viewed in MS Excel(tm)).

Use default_style kwargs of Workbook() constructor to specify your default style (can be constructed with XFStyle() or easyxf()). Refer to upstream documentation for rest features.
Side-effect: passing None as style to write() works like skipping style while upstream version throws an exception.

xlwt
====

Library to create spreadsheet files compatible with MS Excel 97/2000/XP/2003 XLS files, on any platform, with Python 2.3 to 2.6
