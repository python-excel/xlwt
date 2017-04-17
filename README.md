### Python package "xlwt"

**Purpose:** Provide a library for developers to use to generate
spreadsheet files compatible with Microsoft Excel versions 95 to 2003.

**Maintainer:** John Machin, Lingfo Pty Ltd (sjmachin@lexicon.net)

**Licence:** BSD-style (see licences.py)

**Version of xlwt:** 0.7.2

**Version of Python required:** 2.3 to 2.6

**External modules required:** None

The package itself is pure Python with no dependencies on modules or
packages outside the standard Python distribution.

**Quick start:**

    import xlwt
    from datetime import datetime

    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
        num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')

    ws.write(0, 0, 1234.56, style0)
    ws.write(1, 0, datetime.now(), style1)
    ws.write(2, 0, 1)
    ws.write(2, 1, 1)
    ws.write(2, 2, xlwt.Formula("A3+B3"))

    wb.save('example.xls')

**Installation:**

-   Any OS: Unzip the .zip file into a suitable directory, chdir to that
    directory, then do "python setup.py install".
-   If PYDIR is your Python installation directory: the main files are
    in PYDIR/Lib/site-packages/xlwt, docs are in the doc subdirectory.
-   If os.sep != "/": make the appropriate adjustments.

**Download URLs:**

-   Packaged: http://pypi.python.org/pypi/xlwt
-   SVN: https://secure.simplistix.co.uk/svn/xlwt/trunk

**Documentation:**

Documentation can be found in the 'doc' directory of the xlwt package.
If these aren't sufficient, please consult the code in the examples
directory and the source code itself.

**Problems:**

Try the following in this order:

1.  Read the source
2.  Ask a question on http://groups.google.com/group/python-excel/
3.  E-mail the xlwt maintainer (sjmachin at lexicon.net), including
    "[xlwt]" as part of the message subject.

**Acknowledgements:**

-   xlwt is a fork of the pyExcelerator package, which was developed by
    Roman V. Kiseliov. "This product includes software developed by
    Roman V. Kiseliov \<roman@kiseliov.ru\>."
-   xlwt uses ANTLR v 2.7.7 to generate its formula compiler.
-   \<\< a growing list of names; see HISTORY.html \>\>: feedback,
    testing, test files, ...

