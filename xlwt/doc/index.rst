The xlwt Module
===============


**A Python package for generating Microsoft Excel spreadsheet files.**



General information
-------------------



State of Documentation
~~~~~~~~~~~~~~~~~~~~~~

This documentation is currently incomplete. There may be methods and
classes not included and any item marked with a *[NC]* is not complete
and may have further parameters, methods, attributes and functionality
that are not documented. In these cases, you'll have to refer to the
source if the documentation provided is insufficient.




Module Contents *[NC]*
----------------------

: **easyxf** (function): This function is used to create and configure
XFStyle objects for use with (for example) the Worksheet.write method.
: strg_to_parse : A string to be parsed to obtain attribute values for
Alignment, Borders, Font, Pattern and Protection objects. Refer to the
examples in the file .../examples/xlwt_easyxf_simple_demo.py and to
the xf_dict dictionary in Style.py. Various synonyms including
color/colour, center/centre and gray/grey are allowed. Case is
irrelevant (except maybe in font names). '-' may be used instead of
'_'. Example: "font: bold on; align: wrap on, vert centre, horiz
center"
: num_format_str : To get the "number format string" of an existing
cell whose format you want to reproduce, select the cell and click on
Format/Cells/Number/Custom. Otherwise, refer to Excel help. Examples:
"#,##0.00", "dd/mm/yyyy"
:Returns:: An object of the XFstyle class


: **Workbook** (class) [#]: The class to instantiate to create a
workbook For more information about this class, see The Workbook Class.

: **Worksheet** (class) [#]: A class to represent the contents of a
sheet in a workbook. For more information about this class, see The
Worksheet Class .




The Workbook Class *[NC]*
-------------------------

: **Workbook(encoding='ascii',style_compression=0)** (class) [#]: This
is a class representing a workbook and all its contents. When creating
Excel files with xlwt, you will normally start by instantiating an
object of this class.
: encoding : *[NC]*
: style_compression : *[NC]*
:Returns:: An object of the Workbook class


: **add_sheet(sheetname)** [#]: This method is used to create
Worksheets in a Workbook.
: sheetname : The name to use for this sheet, as it will appear in the
tabs at the bottom of the Excel application.
:Returns:: An object of the Worksheet class


: **save(filename_or_stream)** [#]: This method is used to save
Workbook to a file in native Excel format.
: filename_or_stream : This can be a string containing a filename of
the file, in which case the excel file is saved to disk using the name
provided. It can also be a stream object with a write method, such as
a StringIO, in which case the data for the excel file is written to
the stream.






The Worksheet Class *[NC]*
--------------------------

: **Worksheet(sheetname, parent_book)** (class) [#]: This is a class
representing the contents of a sheet in a workbook. WARNING: You don't
normally create instances of this class yourself. They are returned
from calls to Workbook.add_sheet
: **write(r, c, label="", style=Style.default_style)** [#]: This
method is used to write a cell to a Worksheet..
: r : The zero-relative number of the row in the worksheet to which
the cell should be written.
: c : The zero-relative number of the column in the worksheet to which
the cell should be written.
: label : The data value to be written. An int, long, or
decimal.Decimal instance is converted to float. A unicode instance is
written as is. A str instance is converted to unicode using the
encoding (default: 'ascii') specified when the Workbook instance was
created. A datetime.datetime, datetime.date, or datetime.time instance
is converted into Excel date format (a float representing the number
of days since (typically) 1899-12-31T00:00:00, under the pretence that
1900 was a leap year). A bool instance will show up as TRUE or FALSE
in Excel. None causes the cell to be blank -- no data, only
formatting. An xlwt.Formula instance causes an Excel formula to be
written. *[NC]*
: style : A style -- also known as an XF (extended format) -- is an
XFStyle object, which encapsulates the formatting applied to the cell
and its contents. XFStyle objects are best set up using the easyxf
function. They may also be set up by setting attributes in Alignment,
Borders, Pattern, Font and Protection objects then setting those
objects and a format string as attributes of an XFStyle object. *[NC]*






