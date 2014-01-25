__VERSION__ = '0.7.5'

from Workbook import Workbook
from Worksheet import Worksheet
from Row import Row
from Column import Column
from Formatting import Font, Alignment, Borders, Pattern, Protection
from Style import XFStyle, easyxf, easyfont, add_palette_colour
from ExcelFormula import *

import sys
if sys.version_info[0] > 2:
    Workbook.Workbook = Workbook
    Worksheet.Worksheet = Worksheet
    Row.Row = Row
    Column.Column = Column
