#!/usr/bin/env python
# tries stress SST, SAT and MSAT

from __future__ import print_function
from __future__ import absolute_import
from time import *
from xlwt import *

style = XFStyle()

wb = Workbook()
ws0 = wb.add_sheet('0')

colcount = 200 + 1
rowcount = 6000 + 1

t0 = time()

for col in range(colcount):
    for row in range(rowcount):
        ws0.write(row, col, "BIG(%d, %d)" % (row, col))

t1 = time() - t0
print("\nsince starting elapsed %.2f s" % (t1))

print("Storing...")
wb.save('big-35Mb.xls')

t2 = time() - t0
print("since starting elapsed %.2f s" % (t2))


