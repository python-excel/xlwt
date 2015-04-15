Changes
=======

.. currentmodule:: xlwt

0.7.2 (1 June 2009)
-------------------

- Added function Utils.rowcol_pair_to_cellrange.
  ``(0, 0, 65535, 255) -> "A1:IV65536"``

- Removed :class:`Worksheet` property ``show_empty_as_zero``,
  and added attribute :attr:`~Worksheet.show_zero_values`
  (default: ``1 == True``).

- Fixed formula code generation problem with formulas
  including MAX/SUM/etc functions with arguments like A1+123.

- Added .pattern_examples.xls and put a pointer to it
  in the easyxf part of Style.py.

- Fixed Row.set_cell_formula() bug introduced in 0.7.1.

- Fixed bug(?) with SCL/magnification handling causing(?) Excel
  to raise a dialogue box if sheet is set to open in page preview mode
  and user then switches to normal view.

- Added color and colour as synonyms for font.colour_index in easyxf.

- Removed unused attribute Row.__has_default_format.

0.7.1 (4 March 2009)
--------------------

See source control for changes made.

0.7.0 (19 September 2008)
-------------------------

- Fixed more bugs and added more various new bits of functionality

0.7.0a4 (8 October 2007)
------------------------

- fork of pyExcelerator, released to python-excel.

- Fixed various bugs in pyExcelerator and added various new bits of functionality
