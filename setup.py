#!/usr/bin/env python
# -*- coding: ascii -*-

#  Copyright (C) 2005 Roman V. Kiseliov
#  All rights reserved.
# 
#  Redistribution and use in source and binary forms, with or without
#  modification, are permitted provided that the following conditions
#  are met:
# 
#  1. Redistributions of source code must retain the above copyright
#     notice, this list of conditions and the following disclaimer.
# 
#  2. Redistributions in binary form must reproduce the above copyright
#     notice, this list of conditions and the following disclaimer in
#     the documentation and/or other materials provided with the
#     distribution.
# 
#  3. All advertising materials mentioning features or use of this
#     software must display the following acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <roman@kiseliov.ru>."
# 
#  4. Redistributions of any form whatsoever must retain the following
#     acknowledgment:
#     "This product includes software developed by
#      Roman V. Kiseliov <roman@kiseliov.ru>."
# 
#  THIS SOFTWARE IS PROVIDED BY Roman V. Kiseliov ``AS IS'' AND ANY
#  EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
#  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
#  PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL Roman V. Kiseliov OR
#  ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
#  SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
#  NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
#  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
#  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
#  STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
#  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
#  OF THE POSSIBILITY OF SUCH DAMAGE.

__rev_id__ = """$Id$"""

import sys
from distutils.core import setup

DESCRIPTION = \
'Library to create spreadsheet files compatible with MS Excel 97/2000/XP/2003 XLS files, ' \
'on any platform, with Python 2.3 or later'


LONG_DESCRIPTION = \
'''xlwt is a library for generating spreadsheet files that are compatible with
Excel 97/2000/XP/2003, OpenOffice.org Calc, and Gnumeric. xlwt has
full support for Unicode. Excel spreadsheets can be generated on any platform
without needing Excel or a COM server. The only requirement is Python 2.3 or higher. 
'''

CLASSIFIERS = \
[
 'Operating System :: OS Independent',
 'Programming Language :: Python',
 'License :: OSI Approved :: BSD License',
 'Development Status :: 3 - Alpha',
 'Intended Audience :: Developers',
 'Topic :: Software Development :: Libraries :: Python Modules',
 'Topic :: Office/Business :: Financial :: Spreadsheet',
 'Topic :: Database',
 'Topic :: Internet :: WWW/HTTP :: Dynamic Content :: CGI Tools/Libraries',
]

KEYWORDS = \
'xls excel spreadsheet workbook worksheet'

setup(name = 'xlwt',
      version = '0.7.0a4',
      # author = 'Roman V. Kiseliov',
      # author_email = 'roman@kiseliov.ru',
      maintainer = 'John Machin',
      maintainer_email = 'sjmachin@lexicon.net',
      url = 'https://secure.simplistix.co.uk/svn/xlwt/trunk',
      description = DESCRIPTION,
      long_description = LONG_DESCRIPTION,
      license = 'BSD',
      platforms = 'Platform Independent',
      packages = ['xlwt'],
      keywords = KEYWORDS,
      classifiers = CLASSIFIERS
      )
