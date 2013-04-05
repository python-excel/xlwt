#!/usr/bin/env python

from distutils.core import setup
from xlwt import __VERSION__

DESCRIPTION = (
    'Library to create spreadsheet files compatible with '
    'MS Excel 97/2000/XP/2003 XLS files, '
    'on any platform, with Python 2.3 to 2.7'
    )

LONG_DESCRIPTION = """\
xlwt is a library for generating spreadsheet files that are compatible
with Excel 97/2000/XP/2003, OpenOffice.org Calc, and Gnumeric. xlwt has
full support for Unicode. Excel spreadsheets can be generated on any
platform without needing Excel or a COM server. The only requirement is
Python 2.3 to 2.7.
"""

CLASSIFIERS = [
    'Operating System :: OS Independent',
    'Programming Language :: Python',
    'License :: OSI Approved :: BSD License',
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'Topic :: Software Development :: Libraries :: Python Modules',
    'Topic :: Office/Business :: Financial :: Spreadsheet',
    'Topic :: Database',
    'Topic :: Internet :: WWW/HTTP :: Dynamic Content :: CGI Tools/Libraries',
    ]

KEYWORDS = (
    'xls excel spreadsheet workbook worksheet pyExcelerator'
    )

setup(
    name = 'xlwt',
    version = __VERSION__,
    maintainer = 'John Machin',
    maintainer_email = 'sjmachin@lexicon.net',
    url = 'http://www.python-excel.org/',
    download_url = 'http://pypi.python.org/pypi/xlwt',
    description = DESCRIPTION,
    long_description = LONG_DESCRIPTION,
    license = 'BSD',
    platforms = 'Platform Independent',
    packages = ['xlwt'],
    keywords = KEYWORDS,
    classifiers = CLASSIFIERS,
    package_data = {
        'xlwt': [
            'doc/*.*',
            'examples/*.*',
            'tests/*.*',
            ],
        },
    )
