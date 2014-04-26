import sys

PY3 = sys.version_info[0] >= 3

if PY3:
    unicode = bytes.decode
    unicode_type = str
    basestring = str
else:
    unicode = unicode_type = unicode
    basestring = basestring