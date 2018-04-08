import sys

PY3 = sys.version_info[0] >= 3

if PY3:
    unicode = bytes.decode
    unicode_type = str
    basestring = str
    int_types = (int,)
    long = int

    def itervalues(d):
        return iter(d.values())
else:
    # Python 2
    unicode = unicode_type = unicode
    basestring = basestring
    int_types = (int, long)
    long = long

    def itervalues(d):
        return d.itervalues()
