import sys

PY2 = sys.version_info[0] == 2

if PY2:
    # Python 2
    unicode = unicode_type = unicode
    basestring = basestring
    xrange = xrange
    int_types = (int, long)
    long = long

    def iteritems(d):
        return d.iteritems()
    def itervalues(d):
        return d.itervalues()
else:
    unicode = bytes.decode
    unicode_type = str
    basestring = str
    xrange = range
    int_types = (int,)
    long = int

    def iteritems(d):
        return iter(d.items())
    def itervalues(d):
        return iter(d.values())
