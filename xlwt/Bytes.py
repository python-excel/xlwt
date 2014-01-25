# -*- coding: windows-1252 -*-
import sys

if sys.version_info[0] < 3:
    def bytes(value='', encoding='', errors=''):
        return value
