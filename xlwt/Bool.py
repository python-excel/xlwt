# -*- coding: windows-1252 -*-
import sys
version = sys.version.split()[0]
if version < '2.2.1':
    False = 0
if version < '2.3':
    True = not False
