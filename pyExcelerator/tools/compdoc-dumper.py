#!/usr/bin/env python
# -*- coding: windows-1251 -*-

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


import sys, struct

def print_bin_data(data):
    print
    i = 0
    while i < len(data):
        j = 0
        while (i < len(data)) and (j < 16):
            c = '0x%02X' % ord(data[i])
            sys.stdout.write(c)
            sys.stdout.write(' ')
            i += 1
            j += 1
        print
    if i == 0:
        print '<NO DATA>'


def main():
    if len(sys.argv) < 2:
        print 'no input files.'
        print sys.exit(1)

    doc = file(sys.argv[1], 'rb').read()
    header, data = doc[0:512], doc[512:]
    del doc

    doc_magic             = header[0:8]
    file_uid              = header[8:24]
    rev_num               = header[24:26]
    ver_num               = header[26:28]
    byte_order            = header[28:30]
    log_sect_size,        = struct.unpack('<H', header[30:32])
    log_short_sect_size,  = struct.unpack('<H', header[32:34])
    total_sat_sectors,    = struct.unpack('<L', header[44:48])
    dir_start_sid,        = struct.unpack('<l', header[48:52])
    min_stream_size,      = struct.unpack('<L', header[56:60])
    ssat_start_sid,       = struct.unpack('<l', header[60:64])
    total_ssat_sectors,   = struct.unpack('<L', header[64:68])
    msat_start_sid,       = struct.unpack('<l', header[68:72])
    total_msat_sectors,   = struct.unpack('<L', header[72:76])
    MSAT                  = struct.unpack('<109l', header[76:])
     
    sect_size        = 1 << log_sect_size
    short_sect_size  = 1 << log_short_sect_size


    SECTORS = []
    i = 0
    while i < len(data):
        SECTORS.append(data[i:i+sect_size])
        i += sect_size
    del data

    total_sectors = len(SECTORS)

    print 'file magic: '
    print_bin_data(doc_magic)
    print

    print 'file uid: '
    print_bin_data(file_uid)
    print

    print 'revision number: '
    print_bin_data(rev_num)
    print
 
    print 'version number: '
    print_bin_data(ver_num)
    print
    
    print 'byte order: '
    print_bin_data(byte_order)
    print
    
    print 'sector size                                :', hex(sect_size), sect_size
    print 'total sectors in file                      :', hex(total_sectors), total_sectors
    print 'short sector size                          :', hex(short_sect_size), short_sect_size
    print 'Total number of sectors used for the SAT   :', hex(total_sat_sectors), total_sat_sectors
    print 'SID of first sector of the directory stream:', hex(dir_start_sid), dir_start_sid
    print 'Minimum size of a standard stream          :', hex(min_stream_size), min_stream_size
    print 'SID of first sector of the SSAT            :', hex(ssat_start_sid), ssat_start_sid
    print 'Total number of sectors used for the SSAT  :', hex(total_ssat_sectors), total_ssat_sectors
    print 'SID of first additional sector of the MSAT :', hex(msat_start_sid), msat_start_sid
    print 'Total number of sectors used for the MSAT  :', hex(total_msat_sectors), total_msat_sectors
    print 'MSAT (only header part): \n', MSAT

    MSAT_2nd = []
    next = msat_start_sid
    while next > 0:
       msat_sector = struct.unpack('128l', SECTORS[next])
       for x in msat_sector[:128]:
           MSAT_2nd.append(x)
       next = msat_sector[-1]
    print 'additional MSAT sectors: \n', MSAT_2nd

    sat_sectors = [x for x in MSAT if x >=0]
    sat_sectors += [x for x in MSAT_2nd if x >=0]

    print 'SAT resides in following sectors:\n', sat_sectors

    SAT = ''.join([SECTORS[sect] for sect in sat_sectors])
    sat_sids_count = len(SAT) >> 2
    SAT = struct.unpack('<%dl'%sat_sids_count, SAT) # SIDs tuple

    print 'SAT content:\n', SAT

    dir_stream_sectors = [dir_start_sid]
    sid = dir_start_sid
    while SAT[sid] >= 0:
        next_in_chain = SAT[sid]
        dir_stream_sectors.append(next_in_chain)    
        sid = next_in_chain

    print 'Directory sectors:\n', dir_stream_sectors

    dir_stream = ''.join([SECTORS[sect] for sect in dir_stream_sectors])

    dir_entry_list = []
    i = 0
    while i < len(dir_stream):
        dir_entry_list.append(dir_stream[i:i+128]) # 128 -- dir entry size
        i += 128
    del dir_stream

    print 'total directory entries:', len(dir_entry_list)

    dentry_types = {
        0x00: 'Empty',
        0x01: 'User storage',
        0x02: 'User stream',
        0x03: 'LockBytes',
        0x04: 'Property',
        0x05: 'Root storage'
    }
    node_colours = {
        0x00: 'Red',
        0x01: 'Black'
    }
    for dentry in dir_entry_list:
        print 'DID', dir_entry_list.index(dentry)
        name = dentry[0:64]
        name = name.split('\x00')
        name = ''.join(name)
        print 'dir entry name:', name

        sz, = struct.unpack('<H', dentry[64:66])
        print 'Size of the used area of the character buffer of the name:', sz

        t ,  = struct.unpack('B', dentry[66])
        print 'type of entry:', t, dentry_types[t]

        c ,  = struct.unpack('B', dentry[67])
        print 'entry colour:', c, node_colours[c]

        did_left ,  = struct.unpack('<l', dentry[68:72])
        print 'left child DID :', did_left
        did_right ,  = struct.unpack('<l', dentry[72:76])
        print 'right child DID:', did_right
        did_root ,  = struct.unpack('<l', dentry[76:80])
        print 'root DID       :', did_root
        
        dentry_start_sid ,  = struct.unpack('<l', dentry[116:120])
        print 'start SID       :', dentry_start_sid

        stream_sid_chain = [dentry_start_sid]
        sid = dentry_start_sid
        while SAT[sid] >= 0:
            next_in_chain = SAT[sid]
            stream_sid_chain.append(next_in_chain)    
            sid = next_in_chain
        print 'SID chain:\n', stream_sid_chain

        stream_size ,  = struct.unpack('<L', dentry[120:124])
        print 'stream size     :', stream_size

        print

main()
