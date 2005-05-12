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

import UnicodeUtils
import struct

def get_wb_stream_data(filename):

    doc = file(filename, 'rb').read()
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

    #print 'file magic: '
    #print_bin_data(doc_magic)
    #print

    #print 'file uid: '
    #print_bin_data(file_uid)
    #print

    #print 'revision number: '
    #print_bin_data(rev_num)
    #print
 
    #print 'version number: '
    #print_bin_data(ver_num)
    #print
    
    #print 'byte order: '
    #print_bin_data(byte_order)
    #print
    
    #print 'sector size                                :', hex(sect_size), sect_size
    #print 'total sectors in file                      :', hex(total_sectors), total_sectors
    #print 'short sector size                          :', hex(short_sect_size), short_sect_size
    #print 'Total number of sectors used for the SAT   :', hex(total_sat_sectors), total_sat_sectors
    #print 'SID of first sector of the directory stream:', hex(dir_start_sid), dir_start_sid
    #print 'Minimum size of a standard stream          :', hex(min_stream_size), min_stream_size
    #print 'SID of first sector of the SSAT            :', hex(ssat_start_sid), ssat_start_sid
    #print 'Total number of sectors used for the SSAT  :', hex(total_ssat_sectors), total_ssat_sectors
    #print 'SID of first additional sector of the MSAT :', hex(msat_start_sid), msat_start_sid
    #print 'Total number of sectors used for the MSAT  :', hex(total_msat_sectors), total_msat_sectors
    #print 'MSAT (only header part): \n', MSAT

    MSAT_2nd = []
    next = msat_start_sid
    while next > 0:
       msat_sector = struct.unpack('128l', SECTORS[next])
       MSAT_2nd.extend(msat_sector[:127])
       next = msat_sector[-1]
    #print 'additional MSAT sectors: \n', MSAT_2nd

    sat_sectors = [x for x in MSAT if x >=0]
    sat_sectors += [x for x in MSAT_2nd if x >=0]

    #print 'SAT resides in following sectors:\n', sat_sectors

    SAT = ''.join([SECTORS[sect] for sect in sat_sectors])
    sat_sids_count = len(SAT) >> 2
    SAT = struct.unpack('<%dl'%sat_sids_count, SAT) # SIDs tuple

    #print 'SAT content:\n', SAT

    dir_stream_sectors = [dir_start_sid]
    sid = dir_start_sid
    while SAT[sid] >= 0:
        next_in_chain = SAT[sid]
        dir_stream_sectors.append(next_in_chain)    
        sid = next_in_chain

    #print 'Directory sectors:\n', dir_stream_sectors

    dir_stream = ''.join([SECTORS[sect] for sect in dir_stream_sectors])

    dir_entry_list = []
    i = 0
    while i < len(dir_stream):
        dir_entry_list.append(dir_stream[i:i+128]) # 128 -- dir entry size
        i += 128
    del dir_stream

    #print 'total directory entries:', len(dir_entry_list)

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
    
    wb_bin_data = ''
    for dentry in dir_entry_list:
        #print 'DID', dir_entry_list.index(dentry)
        name = dentry[0:64]
        name = name.split('\x00')
        name = ''.join(name)
        #print 'dir entry name:', name

        sz, = struct.unpack('<H', dentry[64:66])
        #print 'Size of the used area of the character buffer of the name:', sz

        t,  = struct.unpack('B', dentry[66])
        #print 'type of entry:', t, dentry_types[t]

        c,  = struct.unpack('B', dentry[67])
        #print 'entry colour:', c, node_colours[c]

        did_left ,  = struct.unpack('<l', dentry[68:72])
        #print 'left child DID :', did_left
        did_right ,  = struct.unpack('<l', dentry[72:76])
        #print 'right child DID:', did_right
        did_root ,  = struct.unpack('<l', dentry[76:80])
        #print 'root DID       :', did_root
        
        dentry_start_sid ,  = struct.unpack('<l', dentry[116:120])
        #print 'start SID       :', dentry_start_sid

        stream_sid_chain = [dentry_start_sid]
        sid = dentry_start_sid
        while SAT[sid] >= 0:
            next_in_chain = SAT[sid]
            stream_sid_chain.append(next_in_chain)    
            sid = next_in_chain
        #print 'SID chain:\n', stream_sid_chain

        stream_size ,  = struct.unpack('<L', dentry[120:124])
        #print 'stream size     :', stream_size
    
        if name == 'Workbook' or name == 'Book':
            for sid in stream_sid_chain:
                wb_bin_data += SECTORS[sid]
            #f = file('workbook.bin', 'wb')
            #f.write(wb_bin_data)
            #f.close()
            break

    return wb_bin_data


def parse_xls(filename):
    
    ##########################################################################

    def process_BOUNDSHEET(biff8, rec_data):
        sheet_stream_pos, visibility, sheet_type = struct.unpack('<I2B', rec_data[:6])
        sheet_name = rec_data[6:]

        if biff8:
            chars_num, options = struct.unpack('2B', sheet_name[:2])
            
            chars_start = 2
            runs_num = 0
            asian_phonetic_size = 0

            result = ''

            compressed = (options & 0x01) == 0
            has_asian_phonetic = (options & 0x04) != 0
            has_format_runs = (options & 0x08) != 0

            if has_format_runs:
                runs_num , = struct.unpack('<H', sheet_name[chars_start:chars_start+2])
                chars_start += 2
            if has_asian_phonetic:
                asian_phonetic_size , = struct.unpack('<I', sheet_name[chars_start:chars_start+4])
                chars_start += 4

            if compressed:
                chars_end = chars_start + chars_num
                result = unicode(sheet_name[chars_start:chars_end], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')
            else:
                chars_end = chars_start + 2*chars_num
                ints = struct.unpack('<'+'H'*chars_num, sheet_name[chars_start:chars_end])
                for i in ints:
                    result += unichr(i)
            
            tail_size = 4*runs_num + asian_phonetic_size
        else:
            result = unicode(sheet_name[1:], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')
        
        return result


    def process_LABEL(biff8, rec_data):
        row_idx, col_idx, xf_idx = struct.unpack('<3H', rec_data[:6])

        label_name = rec_data[6:]

        if biff8:
            chars_num, options = struct.unpack('<HB', label_name[:3])
            
            chars_start = 3
            runs_num = 0
            asian_phonetic_size = 0

            result = ''

            compressed = (options & 0x01) == 0
            has_asian_phonetic = (options & 0x04) != 0
            has_format_runs = (options & 0x08) != 0

            if has_format_runs:
                runs_num , = struct.unpack('<H', label_name[chars_start:chars_start+2])
                chars_start += 2
            if has_asian_phonetic:
                asian_phonetic_size , = struct.unpack('<I', label_name[chars_start:chars_start+4])
                chars_start += 4

            if compressed:
                chars_end = chars_start + chars_num
                result = unicode(label_name[chars_start:chars_end], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')
            else:
                chars_end = chars_start + 2*chars_num
                ints = struct.unpack('<'+'H'*chars_num, label_name[chars_start:chars_end])
                for i in ints:
                    result += unichr(i)
            
            tail_size = 4*runs_num + asian_phonetic_size
        else:
            result = unicode(label_name[2:], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')

        return (row_idx, col_idx, result)


    def process_LABELSST(rec_data):
        row_idx, col_idx, xf_idx, sst_idx = struct.unpack('<3HI', rec_data)
        return (row_idx, col_idx, sst_idx)


    def process_RSTRING(biff8, rec_data):
        if biff8:
            return process_LABEL(biff8, rec_data)
        else:
            row_idx, col_idx, xf_idx, length = struct.unpack('<4H', rec_data[:8])
            result = unicode(rec_data[8:8+length], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')

        return (row_idx, col_idx, result)
        

    def decode_rk(encoded):
        is_multed_100 = (encoded & 0x01) != 0
        is_integer = (encoded & 0x02) != 0

        if is_integer:
            result = float(encoded >> 2)
        else:
            ieee754 = struct.pack('<2I', 0, (encoded & 0xFFFFFFFC))
            result , = struct.unpack('<d', ieee754)
        if is_multed_100:
            result /= 100.0
        
        return result


    def process_RK(rec_data):
        row_idx, col_idx, xf_idx, encoded = struct.unpack('<3HI', rec_data)
        result = decode_rk(encoded)
        return (row_idx, col_idx, result)


    def process_MULRK(rec_data):
        row_idx, first_col_idx = struct.unpack('<2H', rec_data[:4])
        last_col_idx , = struct.unpack('<H', rec_data[-2:])
        xf_rk_num = last_col_idx - first_col_idx + 1

        results = []
        for i in range(xf_rk_num):
            xf_idx, encoded = struct.unpack('<HI', rec_data[4+6*i : 4+6*(i+1)])
            results.append(decode_rk(encoded))

        return zip([row_idx]*xf_rk_num, range(first_col_idx, last_col_idx+1), results)


    def process_NUMBER(rec_data):
        row_idx, col_idx, xf_idx, result = struct.unpack('<3Hd', rec_data)
        return (row_idx, col_idx, result)

    
    def process_SST(rec_data, sst_continues):
        # 0x00FC
        total_refs, total_str = struct.unpack('<2I', rec_data[:8])
        #print total_refs, str_num

        pos = 8
        curr_block = rec_data
        curr_block_num = -1
        curr_str_num = 0
        SST = {}

        while curr_str_num < total_str:
            if pos >= len(curr_block):
                curr_block_num += 1
                curr_block = sst_continues[curr_block_num]
                pos = 0

            chars_num, options = struct.unpack('<HB', curr_block[pos:pos+3])
            #print chars_num, options
            pos += 3

            asian_phonetic_size = 0
            runs_num = 0
            has_asian_phonetic = (options & 0x04) != 0
            has_format_runs = (options & 0x08) != 0
            if has_format_runs:
                runs_num , = struct.unpack('<H', curr_block[pos:pos+2])
                pos += 2
            if has_asian_phonetic:
                asian_phonetic_size , = struct.unpack('<I', curr_block[pos:pos+4])
                pos += 4

            curr_char = 0
            result = ''
            while curr_char < chars_num:
                if pos >= len(curr_block):
                    curr_block_num += 1
                    curr_block = sst_continues[curr_block_num]
                    options = ord(curr_block[0])
                    pos = 1
                #print curr_block_num

                compressed = (options & 0x01) == 0
                if compressed:
                    chars_end = pos + chars_num - curr_char
                else:
                    chars_end = pos + 2*chars_num - 2*curr_char
                #print compressed, has_asian_phonetic, has_format_runs

                splitted = chars_end > len(curr_block)
                if splitted:
                    chars_end = len(curr_block)
                    if compressed:
                        chars_in_block = len(curr_block) - pos
                    else:
                        chars_in_block = (len(curr_block) - pos) >> 1
                else:
                    chars_in_block = chars_num - curr_char
                #print splitted, chars_in_block, chars_num, curr_char, pos

                if compressed:
                    result += unicode(curr_block[pos:chars_end], UnicodeUtils.DEFAULT_ENCODING, 'backslashreplace')
                else:
                    ints = struct.unpack('<'+'H'*chars_in_block, curr_block[pos:chars_end])
                    for i in ints:
                        result += unichr(i)

                pos = chars_end
                curr_char = len(result)
            # end while

            # TODO: handle spanning format runs over CONTINUE blocks ???
            tail_size = 4*runs_num + asian_phonetic_size
            if len(curr_block) < pos + tail_size:
                pos = pos + tail_size - len(curr_block)
                curr_block_num += 1
                curr_block = sst_continues[curr_block_num]
            else:
                pos += tail_size

            #print result.encode('cp866', 'backslashreplace')

            SST[curr_str_num] = result
            curr_str_num += 1

        return SST


    #####################################################################################
    
    import struct

    biff8 = True
    SST = {}
    sheets = []
    sheet_names = []
    values = {}
    ws_num = 0
    BOFs = 0
    EOFs = 0

    # Inside MS Office document looks like filesystem
    # We need extract stream named 'Workbook' or 'Book'
    wb_bin_data = get_wb_stream_data(filename)
    wb_bin_data_len = len(wb_bin_data)
    stream_pos = 0
    
    # Excel's method of data storing is based on 
    # ancient technology "TLV" (Type, Length, Value).
    # In addition, if record size grows to some limit
    # Excel writes CONTINUE records
    while stream_pos < wb_bin_data_len and EOFs <= ws_num:
        rec_id, data_size = struct.unpack('<2H', wb_bin_data[stream_pos:stream_pos+4])
        stream_pos += 4
        
        rec_data = wb_bin_data[stream_pos:stream_pos+data_size]
        stream_pos += data_size

        if rec_id == 0x0809: # BOF
            #print 'BOF', 
            BOFs += 1
            ver, substream_type = struct.unpack('<2H', rec_data[:4])
            if substream_type == 0x0005:
                # workbook global substream
                biff8 = ver >= 0x0600
            elif substream_type == 0x0010:
                # worksheet substream
                pass
            #print 'BIFF8 == ', biff8
        elif rec_id == 0x000A: # EOF
            #print 'EOF'
            if BOFs > 1:
                sheets.extend([values])
                values = {}
            EOFs += 1
        elif rec_id == 0x0085: # BOUNDSHEET
            #print 'BOUNDSHEET',
            ws_num += 1
            b = process_BOUNDSHEET(biff8, rec_data)
            sheet_names.extend([b])
            #print b.encode('cp866', 'backslashreplace')
        elif rec_id == 0x00FC: # SST
            #print 'SST'
            sst_data = rec_data
            sst_continues = []
            rec_id, data_size = struct.unpack('<2H', wb_bin_data[stream_pos:stream_pos+4])
            while rec_id == 0x003C: # CONTINUE
                #print 'SST CONTINUE'
                stream_pos += 4
                rec_data = wb_bin_data[stream_pos:stream_pos+data_size]
                sst_continues.extend([rec_data])
                stream_pos += data_size
                rec_id, data_size = struct.unpack('<2H', wb_bin_data[stream_pos:stream_pos+4])
            SST = process_SST(sst_data, sst_continues)
        elif rec_id == 0x00FD: # LABELSST
            #print 'LABELSST',
            r, c, i = process_LABELSST(rec_data)
            values[(r, c)] = SST[i]
            #print r, c, SST[i].encode('cp866', 'backslashreplace')
        elif rec_id == 0x0204: # LABEL
            #print 'LABEL',
            r, c, b = process_LABEL(biff8, rec_data)
            values[(r, c)] = b
            #print r, c, b.encode('cp866', 'backslashreplace')
        elif rec_id == 0x00D6: # RSTRING
            #print 'RSTRING',
            r, c, b = process_RSTRING(biff8, rec_data)
            values[(r, c)] = b
            #print r, c, b.encode('cp866', 'backslashreplace')
        elif rec_id == 0x027E: # RK
            #print 'RK',
            r, c, b = process_RK(rec_data)
            values[(r, c)] = b
            #print r, c, b
        elif rec_id == 0x00BD: # MULRK
            #print 'MULRK',
            for r, c, b in process_MULRK(rec_data):
                values[(r, c)] = b
            #print r, c, b
        elif rec_id == 0x0203: # NUMBER
            #print 'NUMBER',
            r, c, b = process_NUMBER(rec_data)
            values[(r, c)] = b
            #print r, c, b

    return zip(sheet_names, sheets)
