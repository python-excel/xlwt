# -*- coding: windows-1251 -*-

from BIFFRecords import BiffRecord
from struct import *
import hashlib
import os

'''
Makes it possible to insert a header or footer picture into a sheet.
Currently only works for a single picture in the entire sheet but some
extensibility should be possible.

Usage: wokesheet_object.add_hfpicture(filename, page_position, h_size, v_size, name)

page_position is described with a 2 character string. The first letter, 'R',
'C' or 'L', sets the horizontal position of the picture within the header/footer
while the second letter, 'H' or 'F' determines whether the picture is located in
the header or footer. Note that the corresponding header or footer strings must have
'&G' in them in order for this to work!

All sizes are in pixels. Name is an optional argument, the rest are mandatory.
'''

#going to pack in four variable or 16 byte groups
#this generates the raw data for the first part of the worksheet record.
#It gets rewritten after every use of add_hfpicture.
def generate_ws_header_data(pi_nshapes, ps_name, pi_rec_size):
    ls_data = ''
    #initial record header and the following zeroes
    ls_data += pack('<HLLH',0x0866, 0,0,0)

    #single drawing tag (0x0001) dgcontainer headers and record length
    li_rl1 = 0x98 + len(ps_name)*2 + pi_rec_size
    ls_data += pack('<HHHL',0x0001,0x000f,0xf002,li_rl1)

    #beginning of fdg record. 8 is the length. 
    ls_data += pack('<HHL',0x0010,0xf008,0x00000008)

    #continue fdg record. Some future problems will likely occur here
    ls_data += pack('<BBHBBH',0x00,pi_nshapes+1, 0x0000,0x00,pi_nshapes, 0x0004)

    #beginning of spgrcontainer (namely, the frtrecord).
    #The standard format is recinstance followed by record type and,
    #finally, record length. Common throughout the module.
    li_rl2 = 0x80+len(ps_name)*2 + pi_rec_size
    ls_data += pack('<HHL',0x000f,0xf003,li_rl2)

    #beginning of spcontainer. 28h is the length
    ls_data += pack('<HHL', 0x000f, 0xf004, 0x00000028)

    #spgr record. Full of zeros, essentially, 10h is the length
    ls_data += pack('<HHLLLLL', 0x0001, 0xf009, 0x00000010,0,0,0,0)

    #beginning of fsp record
    ls_data += pack('<HHL', 0x0002, 0xf00a, 0x00000008)

    #continue fsp record. Variable stuff should be here.
    ls_data += pack('<HHHH',0x0400,0x0000,0x0005,0x0000)
    
    return ls_data

#this generates the second part of the worksheet data. Contains shapes information
#and a new record is appended for every use of add_hfpicture
def generate_ws_rec_data(pi_nshapes, ps_pos, pi_hsize, pi_vsize, ps_name):
    ls_data = ''
    #spcontainer record.
    li_rl0 = 0x48 + len(ps_name)*2
    ls_data += pack('<HHL',0x000f,0xf004,li_rl0)

    #beginning of fsp record. 
    ls_data += pack('<HHL', 0x04b2, 0xf00a, 0x00000008)

    #continue fsp record. Variable stuff here
    ls_data += pack('<BBHHH', pi_nshapes,0x04,0x0000,0x0a00,0x0000)

    #beginning of FOPT record
    li_rl1 = 0x20+2*len(ps_name)
    ls_data += pack('<HHL', 0x0043, 0xf00b, li_rl1)

    #contunue FOPT record. Variable stuff follows...
    ls_data += pack('<HHHH', 0x007f, 0x0100, 0x0100, 0x4104)

    #continue FOPT record.
    li_rl2 = 2 + len(ps_name)*2
    ls_data += pack('<BBHHL', pi_nshapes, 0x00, 0x0000, 0xc105, li_rl2)

    #continue FOPT record.
    ls_data += pack('<HHH', 0xc380, 0x0006, 0x0000)

    #insert picture name here. First 2 bytes are byte order mark.
    ls_data += ps_name.encode('utf-16')[2:]

    #some random but obviously important zeros
    ls_data += pack('<H',0x0000)

    #picture position. Again, ignore the byte order mark
    ls_data += ps_pos.upper().encode('utf-16')[2:]

    #another set of important zeros
    ls_data += pack('<H', 0x0000)

    #OfficeArtClientAnchor record
    ls_data += pack('<HHLLL', 0x0000, 0xf010, 0x00000008, pi_hsize, pi_vsize)

    return ls_data

#this generates the raw data for the first part of the workbook record.
#It gets rewritten after every use of add_hfpicture.
def generate_wb_header_data(pi_ntot, pi_tot_size):
    ls_data = ''

    #initial record header and the following zeroes
    ls_data += pack('<HLLH',0x0866, 0,0,0)

    #drawing group tag (0x0002) dgcontainer headers and record length
    #size offset generation. Right now this only works for 1 or 2 pictures on
    #a single sheet
    li_offset1 = 0x85+0x45*(pi_ntot-1)
    ls_data += pack('<HHHL',0x0002,0x000f,0xf000,pi_tot_size+li_offset1)

    #beginning of FDGGBlock. Record length and max id can be variable here if
    #there are hfpictures on multiple sheets (currently not implemented).
    ls_data += pack('<HHL',0x0000,0xf006,0x00000018)

    #continue FDGGBlock. Still some variables possible
    ls_data += pack('<BBHL',pi_ntot+1,0x04,0x0000,0x00000002)

    #continue FDGGBlock.
    ls_data += pack('<HHLLL',pi_ntot+1, 0x0000,0x00000001,0x00000001,pi_ntot+1)

    #BlipStoreContainer header. The rec part follows. Another case of possibly
    #variable offset for hfpictures on multiple sheets. The rest is contained
    #in the rec part.
    li_offset2 = 0x45+0x45*(pi_ntot-1)
    ls_data += pack('<HHL', 0x10*pi_ntot+0xf, 0xf001, pi_tot_size+li_offset2)

    return ls_data

#this generates the second part of the workbook data.
#and a new record is appended for every use of add_hfpicture
def generate_wb_rec_data(ps_filename, pi_ntot):
    ls_data = ''
    li_file_size = os.path.getsize(ps_filename)
    lh_file = open(ps_filename,mode='rb')
    ls_bitmap_data = lh_file.read()
    lh_file.close()

    #beginning of the bstorefilecontainerblock record. Size offset may be variable
    #if multiple sheets version is to be implemented.
    ls_data += pack('<HHLH', 0x0052, 0xf007, li_file_size+0x3d, 0x0505)

    #md4
    lo_md4 = hashlib.new('md4')
    lo_md4.update(ls_bitmap_data)
    ls_md4_data = lo_md4.digest()
    ls_data += ls_md4_data

    #continue bstorefilecontainerblock. Lots of padding zeros
    ls_data += pack('<HLLLL',0x00ff, li_file_size+0x19, 0x00000001,0,0)

    #begin OfficeArtBlibJpeg
    ls_data += pack('<HHL', 0x46a0, 0xf01d, li_file_size+0x11)

    #another md4 followed by ff byte and the bitmap data itself
    ls_data += ls_md4_data + pack('<B',0xFF) + ls_bitmap_data

    return ls_data

#last 24 bytes of the workbook record following the bitmap data, according to
#the record size. Currently completely static
def generate_wb_footer_data():
    ls_data = ''

    #this is just the default recently used colord record
    ls_data += pack('<HHHH',0x0040,0xf1e1,0x0010,0x0000)
    ls_data += pack('<HHHH',0xffff,0x0000,0x0000,0x00ff)
    ls_data += pack('<HHHH', 0x8080, 0x0080, 0x00f7, 0x1000)

    return ls_data

#since continue rules are not quite the same as the those for the standard
#BiffRecord, we cannot just subclass BiffRecord and have to make our own procedure
#for data consolidation. Instead, copying and modifying some code from it.
def consolidate_record(ps_header_data, ps_rec_data, **args):
    ls_data = ps_header_data + ps_rec_data
    if args.has_key('wb_footer') and args['wb_footer'] == True:
        ls_data += generate_wb_footer_data()
    ls_result = ''
    if len(ls_data)==0: return ls_result

    if len(ls_data) > 0x2020: # limit for BIFF7/8
        ll_chunks = []
        li_pos = 0
        while li_pos < len(ls_data):
            if li_pos == 0:
                li_chunk_pos = li_pos + 0x2020
            else:
                li_chunk_pos = li_pos + 0x2012 #accounting for continue bytes here
            ls_chunk = ls_data[li_pos:li_chunk_pos]
            ll_chunks.append(ls_chunk)
            li_pos = li_chunk_pos
        ls_continues = pack('<2H', 0x0866, len(ll_chunks[0])) + ll_chunks[0]
        for ls_chunk in ll_chunks[1:]:
            #the main difference from the standard BiffRecord is here. Don't
            #forget to add 0xee bytes to the length to account for the next 2 records
            ls_continues += pack('<HHHL',0x0866, len(ls_chunk)+0xe,0x0866,0x00000000)
            #remaining zero, continue byte
            ls_continues += pack('<LHH',0x00000000,0x0000,0x0006) + ls_chunk
        ls_result = ls_continues
    else:
        ls_rec_header = pack('<HH', 0x0866, len(ls_data))
        ls_result = ls_rec_header + ls_data

    return ls_result
