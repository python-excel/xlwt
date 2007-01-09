from xlwt import *
import sys
from struct import pack, unpack

def cellname(rowx, colx):
    # quick kludge, works up to 26 cols :-)
    return chr(ord('A') + colx) + str(rowx + 1)

def RK_pack_check(num, anint, case=None):
    if not(-0x7fffffff - 1 <= anint <= 0x7fffffff):
        print "RK_pack_check: not a signed 32-bit int: %r (%r); case: %r" \
            % (anint, hex(anint), case)
    pstr = pack("<i", anint)
    actual = unpack_RK(pstr)
    if actual != num:
        print "RK_pack_check: round trip failure: %r (%r); case %r;  %r in, %r out" \
            % (anint, hex(anint), case, num, actual)
 

def RK_encode(num, blah=0):
    """\
    Return a 4-byte RK encoding of the input float value
    if possible, else return None.
    """
    rk_encoded = 0
    packed = pack('<d', num)

    if blah: print
    if blah: print repr(num)
    w01, w23 = unpack('<2i', packed)
    if not w01 and not(w23 & 3):
        # 34 lsb are 0
        if blah: print "float RK", w23, hex(w23)
        return RK_pack_check(num, w23, 0)
        # return RKRecord(
        #    self.__parent.get_index(), self.__idx, self.__xf_idx, w23).get()

    if -0x20000000 <= num < 0x20000000:
        inum = int(num)
        if inum == num:
            if blah: print "30-bit integer RK", inum, hex(inum)
            rk_encoded = 2 | (inum << 2)
            if blah: print "rk", rk_encoded, hex(rk_encoded)
            return RK_pack_check(num, rk_encoded, 2)
            # return RKRecord(
            #     self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

    temp = num * 100
    packed100 = pack('<d', temp)
    w01, w23 = unpack('<2i', packed100)
    if not w01 and not(w23 & 3):
        # 34 lsb are 0
        if blah: print "float RK*100", w23, hex(w23)
        return RK_pack_check(num, w23 | 1, 1)
        # return RKRecord(
        #    self.__parent.get_index(), self.__idx, self.__xf_idx, w23 | 1).get()

    if -0x20000000 <= temp < 0x20000000:
        itemp = int(round(temp, 0))
        if blah: print (itemp == temp), (itemp / 100.0 == num)
        if itemp / 100.0 == num:
            if blah: print "30-bit integer RK*100", itemp, hex(itemp)
            rk_encoded = 3 | (itemp << 2)
            return RK_pack_check(num, rk_encoded, 3)
            # return RKRecord(
            #    self.__parent.get_index(), self.__idx, self.__xf_idx, rk_encoded).get()

    if blah: print "Number" 
    # return NumberRecord(
    #    self.__parent.get_index(), self.__idx, self.__xf_idx, num).get()

def unpack_RK(rk_str):
    flags = ord(rk_str[0])
    if flags & 2:
        # There's a SIGNED 30-bit integer in there!
        i,  = unpack('<i', rk_str)
        i >>= 2 # div by 4 to drop the 2 flag bits
        if flags & 1:
            return i / 100.0
        return float(i)
    else:
        # It's the most significant 30 bits of an IEEE 754 64-bit FP number
        d, = unpack('<d', '\0\0\0\0' + chr(flags & 252) + rk_str[1:4])
        if flags & 1:
            return d / 100.0
        return d

testvals = (
    130.63999999999999,
    130.64,
    -18.649999999999999,
    -18.65,
    137.19999999999999,
    137.20,
    -16.079999999999998,
    -16.08,
    0,
    1,
    2,
    3,
    0x1fffffff,
    0x20000000,
    0x20000001,
    1000999999,
    0x3fffffff,
    0x40000000,
    0x40000001,
    0x7fffffff,
    0x80000000,
    0x80000001,
    0xffffffff,
    0x100000000,
    0x100000001,
    )

XLS = 1
BLAH = 1

def main(do_neg):
    if XLS:
        w = Workbook()
        ws = w.add_sheet('Test RK encoding')
        for colx, heading in enumerate(('actual', 'expected', 'OK') * 2):
            ws.write(0, colx, heading)
    rx = 0
    for neg in range(do_neg + 1):
        for seed in testvals:
            rx += 1
            for i in range(2):
                bv = [seed, seed /100.00][i] * (1 - 2 * neg)
                bv = float(bv) # pyExcelerator cracks it with longs
                cx = i * 3
                if XLS:
                    ws.write(rx, cx, bv)
                    ws.write(rx, cx + 1, repr(bv))
                    ws.write(rx, cx + 2, Formula(
                        '%s=VALUE(%s)' % (cellname(rx, cx), cellname(rx, cx + 1))
                        ))
                else:
                    RK_encode(bv, blah=BLAH)
    if XLS:                        
        w.save('RKbug%d.xls' % do_neg)

if __name__ == "__main__":
    # arg == 0: only positive test values used
    # arg == 1: both positive & negative test values used
    main(int(sys.argv[1]))
