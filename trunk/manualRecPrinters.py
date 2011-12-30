#!/usr/bin/python

import struct
from util import hexdump

#i hacked this together on a whim to add some extra support, it's not pretty

#
# an object record consists of tlv data in teh same format as a normal excel record
# 2 bytes type
# 2 bytes len
# 'len' bytes data
# the following dictionary contains the basic subtypes, there are more sub-subtypes

objSubTypeMap = {
        0x00:["ftEnd", "End of OBJ Record"],
        0x01:["(Reserved)", "(Reserved)"],
        0x02:["(Reserved)", "(Reserved)"],
        0x03:["(Reserved)", "(Reserved)"],
        0x04:["ftMacro", "Fmla-style macro"],
        0x05:["ftButton", "Command button"],
        0x06:["ftGmo", "Group Marker"],
        0x07:["ftCf", "Clipboard format"],
        0x08:["ftPioGrbit", "Picture option flags"],
        0x09:["ftPictFmla", "Picture fmla-style macro"],
        0x0a:["ftCbls", "Checkboxlink"],
        0x0b:["ftRbo", "Radio button"],
        0x0c:["ftSbs", "Scroll bar"],
        0x0d:["ftNts", "Note structure"],
        0x0e:["ftSbsFmla", "Scroll bar fmla-style macro"],
        0x0f:["ftGboData", "Group box data"],
        0x10:["ftEdoData", "Edit control data"],
        0x11:["ftRboData", "Radio button data"],
        0x12:["ftCblsData", "Check box data"],
        0x13:["ftLbsData", "List box data"],
        0x14:["ftCblsFmla", "Check box link fmla-style macro"],
        0x15:["ftCmo", "Common object data"],
        }
def objSubPrinter(data, dLen):

    off = 0
    while off <= dLen - 4:
        t, l = struct.unpack("<HH", data[off:off+4])
        try:
            desc = objSubTypeMap[t]
            print("        Subtype %s [%#x (%d)] offset %#x, len %#x (%d) (%s)" %
                    (desc[0], t, t, off, l, l, desc[1]))
        except KeyError:
            print "        obj.subtype %#x (%d) obj.sublen %#x (%d)" % (t, t, l, l)
        off += l + 4
    return -1

extraPrinters = {
        "[obj]":objSubPrinter,
}
