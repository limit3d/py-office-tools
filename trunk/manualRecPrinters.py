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

#
def optPrinter(data, dLen, propCount, depth):
    
    cProps = []
    off = 0
    
    #first we have an array of property descriptors packed together
    while off <= dLen - 6 and propCount > 0:
        try:
            pff, op = struct.unpack("<HL", data[off:off+6])
        except struct.error:
            return

        pid = pff & 0x3fff
        fBid = (pff & 0x4000) >> 14
        fComplex = (pff & 0x8000) >> 15

        #op contains the length
        if fComplex:
            cProps.append((op, pid))

        print(" "*(depth*4) + "Prop ID %#x (%d), Blip ID %#x (%d), Complex %#x, Op %#x (%#d)" %
                        (pid, pid, fBid, fBid, fComplex, op, op))
        
        off += 6
        propCount -= 1

    #dump the data for any complex properties
    if off < dLen:

        print("\n" + " "*(depth*4) + "Dumping data for %d complex props\n" % len(cProps))

        for (cLen, pid) in cProps:

            if off + cLen > dLen:
                print "Warning: breaking early, not enough to print remaining props"
                break
            
            lenLeft = min(cLen, dLen - off)
            print(" "*(depth*4) + "Dumping property %#x (%d)...\n" % (pid, pid))
            print hexdump(data[off:off+lenLeft], indent=(depth)*4)

            off += cLen

#
msoDrawTypeMap = {
        0xf002:["msofbtDgContainer", "per sheet/page/slide data", None],
        0xf008:["msofbtDg", "an FDG", None],
        0xf118:["msofbtRegroupItems", "several FRITs", None],
        0xf120:["msofbtColorScheme", "the colors of the source host's color scheme", None],
        0xf003:["msofbtSpgrContainer", "several SpContainers, the first of which is the group shape itself", None],
        0xf004:["msofbtSpContainer", "a shape", None],
        0xf009:["msofbtSpgr", "an FSPGR", None],
        0xf00a:["msofbtSp", "an FSP", None],
        0xf00b:["msofbtOPT", "a shape property table", optPrinter],
        0xf121:["msofbtSecondaryOPT", "a shape property table", optPrinter],
        0xf122:["msofbtTertiaryOPT", "a shape property table", optPrinter],
        0xf00c:["msofbtTextbox", "RTF Text", None],
        0xf00d:["msofbtClientTextbox", "the host defined text in textbox", None],
        0xf00e:["msofbtAnchor", "a RECT, in 100000ths of an inch", None],
        0xf00f:["msofbtChildAnchor", "a RECT, in parent relative units", None],
        0xf010:["msofbtClientAnchor", "shape location, host defined format", None],
        0xf011:["msofbtClientData", "host specific data", None],
        0xf11f:["msofbtOleObject", "a serialized IStorage for an object", None],
        0xf11d:["msofbtDeletedPspl", "an FPSPL", None],
        0xf005:["msofbtSolverContainer", "the rules governing shapes", None],
        0xf012:["msofbtConnectorRule", "an FConnector Rule", None],
        0xf013:["msofbtAlignRule", "an FAlignRule", None],
        0xf014:["msofbtArcRule", "an FArcRule", None],
        0xf015:["msofbtClientRule", "host defined rule", None],
        0xf017:["msofbtCalloutRule", "an FCORU", None],
}

#
def msoDrawPrinter(data, dLen, depth=2):
    off = 0
    while off <= dLen - 8:
        try:
            vif, l = struct.unpack("<LL", data[off:off+8])
        except struct.error:
            return

        ver = vif & 0xf
        inst = (vif & 0xfff0) >> 4
        fbt = (vif & 0xffff0000) >> 16

        try:
            desc = msoDrawTypeMap[fbt]
            print(" "*(depth*4) + "Subtype %s [%#x (%d)] ver %#x inst %#x offset %#x, len %#x (%d) (%s)" %
                    (desc[0], fbt, fbt, ver, inst, off, l, l, desc[1]))
            if ver == 0xf:
                msoDrawPrinter(data[off+8:off+8+l], l, depth + 1)
            else:

                #CONTINUE records can make the actual length of the property span across
                #multiple records, which we don't currently support, so we chop it at what
                #we have left in the current record
                lenLeft = min(l, dLen - off - 8)
                if lenLeft != l:
                    print "!!Warning, data chopped short, wanted %d, got %d, CONTINUE rec?" % \
                            (l, lenLeft)
                
                #print out sub-record data with custom printer if available
                if desc[2]:
                    desc[2](data[off+8:off+8+lenLeft], lenLeft, inst, depth + 1)
                else:
                    print hexdump(data[off+8:off+8+lenLeft], indent=(depth + 1)*4),
        except KeyError:
            print " "*(depth*4) + "obj.subtype %#x (%d) ver %#x inst %#x obj.sublen %#x (%d)" % \
                        (fbt, fbt, ver, inst, l, l)
            print hexdump(data[off+8:min(off+8+l, dLen - 8)], indent=(depth + 1)*4),
        off += l + 8
    return -1

#
extraPrinters = {
        "[obj]":objSubPrinter,
        "[msodrawing]":msoDrawPrinter,
}
