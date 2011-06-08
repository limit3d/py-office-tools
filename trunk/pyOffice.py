#!/usr/bin/python

import OleFileIO_PL
import sys
import dumpXL
import dumpPPT
from dumpProps import dumpPropSetStream
import getopt


#
def usage(prog):
    print "Usage: %s [ -f OLE file ] [ -h help ] [ -d debug ]" % \
            (prog)
    sys.exit(1)

#
if __name__ == "__main__":

    fName = None

    try:
        opts, args = getopt.getopt(sys.argv[1:], "hf:d")
    except getopt.GetoptError, err:
        print str(err)
        usage(sys.argv[0])

    for o, a in opts:
        if o == "-h":
            usage(sys.argv[0])
        elif o == "-f":
            fName = a
        elif o == "-d":
            OleFileIO_PL.set_debug_mode(True)
        else:
            usage(sys.argv[0])

    if not fName:
        usage(sys.argv[0])

    # Test if a file is an OLE container:
    if not OleFileIO_PL.isOleFile(fName):
        print "File %s is not an OLE file" % (fName)
        sys.exit(1)

    print "[*]Opening file %s" % (fName)

    # Open OLE file:
    ole = OleFileIO_PL.OleFileIO(fName)

    # Get list of storages/streams
    objs = ole.listdir()
    print "[*]Listing streams/storages:\n"
    ole.dumpdirectory()

    #
    if ole.exists("Workbook"):
        print "\n[**]Detected Excel file %s" % fName
        dumpXL.dump(ole)
    
    #
    if ole.exists("PowerPoint Document"):
        print "\n[**]Detected PowerPoint file %s" % fName
        dumpPPT.dump(ole)

    #
    if ole.exists("\x05SummaryInformation"):
        dumpPropSetStream(ole, "\x05SummaryInformation")
    
    #
    if ole.exists("\x05DocumentSummaryInformation"):
        dumpPropSetStream(ole, "\x05DocumentSummaryInformation")
