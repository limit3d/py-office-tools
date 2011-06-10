#!/usr/bin/python

import OleFileIO_PL
import sys
import dumpXL
import dumpPPT
from dumpProps import dumpPropSetStream
import getopt


#
def usage(prog):
    print("Usage: %s [ -f OLE file ] [ -h help ] [ -d debug ] [ -x extract stream (use -o for output file)]"
            "[ -o output file ] [ -O offset into stream (extraction) ]" %
            (prog))
    sys.exit(1)

#
if __name__ == "__main__":

    streamOffset = 0
    outputFile = fName = extractStream = None

    try:
        opts, args = getopt.getopt(sys.argv[1:], "hf:dx:o:O:")
    except getopt.GetoptError, err:
        print str(err)
        usage(sys.argv[0])

    for o, a in opts:
        if o == "-h":
            usage(sys.argv[0])
        if o == "-O":
            streamOffset = int(a)
        elif o == "-x":
            extractStream = a
        elif o == "-f":
            fName = a
        elif o == "-o":
            outputFile = a
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
    if extractStream is not None:
        if outputFile is None:
            print "ERROR: extracting a stream requires output file [ -o  file ]"
            sys.exit(1)
        
        stream = ole.openstream(extractStream)
        buf = stream.read()
        outFile = open(outputFile, "wb")
        outFile.write(buf[streamOffset:])
        outFile.close()

        print "Extracted stream %s (%d of %d available bytes) to file %s" % \
                    (extractStream, len(buf) - streamOffset, len(buf), outputFile)
        sys.exit(0)

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
