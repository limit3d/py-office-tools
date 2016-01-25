#Brief introduction to py-office-tools

# Overview #

This is a set of tools to support reading (and eventually writing) of the Excel and PowerPoint binary file formats. It relies on the OleFileIO\_PL module to deal with the underlying CBFF format. This tool focuses on the storages/streams that contain the application level data for Excel and PowerPoint.

## What it does ##

Prints out detailed information about Excel and PowerPoint records ("Workbook" and "PowerPoint Document" streams), as well as property sets in the DocumentSummary and DocumentSummaryInformation streams.

## How it does ##

I took the Excel and PowerPoint specs, converted them to HTML, fixed up lots of broken ass stuff, and then used those as input to a code generator that spit out Python structures describing the records.

## If you don't like my code (likely) or python ##

Take the HTML files and use them to generate code in a language of your choice.

## Usage + output ##
<pre>
./pyOffice.py -f somefile.xls<br>
<br>
[*]Opening file /home/root/tmp/a3730uez.xls<br>
[*]Listing streams/storages:<br>
<br>
'Root Entry' (root) 0 bytes<br>
{00020820-0000-0000-C000-000000000046}<br>
'\x05DocumentSummaryInformation' (stream) 4096 bytes<br>
'\x05SummaryInformation' (stream) 4096 bytes<br>
'Workbook' (stream) 15119 bytes<br>
<br>
[**]Detected Excel file /home/root/tmp/a3730uez.xls<br>
********************************************************************************<br>
[*]Dumping Workbook stream 0x3b0f (15119) bytes...<br>
<br>
[ii]BOF record: current count 1<br>
[0]Record BOF [0x809 (2057)] offset 0x0 (0), len 0x10 (16) (Beginning of File)<br>
WORD vers = 0x600 (1536)<br>
WORD dt = 0x5 (5)<br>
WORD rupBuild = 0x1fe9 (8169)<br>
WORD rupYear = 0x7cd (1997)<br>
DWORD bfh = 0xc0c9 (49353)<br>
DWORD sfo = 0x306 (774)<br>
[1]Record INTERFACEHDR [0xe1 (225)] offset 0x14 (20), len 0x2 (2) (Beginning of User Interface Records)<br>
WORD Cv = 0x4b0 (1200)<br>
[2]Record MMS [0xc1 (193)] offset 0x1a (26), len 0x2 (2) (ADDMENU/DELMENU Record Group Count)<br>
BYTE caitm = 0x0 (0)<br>
BYTE cditm = 0x0 (0)<br>
[3]Record INTERFACEEND [0xe2 (226)] offset 0x20 (32), len 0x0 (0) (End of User Interface Records)<br>
[4]Record WRITEACCESS [0x5c (92)] offset 0x24 (36), len 0x70 (112) (Write Access User Name)<br>
Field stName is 0x70 (112) bytes, dumping:<br>
0000000000   09 00 00 47 69 6E 61 20 50 61 70 70 61 6C 67 6F    ...Gina Pappalgo<br>
0000000010   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
0000000020   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
0000000030   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
0000000040   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
0000000050   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
0000000060   20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20<br>
<br>
---<br>
<br>
./pyOffice.py -f someFile.ppt<br>
<br>
[*]Opening file /home/root/tmp/magicmktwhitepaperintro.ppt<br>
[*]Listing streams/storages:<br>
<br>
'Root Entry' (root) 3008 bytes<br>
{64818D10-4F9B-11CF-86EA-00AA00B929E8}<br>
'\x05DocumentSummaryInformation' (stream) 2924 bytes<br>
'\x05SummaryInformation' (stream) 21124 bytes<br>
'Current User' (stream) 50 bytes<br>
'Pictures' (stream) 11402 bytes<br>
'PowerPoint Document' (stream) 106945 bytes<br>
<br>
[**]Detected PowerPoint file /home/root/tmp/magicmktwhitepaperintro.ppt<br>
********************************************************************************<br>
[*]Dumping 'Current User' stream 0x32 (50) bytes...<br>
<br>
{Atom(1) CurrentUserAtom (0xff6, 4086) size 0x2a (42), instance 0x0 version 0x0 [offset 0x0 (0)]<br>
DWORD size 0x14 (20)<br>
DWORD headerToken 0xe391c05f (3817979999)<br>
DWORD offsetToCurrentEdit 0x1a19d (106909)<br>
WORD lenUserName 0x12 (18)<br>
WORD docFileVersion 0x3f4 (1012)<br>
BYTE majorVersion 0x3 (3)<br>
BYTE minorVersion 0x0 (0)<br>
WORD unused 0xffff (65535)<br>
ASCIIZ AnsiUserName 'Preferred Customer'<br>
DWORD relVersion 0x8 (8)<br>
UNICODE uniUserName:<br>
<br>
<br>
********************************************************************************<br>
[*]Dumping 'PowerPoint Document' stream 0x1a1c1 (106945) bytes...<br>
<br>
[Container(1) Document (0x3e8, 1000) size 0x6d27 (27943), instance 0x0 version 0xf [offset 0x0 (0)]<br>
{{Atom(2) DocumentAtom (0x3e9, 1001) size 0x28 (40), instance 0x0 version 0x1 [offset 0x8 (8)]<br>
GPointAtom.DWORD slideSize = 0x1680 (5760)<br>
GPointAtom.DWORD slideSize.1 = 0x10e0 (4320)<br>
GPointAtom.DWORD notesSize = 0x113c (4412)<br>
GPointAtom.DWORD notesSize.1 = 0x16b2 (5810)<br>
GRatioAtom.DWORD serverZoom = 0x5 (5)<br>
GRatioAtom.DWORD serverZoom.1 = 0xa (10)<br>
DWORD notesMasterPersist = 0x2 (2)<br>
DWORD handoutMasterPersist = 0x0 (0)<br>
WORD firstSlideNum = 0x1 (1)<br>
WORD slideSizeType = 0x0 (0)<br>
BYTE saveWithFonts = 0x0 (0)<br>
BYTE omitTitlePlace = 0x0 (0)<br>
BYTE rightToLeft = 0x0 (0)<br>
BYTE showComments = 0x0 (0)<br>
[[Container(3) ExObjList (0x409, 1033) size 0x568 (1384), instance 0x0 version 0xf [offset 0x38 (56)]<br>
{{{Atom(4) ExObjListAtom (0x40a, 1034) size 0x4 (4), instance 0x0 version 0x0 [offset 0x40 (64)]<br>
DWORD objectIdSeed = 0xa7 (167)<br>
[[[Container(5) ExEmbed (0xfcc, 4044) size 0xa4 (164), instance 0x0 version 0xf [offset 0x4c (76)]<br>
{{{{Atom(6) ExEmbedAtom (0xfcd, 4045) size 0x8 (8), instance 0x0 version 0x0 [offset 0x54 (84)]<br>
DWORD followColorScheme = 0x0 (0)<br>
BYTE cantLockServerB = 0x1 (1)<br>
BYTE noSizeToServerB = 0x0 (0)<br>
BYTE isTable = 0x0 (0)<br>
<br>
</pre>