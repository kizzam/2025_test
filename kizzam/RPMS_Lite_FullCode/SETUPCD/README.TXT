======================================================================
  Microsoft(R) Product Support Services Application Note (Text File)
                  WW1000: VSHARE.386 FOR WINDOWS(TM)
======================================================================
                                                   Revision Date: 2/94
                                                       1 Disk Included

The following information applies to Microsoft Windows, version 3.1.

-----------------------------------------------------------------------
| INFORMATION PROVIDED IN THIS DOCUMENT AND ANY SOFTWARE THAT MAY     |
| ACCOMPANY THIS DOCUMENT (collectively referred to as an Application |
| Note) IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER      |
| EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED      |
| WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR       |
| PURPOSE. The user assumes the entire risk as to the accuracy and    |
| the use of this Application Note. This Application Note may be      |
| copied and distributed subject to the following conditions:  1) All |
| text must be copied without modification and all pages must be      |
| included;  2) If software is included, all files on the disk(s)     |
| must be copied without modification (the MS-DOS(R)  utility         |
| diskcopy is appropriate for this purpose);  3) All components of    |
| this Application Note must be distributed together;  and  4) This   |
| Application Note may not be distributed for profit.                 |
|                                                                     |
| Copyright (C) 1994 Microsoft Corporation.  All Rights Reserved.     |
| Microsoft and MS-DOS are registered trademarks and Windows is a     |
| trademark of Microsoft Corporation.                                 |
|---------------------------------------------------------------------|

Introduction
------------

The WW1000 disk included with this Application Note contains
VSHARE.386 version 3.11.0.401. This version of VSHARE.386 is
compatible with Windows 3.1, Windows for Workgroups 3.1, and Windows
for Workgroups 3.11.

NOTE: Microsoft Word version 6.0a for Windows and Microsoft Office
Manager version 4.2 install the same version of VSHARE.386 included
with this Application Note (3.11.0.401).

The only difference between the version of VSHARE.386 included with
this Application Note and the one supplied with Windows for Workgroups
3.11 (version 3.11.0.300) is the added support for Windows 3.1.

Using VSHARE.386
----------------

If you are using applications that support object linking and
embedding (OLE) 2.0, you must run either SHARE.EXE or VSHARE.386.
VSHARE.386 eliminates the need for SHARE.EXE when you run Windows 3.1
or Windows for Workgroups in 386 enhanced mode. If you run Windows 3.1
in standard mode, you still need to run SHARE.EXE. If you run
applications that are not compatible with SHARE.EXE, and you run
Windows 3.1 in 386 enhanced mode, you may be able to use VSHARE.386
instead of SHARE.EXE.

To install the updated VSHARE.386 file
--------------------------------------

If you are running Windows for Workgroups 3.1 or 3.11 in 386 enhanced
mode, you are already using VSHARE.386. If you are running Windows
3.1, use the following procedure to install VSHARE.386 version
3.11.0.401:

1. Quit Windows 3.1.

2. Insert the enclosed WW1000 disk in the appropriate floppy disk
   drive.

3. To copy the file to the Windows SYSTEM subdirectory, type the
   following at the MS-DOS command prompt and then press ENTER

      copy <drive>:\vshare.386 <destination>

   where <drive> is the floppy disk drive containing the WW1000 disk
   and <destination> is the drive and directory of your Windows SYSTEM
   subdirectory.

   For example, if the WW1000 disk is in drive A and your Windows
   directory is called WINDOWS and is on drive C, type the following
   at the MS-DOS command prompt and then press ENTER:

      copy a:\vshare.386 c:\windows\system

4. Using a text editor (such as MS-DOS Editor), edit your AUTOEXEC.BAT
   file and remove the command for SHARE.EXE.

5. Edit the [386Enh] section in your SYSTEM.INI file to add the
   following line:

      device=vshare.386
