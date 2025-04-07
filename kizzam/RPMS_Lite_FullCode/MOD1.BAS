Option Explicit

' CreateDatabase and CompactDatabase Language constants
Global Const DB_LANG_GENERAL = ";LANGID=0x0809;CP=1252;COUNTRY=0"

' CreateDatabase and CompactDatabase options
Global Const DB_VERSION10 = 1        ' Microsoft Access Version 1.0
Global Const DB_ENCRYPT = 2          ' Make database encrypted.
Global Const DB_DECRYPT = 4          ' Decrypt database while compacting.
Global Const KEY_F12 = &H7B

Global GLFont As String
Global GLLogo As String

Global DBFileName As String
Global DBFullFileName As String
Global DBPath As String
Global DBFile As Database

Global WinPath As String
Global FileNameini As String

Global GLOwner As String
Global GLOwnerClub As String
Global GLOwner1 As String
Global GLAddress1 As String
Global GLAddress2 As String
Global GLCity As String
Global GLState As String
Global GLPostcode As String
Global GLLoft As String
Global GLPhoneNo As String
Global GLSerialNo As String
Global GLmark As String
Global GLFed As String
Global GLbscmonth As Integer
Global GLRSLpp As Integer
Global GLRSRaceFormType As Integer
Global GLDBVersion As String
Global GLAPPversion As String
Global GLLoftLat As Double
Global GLLoftLng As Double
Global GLLastBU As Variant
Global GLDistType As String
Global GLracepoint As String

'GLLvars Control Load of Screens
Global GLLfmbird As Integer
Global GLLfmerror As Integer

Global fm As String
Global GLBackup As String
Global GLSystem As String

Global GLDemo As Integer

' Colors
Global Const vbMAGENTA = &HFF00FF
Global Const vbRED = &HFF&
Global Const vbGREEN = &HFF00&
Global Const vbBLACK = &H0&
Global Const vbBLUE = &HFF0000
Global Const vbWHITE = &HFFFFFF
Global Const vbYELLOW = &HFFFF&

Type OrientStructure
      Orientation As Long
      Pad As String * 16
End Type

' Enter the following Declare statement on one, single line:
Declare Function Escape% Lib "GDI" (ByVal hDc%, ByVal nEsc%, ByVal nLen%, lpData As OrientStructure, lpOut As Any)


Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%, ByVal lpFileName$)
Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer)
Declare Function GetProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%)

Sub CalcDist ()
Dim DB As Database
Dim tblPoints As Table
Dim tblMembers As Table

Dim RecCnt As Integer

'Calc fileds
Dim RLatDeg As Double
Dim RLngDeg As Double
Dim MLatDeg As Double
Dim MLngDeg As Double
Dim rLatRad As Double
Dim rLngRad As Double
Dim mLatRad As Double
Dim mLngRad As Double
Dim DistGCM As Double
Dim DistRLM As Double
Dim CalcA As Double
Dim CalcB As Double
Dim CalcC As Double
Dim CalcD As Double
Dim CalcE As Double
Dim CalcF As Double
Dim reply As Integer
Dim WGS84_f As Double
Dim WGS84_eps As Double
Dim DistVincenty As Double
Dim DistGCM_UK As Double
Dim DistGCM_Intl As Double
Dim DistGCM_TDPF As Double
Dim NM As Double

On Error GoTo DB_Err

'Print the Pedigree
screen.MousePointer = 11

WGS84_f = 298.257223563
WGS84_eps = 6378137

Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
Set tblMembers = DB.OpenTable("Loft")    'Open Table
tblMembers.MoveLast
RecCnt% = tblMembers.RecordCount

'Can't print distance if no race points
If RecCnt% < 1 Then
    MsgBox "No Members on file", 16, "CANNOT PRINT DISTANCES"
    tblMembers.Close
    DB.Close
    Set DB = Nothing
    Set tblMembers = Nothing
    screen.MousePointer = 0
    Exit Sub
End If

Set tblPoints = DB.OpenTable("Points")       'Open Table
tblPoints.Index = "Points"                      'Open Index
tblPoints.MoveLast
RecCnt% = tblPoints.RecordCount

'Can't print distance if no race points
If RecCnt% < 1 Then
    MsgBox "No Race Points", 16, "CANNOT PRINT DISTANCES"
    tblPoints.Close
    tblMembers.Close
    DB.Close
    Set DB = Nothing
    Set tblPoints = Nothing
    Set tblMembers = Nothing
    screen.MousePointer = 0
    Exit Sub
End If

tblMembers.MoveFirst

While Not tblMembers.EOF
  tblPoints.MoveFirst
  'Convert to degrees decimal
  MLatDeg# = IIf(tblMembers("Lat_Deg") < 0, -(Abs(tblMembers("Lat_Deg")) + tblMembers("Lat_Min") / 60 + tblMembers("Lat_Sec") / 3600), tblMembers("Lat_Deg") + tblMembers("Lat_Min") / 60 + tblMembers("Lat_Sec") / 3600)
  MLngDeg# = IIf(tblMembers("Lng_Deg") < 0, -(Abs(tblMembers("Lng_Deg")) + tblMembers("Lng_Min") / 60 + tblMembers("Lng_Sec") / 3600), tblMembers("Lng_Deg") + tblMembers("Lng_Min") / 60 + tblMembers("Lng_Sec") / 3600)
  
  If MLatDeg# = 0 Or MLngDeg = 0 Then
    MsgBox "No Coordinates for Loft, Goto Main Menu then File Owner and add then try again", 16, "CANNOT CALCULATE DISTANCES"
    tblPoints.Close
    tblMembers.Close
    DB.Close
    Set DB = Nothing
    Set tblPoints = Nothing
    Set tblMembers = Nothing
    screen.MousePointer = 0
    Exit Sub
  End If

  'Convert to Radians
  mLatRad# = MakRad(MLatDeg#)
  mLngRad# = MakRad(MLngDeg#)

  While Not tblPoints.EOF
    'MsgPanel.Caption = "Please WAIT - Distance Calculation PRINTING," & tblPoints("PntName")
	    
    'if no Long or Lat just ignor
    'If RLngDeg# = 0 Or RLatDeg# = 0 Then
    If IsNull(tblPoints("Lat_Deg")) Or IsNull(tblPoints("Lat_Min")) Or IsNull(tblPoints("Lat_Sec")) Or IsNull(tblPoints("Lng_Deg")) Or IsNull(tblPoints("Lng_Min")) Or IsNull(tblPoints("Lng_Sec")) Then
	'Set distance to Zero
	DistGCM# = 0
    Else
	'Convert to degrees decimal
	RLatDeg# = IIf(tblPoints("Lat_Deg") < 0, -(Abs(tblPoints("Lat_Deg")) + tblPoints("Lat_Min") / 60 + tblPoints("Lat_Sec") / 3600), tblPoints("Lat_Deg") + tblPoints("Lat_Min") / 60 + tblPoints("Lat_Sec") / 3600)
	RLngDeg# = IIf(tblPoints("Lng_Deg") < 0, -(Abs(tblPoints("Lng_Deg")) + tblPoints("Lng_Min") / 60 + tblPoints("Lng_Sec") / 3600), tblPoints("Lng_Deg") + tblPoints("Lng_Min") / 60 + tblPoints("Lng_Sec") / 3600)

	'Convert Degrees to Radians
	rLatRad# = MakRad(RLatDeg#)
	rLngRad# = MakRad(RLngDeg#)

	'Calculate Distances
	'UK -   1NM = 1853.1833984418433594 Metres
	'Intl - 1NM = 1851.9993258722453125 Metres
	DistVincenty# = fnVincenty(mLatRad#, mLngRad#, rLatRad#, rLngRad#, "WGS84", WGS84_f, WGS84_eps)
	DistGCM_UK# = fnGCM_Dist(mLatRad#, mLngRad#, rLatRad#, rLngRad#, 1.85318339844184) * 1000
	DistGCM_Intl# = fnGCM_Dist(mLatRad#, mLngRad#, rLatRad#, rLngRad#, 1.85199932587225) * 1000
	DistGCM_TDPF# = fnGCM_DistTDPF(mLatRad#, mLngRad#, rLatRad#, rLngRad#)
	
	'DistVincenty# = fnVincenty(mLatRad#, mLngRad#, RLatDeg#, RLngRad#, "WGS84", WGS84_f, WGS84_eps)
	'NM# = 1853.18339844184
	'DistGCM_UK# = fnGCM_Dist(mLatRad#, mLngRad#, RLatDeg#, RLngRad#, NM#)
	'NM# = 1851.99932587225
	'DistGCM_Intl# = fnGCM_Dist(mLatRad#, mLngRad#, RLatDeg#, RLngRad#, NM#)
	'DistGCM_TDPF# = fnGCM_DistTDPF(mLatRad#, mLngRad#, RLatDeg#, RLngRad#)
	
	'Formulae Dist in NMiles = ACos((Sin(MLat)*Sin(RLat))+(Cos(MLat)*Cos(RLat)*Cos(MLng - RLng)))
	'CalcA# = Sin(MLatRad#) * Sin(RLatRad#)
	'CalcB# = Cos(MLatRad#) * Cos(RLatRad#)
	'CalcC# = Cos(MLngRad# - RLngRad#)
    
	'DistGCM# = ArcCos(CalcA# + (CalcB# * CalcC#))
	'No ArcCOS function so calculate
	'CalcD# = (CalcA# + (CalcB# * CalcC#))
	'CalcE# = ArcCos(CalcD#)
    
	'Convert to Degrees
	'CalcF# = MakDeg(CalcE#)
	'DistGCM# = CalcF# * 111139

	'DistGCM# = CalcF# * 60
	'Convert to Metres
	' DistGCM# = (DistGCM# * 1853 * 100000) / 100000
    End If
    
    'Update file with calculated distance
    tblPoints.Edit
    tblPoints("DistCalcGCM_TDPF") = DistGCM_TDPF#
    tblPoints("DistCalcGCM_UK") = DistGCM_UK#
    tblPoints("DistCalcGCM_Intl") = DistGCM_Intl#
    tblPoints("DistCalcGEOD") = DistVincenty#
    tblPoints("DistCalcGCM_TDPFyds") = DistGCM_TDPF# * 1.09361
    tblPoints("DistCalcGCM_UKyds") = DistGCM_UK# * 1.09361
    tblPoints("DistCalcGCM_Intlyds") = DistGCM_Intl# * 1.09361
    tblPoints("DistCalcGEODyds") = DistVincenty# * 1.09361
    tblPoints.Update
    
    tblPoints.MoveNext
  Wend

  tblMembers.MoveNext
Wend

tblPoints.Close
tblMembers.Close
Set tblPoints = Nothing
Set tblMembers = Nothing
DB.Close
screen.MousePointer = 0
Exit Sub

DB_Err:
    Select Case Err
    Case 3021
	MsgBox "Error " & Err & " " & Val(Error) & Chr(10) & "(Database Empty)", 16, "INFORMATION ONLY"
	Resume Next
    Case 3022
	MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
	MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
	MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
    End Select
    Resume Next
    'what about closing files?
End Sub

Function conf_del% (ObjName As String)
    ' Confirm the users wishes to delete.
    ' Returns 1 if the delete was confirmed otherwise
    ' returns 0.
    Dim Full_Message As String
    Dim Result As Integer
    
    Full_Message$ = "Are you sure you wish to delete this " & ObjName$ & "."
    
    Result% = MsgBox(Full_Message$, 4, "Confirm Deletion")

    If Result% = 6 Then
	conf_del% = 1
    Else
	conf_del% = 0
    End If

End Function

Function DateFix (ByVal dDate As String) As Variant
    'assume Date is valid if it got this far
    'if year is entered as 00 then assume 2000 otherwise accept
    
	If Len(dDate) < 10 Then
	    If InStr(1, dDate, "/") < 3 Then
		dDate = "0" & dDate
	    End If
	    If Len(dDate) < 9 Then
		dDate = Format(dDate, "ddddd")
		If Year(dDate) < 1930 And Year(dDate) > 1899 Then
		    dDate = DateAdd("yyyy", 100, dDate)
		End If
	    End If
	Else
	    dDate = Format(dDate, "ddddd")
	End If
	DateFix = dDate
End Function

Function DB_Backup () As String
    Dim DestFile As String
    Dim F As String
    Dim Msg As String   ' Declare variables.
    Dim x As String
    Dim reply As Integer
    Dim TbLoft As Table
    Dim TBMaster As Table
    Dim y As Integer

    On Error GoTo ErrHandler
    
    DB_Backup$ = "FAILED"
    y = 0

    F$ = Right$(Year(Now), 1) & Format$(Month(Now), "00") & Format$(Day(Now), "00")
    DestFile$ = "RP" & Trim(F$) & ".MDB"
    'DestFile$ = InputBox("Copy to where?")
    fmmain.MsgPanel.Caption = "REMEMBER - ALWAYS KEEP AT LEAST 3 SEPARATE BACKUPS OF YOUR DATA"
    reply% = MsgBox("Please insert USB Thumb drive in USB Port '" & GLBackup & "'", 49, "BACKUP DATA FILES")
    If reply% = 1 Then
	screen.MousePointer = 11
	fmmain.MsgPanel.Caption = "BACKUP " & DBFullFileName$ & " in progress - Please WAIT ..."
	FileCopy DBFullFileName$, DBPath$ & DestFile$   ' Copy file to destination.
	FileCopy DBFullFileName$, GLBackup & DestFile$   ' Copy file to destination.
	'FileCopy DBPath$ & "RP.DAT", GLBackup & "RP.DAT"   ' Copy file to destination.
	FreeLocks
	DoEvents
	fmmain.MsgPanel.Caption = ""
	screen.MousePointer = 0
	'Determine if backup rp.mdb exists
	x$ = Dir$(GLBackup & DestFile$)
	DoEvents
	If x$ = "" Then 'rp.mdb doesn't exist
	    MsgBox "Backup Not successful - use another method to backup RP.MDB"
	Else
	    'Open Loft & save new Backup date
	    MsgBox "Remove USB Key/Thumbdrive " & Chr(10) & "label as Backup with Date" & Chr(10) & "(Advisable to check USB Key has database file via My Computer/Filemanager or Windows Explorer)", 64, "BACKUP COMPLETE"
	    DB_Backup$ = "SUCCESS"
	    Set DBFile = OpenDatabase(DBFullFileName$)
	    Set TbLoft = DBFile.OpenTable("loft")'Open Table
	    TbLoft.MoveFirst
	    TbLoft.Edit
	    TbLoft("LastBackup") = Now
	    TbLoft.Update
	    FreeLocks
	    TbLoft.Close
	    DBFile.Close
	    Set TbLoft = Nothing
	    Set DBFile = Nothing
	    GetLoftDets
	End If
    Else
	MsgBox "User requests No Backup be run", 0, "BACKUP NOT RUN"
	DB_Backup$ = "Cancelled"
	Exit Function
    End If

    Exit Function

ErrHandler:
    y% = y% + 1
    If y% > 5 Then
	MsgBox "Backup FAILED - too many attempts", 48, "Retry later or use another method"
	Exit Function
    End If
    Select Case Err
	Case 55 'File already open
	    screen.MousePointer = 0
	    MsgBox Err & " " & Error(Err) & " Cannot copy an open file. Close it and try again.", 16, "WARNING - Backup UNSUCCESSFUL"
	    Exit Function
	Case 61 'Disk Full
	    screen.MousePointer = 0
	    MsgBox " Cannot backup as USB Thumb Drive/Key is FULL - replace", 48, "Click OK when ready to restart"
	    y% = y% + 2
	    Resume
	Case 68 'Device unavailable
	    screen.MousePointer = 0
	    MsgBox Err & " " & Error(Err) & " Device is unavailable", 48, "Click OK when ready to retry"
	    Resume
	Case 70 'Permission denied
	    screen.MousePointer = 0
	    MsgBox Err & " " & Error(Err) & " Check backup media, " & Chr(10) & "suspect too many files already on USB Key, not enough free space" & Chr(10) & "Hint: Try new USB Key", 48, "Click OK when ready to retry"
	    Resume
	Case 71 'Disk not ready
	    screen.MousePointer = 0
	    MsgBox "There is no disk in drive" & Chr(10) & "Hint: Put blank USB Key in USB port " & GLBackup$, 48, "Click OK when ready to retry"
	    Resume
	Case Else
	    screen.MousePointer = 0
	    MsgBox Err & " " & Error(Err) & " Try again later or use another method", 16, "WARNING - Backup UNSUCCESSFUL"
	    Exit Function
    End Select
    Resume Next
End Function

Function DB_Repair (ByVal TypeWork As String) As Integer
    Dim reply As Integer
    Dim F As String

    On Error GoTo ErrHandler2
    DB_Repair = 0
    
    '+++ Notify user to backup prior to Compression
    reply% = MsgBox("Have you done a backup just prior to this operation?", 36, "BACKUP COMPLETED? - Must be run prior to this repair")
    If reply% = 6 Then 'Yes
	reply% = MsgBox("Double Check - Have you done a backup just prior to this operation?", 36, "BACKUP COMPLETED?")
	If reply% = 6 Then 'Yes
	    screen.MousePointer = 11
	    '+++ Check enough space
	    fmmain.MsgPanel.Caption = TypeWork & "ING working Please Wait . . . . ."
	    DB_Repair = -1
	    If TypeWork = "REPAIR" Then
		RepairDatabase DBFullFileName$
		DoEvents
		FreeLocks
		fmmain.MsgPanel.Caption = "ATTEMPTED " & TypeWork & " COMPLETED"
	    Else 'TypeWork = "COMPRESS"
		'Open Database for exclusive use
		Set DBFile = OpenDatabase(DBFullFileName$, True, False)
		'Close Database
		DBFile.Close
		
		F = Dir(DBPath$ & "COMPRESS.MDB")
		If F <> "" Then
		    Kill DBPath$ & "COMPRESS.MDB"
		End If

		F = Dir(DBPath$ & "COMPRESS.LDB")
		If F <> "" Then
		    Kill DBPath$ & "COMPRESS.LDB"
		End If

		CompactDatabase DBFullFileName$, DBPath$ & "COMPRESS.MDB", DB_LANG_GENERAL, DB_ENCRYPT
		DoEvents
		F = Dir(DBPath$ & "rpMDB.BAK")
		If F = "" Then
		    Name DBFullFileName$ As DBPath$ & "rpMDB.BAK"
		Else
		    F = Dir(DBPath$ & "rpMDB1.BAK")
		    If F <> "" Then
			Kill DBPath$ & "rpMDB1.BAK"
			DoEvents
			Name DBPath$ & "rpMDB.BAK" As DBPath$ & "rpMDB1.BAK"
			DoEvents
		    Else
			Name DBPath$ & "rpMDB.BAK" As DBPath$ & "rpMDB1.BAK"
			DoEvents
		    End If
		    Name DBFullFileName$ As DBPath$ & "rpMDB.BAK"
		End If
		DoEvents
		Name DBPath$ & "COMPRESS.MDB" As DBFullFileName$
	    End If
	Else
	    MsgBox "Must do a backup prior to " & TypeWork, 48, "OPERATION ABORTED BY SYSTEM"
	End If
    Else
	MsgBox "Must do a backup prior to " & TypeWork, 48, "OPERATION ABORTED BY SYSTEM"
    End If
    screen.MousePointer = 0
Exit Function

ErrHandler2:
    If Err = 55 Then    ' File already open.
	screen.MousePointer = 0
	DB_Repair = 0
	MsgBox "Cannot copy an open file. Close it and try again."
	End
    ElseIf Err = 3049 Then
	MsgBox "Database filename - '" & DBFullFileName$ & "' -> " & TypeWork & " failed (NB: '" & FileNameini$ & "')", 48, "ERROR: Revert to previous good data backup"
	DB_Repair = 0
	End
    Else
	MsgBox Err & Error$
	DB_Repair = 0
	screen.MousePointer = 0
	End
    End If
End Function

Function Decrypt (ByVal Native As String) As String
    ' Encrypt also decrypts a string.
    Decrypt = Encrypt(Native$)
End Function

Function Encrypt (ByVal Native As String) As String
    Dim CounterP As Integer
    Dim ThisChar As String, NewChar As String
    Dim NewString As String

    Native$ = Trim$(Native$)

    NewString$ = ""
   
    For CounterP = Len(Native$) To 1 Step -1
	ThisChar$ = Mid$(Native$, CounterP%, 1)
	NewChar$ = ""
	If ThisChar$ >= "A" And ThisChar$ <= "Z" Then
	    NewChar$ = Chr(90 - (Asc(ThisChar$) - 65))
	End If
	If ThisChar$ >= "a" And ThisChar$ <= "z" Then
	    NewChar$ = Chr(122 - (Asc(ThisChar$) - 97))
	End If
	If ThisChar$ >= "0" And ThisChar$ <= "9" Then
	    NewChar$ = Chr(57 - (Asc(ThisChar$) - 48))
	End If
	If NewChar$ = "" Then NewChar$ = ThisChar$

	' Add to the string.
	NewString$ = NewString$ & NewChar$

    Next
  
    Encrypt = NewString$

End Function

Function Find_Value (ByVal fvSection As String, ByVal fvKey As String) As String
    Dim lpBuffer As String, lpDefault As String
    Dim lpSize As Integer, lpValid As Integer
    Dim lpReturnString As String
    
    On Error GoTo Conv_error

    ' Set up some of the variables needed to lookup the INI file.
    lpDefault$ = ""
    lpReturnString$ = Space$(128)
    lpSize% = Len(lpReturnString$)
    lpValid% = GetPrivateProfileString%(fvSection$, fvKey$, lpDefault$, lpReturnString$, lpSize%, FileNameini$)
    lpBuffer$ = Left$(lpReturnString$, lpValid%)
    'Find_Value = CDbl(lpBuffer$)
    Find_Value = lpBuffer$
    Exit Function

Conv_error:
    Find_Value = ""

    Exit Function
End Function

Function GetLicInfo ()
'reads security file RP.DAT'
'RP.DAT is encrypted so it must be unencrypted
'This function returns 32 char string from RP.DAT
'   NOT USED YET 13/2/99
    Dim x As String
    Dim U As String
    Dim U1 As String
    Dim char As String
    Dim ii As Integer
    Dim C1 As Integer

    x$ = "Read RP.DAT"
    Open DBPath$ & "RP.DAT" For Input As #1   ' Open file.
    U$ = ""
    U1$ = ""
    For ii% = 1 To 24 'read user name
	char$ = Input$(1, #1)
	C1% = Asc(char$)
	U$ = U$ + Chr$(C1% Xor &HFF)
    Next ii%
    For ii% = 1 To 8 'read lic
	char$ = Input$(1, #1)
	C1% = Asc(char$)
	U1$ = U1$ + Chr$(C1% Xor &HFF)
    Next ii%

End Function

Sub GetLoftDets ()
Dim x As String
Dim TbLoft As Table
Dim reply As Integer

On Error GoTo err1:

'Open Files
    x$ = "Open Database(GetLoftDets)"
    Set DBFile = OpenDatabase(DBFullFileName$)
    x$ = "Open Loft(GetLoftDets)"
    Set TbLoft = DBFile.OpenTable("loft")'Open Table
    TbLoft.MoveFirst

    x$ = "owneradd1"
    If IsNull(TbLoft("owneradd1")) Then
	GLAddress1$ = "Demo Address line 1"
	GLAddress2$ = "Demo Address line 2"
    Else
	GLAddress1$ = "" & TbLoft("owneradd1")
	GLAddress2$ = "" & TbLoft("owneradd2")
    End If

    x$ = "ownercity"
    If IsNull(TbLoft("ownercity")) Then
	GLCity$ = "City"
    Else
	GLCity$ = "" & TbLoft("ownercity")
    End If
    
    x$ = "ownerstate"
    If IsNull(TbLoft("ownerstate")) Then
	GLState$ = "State"
    Else
	GLState$ = "" & TbLoft("ownerstate")
    End If
    x$ = "ownerpc"
    If IsNull(TbLoft("ownerpc")) Then
	GLPostcode$ = "9999"
    Else
	GLPostcode$ = "" & TbLoft("ownerpc")
    End If
    x$ = "ownerph"
    If IsNull(TbLoft("ownerph")) Then
	GLPhoneNo$ = "61 7 46309009"
    Else
	GLPhoneNo$ = "" & TbLoft("ownerph")
    End If
    x$ = "license"
    If IsNull(TbLoft("license")) Then
	GLSerialNo$ = "Demo"
    Else
	GLSerialNo$ = "" & TbLoft("license")
    End If
    x$ = "ownermark"
    If IsNull(TbLoft("ownermark")) Then
	GLmark$ = "DEMO"
    Else
	GLmark$ = "" & TbLoft("ownermark")
    End If
    x$ = "ownerfed"
    If IsNull(TbLoft("ownerfed")) Then
	GLFed$ = "DEMO"
    Else
	GLFed$ = "" & TbLoft("ownerfed")
    End If
    x$ = "bscmonth"
    If IsNull(TbLoft("bscmonth")) Then
	GLbscmonth% = 6
    ElseIf Not IsNumeric(TbLoft("bscmonth")) Then
	GLbscmonth% = 6
    Else
	GLbscmonth% = TbLoft("bscmonth")
    End If
    x$ = "name"
    If IsNull(TbLoft("name")) Then
	GLLoft$ = "DEMO Loft"
    Else
	GLLoft$ = "" & TbLoft("name")
    End If
    
    x$ = "ownername"
    If IsNull(TbLoft("ownername")) Then
	GLOwner$ = "Demonstration Copy"
    Else
	GLOwner$ = "" & TbLoft("ownername")
    End If
    x$ = "RSLPP"
    If IsNull(TbLoft("RSLPP")) Then
	GLRSLpp% = 30
    Else
	GLRSLpp% = TbLoft("RSLPP")
    End If
    x$ = "RSRaceFormType"
    If IsNull(TbLoft("RSRaceFormType")) Then
	GLRSRaceFormType% = 1
    Else
	GLRSRaceFormType% = TbLoft("RSRaceFormType")
    End If
    x$ = "DBVersion"
    If IsNull(TbLoft("DBVersion")) Then
	GLDBVersion = "DEMO V3.2p"
    Else
	GLDBVersion = TbLoft("DBVersion")
    End If
    x$ = "OwnerClub"
    If GLOwner$ = "Demonstration Copy" Then
	GLOwnerClub$ = "DEMONSTRATION USE ONLY"
    ElseIf IsNull(TbLoft("ownerclub")) Then
	GLOwnerClub$ = "DEMONSTRATION USE ONLY"
    Else
	GLOwnerClub$ = TbLoft("ownerclub")
    End If
       
    'Store Loft Longitude & Latitude for use later
    If IsNull(TbLoft("Lat_deg") Or TbLoft("Lat_min") Or TbLoft("Lat_sec")) Then
	    GLLoftLat = 0
	    GLLoftLng = 0
    ElseIf IsNull(TbLoft("Lng_deg") Or TbLoft("Lng_min") Or TbLoft("Lng_sec")) Then
	    GLLoftLat = 0
	    GLLoftLng = 0
    Else
	    GLLoftLat = CDbl(TbLoft("Lat_deg") + (TbLoft("Lat_min") / 60) + (TbLoft("Lat_Sec") / 3600))
	    GLLoftLng = CDbl(TbLoft("Lng_deg") + (TbLoft("Lng_min") / 60) + (TbLoft("Lng_Sec") / 3600))
    End If
    If IsNull(TbLoft("LastBackup")) Then
	GLLastBU = CVDate("01/01/1980")
    Else
	GLLastBU = TbLoft("LastBackup")
    End If

    If IsNull(TbLoft("DistanceType")) Then
	GLDistType = "M"
    ElseIf TbLoft("DistanceType") = "M" Then
	GLDistType = "M"
    ElseIf TbLoft("DistanceType") = "Y" Then
	GLDistType = "Y"
    End If

    TbLoft.Close    'Close Table
    DBFile.Close    'Close Database
    Set TbLoft = Nothing
    
    Exit Sub

err1:
    Select Case Err
    Case 0
	Resume Next
    Case 76
	MsgBox "'" & FileNameini$ & "' points to Database filename - '" & DBFullFileName$ & "' which doesn't exist. Hint: modify rp.ini.", 48, "ERROR: File doesn't exist"
	GoTo ErrClose:
    Case 3011
	MsgBox "Database corrupted - or incorrect version - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "To Fix run Database Repair"
	GoTo ErrClose:
    Case 3018
	MsgBox "Field '" & x$ & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
	GoTo ErrClose:
    Case 3021
	'MsgBox "No records", 0, "Error"
	Resume Next
    Case 3022
	MsgBox "record already selected", 16, "Cannot add record"
	GoTo ErrClose:
    Case 3024
	MsgBox "CANNOT find Database filename - '" & DBFullFileName$ & "'. To fix EITHER restore database file OR modify file '" & FileNameini$ & "' to look to correct directory (ie., to where your rp.mdb database is located.)", 48, "ERROR: Cannot start program"
	GoTo ErrClose:
    Case 3026
	MsgBox "Disk is full no more additions possible.", 48, "Add Error"
	GoTo ErrClose:
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
	GoTo ErrClose:
    Case 3049
	MsgBox "Database filename - '" & DBFullFileName$ & "'. To fix Click OK and software will attempt repair.", 48, "ERROR: Database needs repair"
	Rem MsgBox "Database filename - '" & DBFullFileName$ & "' -> " & TypeWork & " failed (NB: '" & FileNameini$ & "')", 48, "ERROR: Revert to previous good data backup"
	reply% = DB_Repair("REPAIR")
	'reply% = DB_Repair("COMPRESS")
	If reply% = -1 Then
	    MsgBox "Attempted repair appears successful", 16, "Repair completed"
	    Resume
	Else
	    MsgBox "Repair not completed", 16, "Program Ending"
	    GoTo ErrClose:
	End If
    Case 48
	MsgBox "Corrupt Windows Dynamic Link Library, Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
	GoTo ErrClose:
    Case Else
	If x$ = "Open Database(GetLoftDets)" Then
	    MsgBox Err & " " & Error$(Err) & " '" & x$ & "' failed - " & DBFullFileName$ & "'" & Chr(10) & "System finds Database is already open" & Chr(10) & "Suggestion: Ensure program not already running, if not then try" & Chr(10) & "Doing normal shutdown of computer & power off then power on & retry" & Chr(10) & "If this doesn't work, repair database", 16, "Can't start system"
	Else
	    MsgBox "Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
	End If
	GoTo ErrClose:
    End Select

ErrClose:
    TbLoft.Close    'Close Table
    DBFile.Close    'Close Database
    Set TbLoft = Nothing
    Set DBFile = Nothing
    End
End Sub

Function LicenseDisplay (ByVal Lic)
    Dim x As Integer
    Dim x1 As String
    Dim y1 As String
    Dim y As Double
    
    x% = Len(Lic)
    x1$ = Mid$(Lic, 1, x% - 1)
    y1$ = Mid$(Lic, x%, 1)
    y = Val(x1$)
    y = Int((((y / 52) + .942) / 1.5) * 93)
    LicenseDisplay = Trim(Str(y)) & y1$
End Function

Function LicNo (ByVal usr As String) As String

ReDim iUserName(24) As Integer
Dim dLic As Double
Dim iX As Integer
Dim sA As String
Dim txtSecIDa As String
Dim txtSecID As String

dLic# = 0

usr$ = Trim(usr)

For iX% = 1 To Len(usr$)
    iUserName(iX%) = Asc(Mid$(usr$, iX%, 1))
    If iX% = 2 Then
	dLic# = dLic# * 12
    Else
	dLic# = dLic# + (iUserName(iX%) * iX%)
    End If
Next iX%

sA$ = Mid$(Str(dLic#), 3, 2)

If Val(Mid(Str(dLic#), 3, 2)) > Asc("A") And Val(Mid(Str(dLic#), 3, 2)) < Asc("Z") Then
	txtSecIDa$ = Chr(Val(Mid(Str(dLic#), 3, 2)))
ElseIf Val(Mid(Str(dLic#), 3, 3)) > Asc("a") And Val(Mid(Str(dLic#), 3, 3)) < Asc("z") Then
	txtSecIDa$ = Chr(Val(Mid(Str(dLic#), 3, 3)))
ElseIf Val(Mid(Str(dLic#), 3, 2)) < 26 Then
    txtSecIDa$ = Chr(Val(Mid(Str(dLic#), 3, 2)) + 65)
ElseIf Val(Mid(Str(dLic#), 3, 2)) / 4 < 25 Then
    txtSecIDa$ = Chr(Val((Mid(Str(dLic#), 3, 2)) / 4) + 65)
Else
    txtSecIDa$ = "A"
End If

dLic# = Int(dLic# * 119.2) + 100000

txtSecID$ = Str(dLic#)

LicNo = Trim(txtSecID$ & txtSecIDa$)

End Function

Sub Main ()

    screen.MousePointer = 11

    ' Get system parameters.
    System_Init
    
    ' Show introduction banner.
    fmmain.MsgPanel.Caption = "Please wait while system is being initialised.."
    fmmain.MsgPanel.FontBold = True
    If GLSystem = "RP" Then
	fmmain.Show
    Else
	fmmain1.Show
    End If
    
    screen.MousePointer = 0

End Sub

Sub Open_DBs ()
    On Error GoTo syserr2
    Set DBFile = OpenDatabase(DBFullFileName$)

syserr2:
    Select Case Err
    Case 0
	Resume Next
    Case 76
	MsgBox "'" & FileNameini$ & "' points to Database filename - '" & DBFullFileName$ & "' which doesn't exist. Hint: modify rp.ini.", 48, "ERROR: File doesn't exist"
	End
    Case 3011
	MsgBox "Database corrupted - or incorrect version - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "To Fix run Database Repair"
	End
    Case 3018
	MsgBox "Field '" & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
	End
    Case 3021
	'MsgBox "No records", 0, "Error"
	Resume Next
    Case 3022
	MsgBox "record already selected", 16, "Cannot add record"
	End
    Case 3024
	MsgBox "CANNOT find Database filename - '" & DBFullFileName$ & "'. To fix EITHER restore database file OR modify file '" & FileNameini$ & "' to look to correct directory (ie., to where your rp.mdb database is located.)", 48, "ERROR: Cannot start program"
	End
    Case 3026
	MsgBox "Disk is full no more additions possible.", 48, "Add Error"
	End
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
	End
    Case 3049
	MsgBox "Database filename - '" & DBFullFileName$ & "' " & FileNameini$ & "'. To fix Click OK and software will attempt repair.", 48, "ERROR: Database needs repair"
	Resume
    Case 48
	MsgBox "Corrupt Windows Dynamic Link Library, Error - " & Err & " " & Error$(Err) & " - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
	End
    Case Else
	MsgBox "Error - " & Err & " " & Error$(Err) & " - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
	End
    End Select
    Resume Next
    
End Sub

Sub PrtOrder ()

Dim MlMargin As Integer
Dim mMsg As String
Dim iResponse As Integer

mMsg$ = "Click OK when PRINTER is " & Chr(10) & "POWERED ON and ON LINE?"
iResponse% = MsgBox(mMsg$, 1, "PRINT ORDER FORM")

If iResponse% <> 1 Then   '0 = OK
    Exit Sub
End If

'Print the Pedigree
screen.MousePointer = 11 'Set Mouse pointer to Hourglass

MlMargin% = 350

'PRINT HEADING
Printer.FontName = "Arial"   ' Select font and point size.
Printer.ScaleMode = 1          ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "ORDER"

Printer.CurrentX = 4700
Printer.CurrentY = 950
Printer.Print "FORM"

'Print "Post to" line
Printer.FontSize = 14
Printer.FontBold = False
Printer.CurrentX = MlMargin%
Printer.CurrentY = 2100
Printer.Print "TO:"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 2450
Printer.Print "Mr Kieran Malone (Ph 0409495571)"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 2800
Printer.Print "The Lodge"
Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 3150
Printer.Print "71 Hunter Crc"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 3500
Printer.Print "PETRIE QLD 4502"

'Print the request line
Printer.CurrentX = MlMargin% + 750
Printer.CurrentY = 4500
Printer.Print "Please rush me The Racing Pigeon Management Software"

Printer.CurrentX = MlMargin% + 750
Printer.CurrentY = 4900
Printer.Print "As payment please find enclosed my Cheque or Money Order"
Printer.Print ""

Printer.CurrentX = MlMargin% + 1000
Printer.CurrentY = 5500
Printer.Print "Please provide on DVD or USB Key (Circle Preference)"
Printer.CurrentY = 5500
Printer.CurrentX = MlMargin% + 6900
Printer.Print "$125.00"
Printer.CurrentY = 6000
Printer.CurrentX = MlMargin% + 1000
Printer.Print "OR"
Printer.CurrentY = 6500
Printer.CurrentX = MlMargin% + 1000
Printer.Print "Please provide with training videos"
Printer.CurrentY = 6500
Printer.CurrentX = MlMargin% + 6900
Printer.Print "$125.00 ........."
Printer.Print ""

'Post To
Printer.CurrentX = MlMargin%
Printer.CurrentY = 7500
Printer.Print "DELIVERY ADDRESS:"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 8150
Printer.Print "Name:"

Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 8650
Printer.Print "......................................................................"

Printer.FontSize = 14
Printer.FontBold = False
Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 9150
Printer.Print "Address:"

Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 9550
Printer.Print "......................................................................"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 9950
Printer.Print "......................................................................"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 10350
Printer.Print "......................................................................"


Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 10750
Printer.Print "......................................................................"

Printer.CurrentX = MlMargin%
Printer.CurrentY = 13000
Printer.Print "_______________________________________________________________________________________"

'Print Software details & Licence No. & Copyright notice.
Printer.FontItalic = False
Printer.FontSize = 12
Printer.FontBold = True
Printer.CurrentX = MlMargin%
Printer.CurrentY = 13400

'Print Software Serial No.
Printer.FontBold = False
Printer.FontSize = 10
Printer.CurrentX = MlMargin%
Printer.Print "The Racing Pigeon Management software (Copyright 1988-2013, All rights reserved, Kieran Malone Ph 0409495571)"
Printer.CurrentX = MlMargin%
Printer.Print "Software Licence to:"
Printer.CurrentX = MlMargin%
Printer.Print "Name: " & GLOwner
Printer.CurrentX = MlMargin%
Printer.Print "Loft: " & GLLoft
Printer.CurrentX = MlMargin%
Printer.Print GLAddress1
Printer.CurrentX = MlMargin%
Printer.Print GLAddress2
Printer.CurrentX = MlMargin%
Printer.Print GLCity & " "; GLState & " "; GLPostcode
Printer.CurrentX = MlMargin%
Printer.Print "Phone: " & GLPhoneNo
Printer.CurrentX = MlMargin%
Printer.Print "Federation: " & GLFed
Printer.CurrentX = MlMargin%
Printer.Print "Club: "; GLOwnerClub
Printer.CurrentX = MlMargin%
Printer.Print "Serial No. "; LicenseDisplay(GLSerialNo$)
Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
screen.MousePointer = 0

End Sub

Sub prtsys ()
Dim MlMargin As Integer
Dim mMsg As String
Dim iResponse As Integer

screen.MousePointer = 11

mMsg$ = "Click OK when PRINTER is " & Chr(10) & "POWERED ON and ON LINE?"
iResponse% = MsgBox(mMsg$, 1, "PRINT SYSTEM DETAILS")

If iResponse% <> 1 Then   '0 = OK
    screen.MousePointer = 0
    Exit Sub
End If

screen.MousePointer = 11 'Set Mouse pointer to Hourglass

MlMargin% = 250

'PRINT HEADING
Printer.FontName = "Arial"   ' Select font and point size.
Printer.ScaleMode = 1          ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "SYSTEM DETAILS"

'Print "Post to" line
Printer.FontSize = 12
Printer.FontBold = False
Printer.CurrentX = MlMargin%
Printer.CurrentY = 2100
Printer.Print "TO:"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 2450
Printer.Print "Mr Kieran Malone"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 2800
Printer.Print "Goolawan"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 3150
Printer.Print "182 Breydons Rd"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 3500
Printer.Print "HODGSONVALE QLD 4352"

Printer.CurrentX = MlMargin% + 550
Printer.CurrentY = 3900
Printer.Print "Ph 07 46309009"

'Print the request line
Printer.CurrentY = 4500
Printer.Print "_______________________________________________________________________________________"

'Print Software details & Licence No. & Copyright notice.
Printer.FontBold = False
Printer.FontSize = 8

Printer.CurrentX = MlMargin% + 550
Printer.Print "Windows Directory - " & WinPath$
Printer.CurrentX = MlMargin% + 550
Printer.Print "Application Directory - " & App.Path
Printer.CurrentX = MlMargin% + 550
Printer.Print "Full Database Name - " & DBFullFileName$
Printer.CurrentX = MlMargin% + 550
Printer.Print "Database Name - " & DBFileName$
Printer.CurrentX = MlMargin% + 550
Printer.Print "Database Path - " & DBPath$
Printer.CurrentX = MlMargin% + 550
Printer.Print "ini filename - " & FileNameini$
Printer.CurrentX = MlMargin% + 550
Printer.Print "Registered Owner - " & GLOwner
Printer.CurrentX = MlMargin% + 550
Printer.Print GLOwner1
Printer.CurrentX = MlMargin% + 550
Printer.Print GLAddress1
Printer.CurrentX = MlMargin% + 550
Printer.Print GLAddress2
Printer.CurrentX = MlMargin% + 550
Printer.Print GLCity
Printer.CurrentX = MlMargin% + 550
Printer.Print GLState
Printer.CurrentX = MlMargin% + 550
Printer.Print GLPostcode
Printer.CurrentX = MlMargin% + 550
Printer.Print "Loft Name - " & GLLoft
Printer.CurrentX = MlMargin% + 550
Printer.Print GLPhoneNo
Printer.CurrentX = MlMargin% + 550
Printer.Print "Serial No. - " & LicenseDisplay(GLSerialNo$)
Printer.CurrentX = MlMargin% + 550
Printer.Print "Default Ring Mark - " & GLmark
Printer.CurrentX = MlMargin% + 550
Printer.Print "Default Fed Name (as appears on Race sheets) - " & GLFed
Printer.CurrentX = MlMargin% + 550
Printer.Print "Month for start of breeding season - " & Str(GLbscmonth)
Printer.CurrentX = MlMargin% + 550
Printer.Print "Lines per page on race sheet - " & Str(GLRSLpp)
Printer.CurrentX = MlMargin% + 550
Printer.Print "Type of race sheet used (system allows 2 different type) - " & GLRSRaceFormType
Printer.CurrentX = MlMargin% + 550
Printer.Print "Database version in use - " & GLDBVersion
Printer.CurrentX = MlMargin% + 550
Printer.Print "Application Version - " & GLAPPversion
Printer.CurrentX = MlMargin% + 550
Printer.Print "GLLfmbird - " & GLLfmbird
Printer.CurrentX = MlMargin% + 550
Printer.Print "GLLfmerror - " & GLLfmerror
Printer.CurrentX = MlMargin% + 550
Printer.Print "fm - " & fm
Printer.CurrentX = MlMargin% + 550
Printer.Print "Demo setup - " & Str(GLDemo)
Printer.CurrentX = MlMargin% + 550
Printer.Print "Software Serial No. " & LicenseDisplay(GLSerialNo$)
Printer.CurrentX = MlMargin%
Printer.CurrentY = 14350
Printer.Print "______________________________________________________________________________________________________________"

Printer.CurrentX = MlMargin% + 550
Printer.Print "The Racing Pigeon Management software (Copyright 1988-2013, All rights reserved, Kieran Malone Ph 0409495571)"

Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
screen.MousePointer = 0

End Sub

Function strMth (ByVal Pmon As Integer)
Select Case Pmon
    Case 1
	strMth = "January"
    Case 2
	strMth = "February"
    Case 3
	strMth = "March"
    Case 4
	strMth = "April"
    Case 5
	strMth = "May"
    Case 6
	strMth = "June"
    Case 7
	strMth = "July"
    Case 8
	strMth = "August"
    Case 9
	strMth = "September"
    Case 10
	strMth = "October"
    Case 11
	strMth = "November"
    Case 12
	strMth = "December"
    End Select

End Function

Sub System_Init ()
    On Error GoTo syserr:
    
    Dim TBMaster As Table
    Dim lpBuffer As String
    Dim lpSize As Integer, lpValid As Integer
    Dim ii As Integer
    Dim U As String
    Dim U1 As String
    Dim u2 As String
    Dim char As String
    Dim C1 As Integer
    Dim x As String
    Dim Recs As Long
    Dim SaveTitle As String
    Dim reply As Integer
    
    GLBackup$ = "A:\"
    GLSystem = "RP"

    'Find the windows directory.
    lpBuffer$ = String$(145, 0)
    lpSize% = GetWindowsDirectory(lpBuffer$, 145)
    WinPath$ = Left$(lpBuffer$, lpSize%)
    FileNameini$ = WinPath$ & "\RP.INI"

    'Determine if rp.ini exists, if not then advise user
    x$ = Dir$(FileNameini$)
    If x$ = "" Then 'doesn't exist so workaround
	If App.PrevInstance Then
	    'read rp.ini file & if [Bird Stud] then start reading 2nd DB otherwise
	    SaveTitle$ = App.Title
	    App.Title = "... duplicate instance."
	    fmmain.Caption = "... duplicate instance."
	    AppActivate SaveTitle$
	    SendKeys "% R", True
	    End
	End If
	DBPath$ = App.Path & "\"
	DBFileName$ = "rp.mdb"
	DBFullFileName$ = DBPath$ & DBFileName$
	GLBackup$ = "A:\"
	GLSystem = "RP"
	ii% = MsgBox("Click on 'Yes' if you want System to use default location for database eg. '" & DBFullFileName$ & "' otherwise click 'No' and program will abort allowing you to restore/recreate the rp.ini file.", 36, "WARNING RP.INI Missing - Proceed with caution")
	If ii% <> 6 Then
	    End
	End If
    Else
	If App.PrevInstance Then
	    'read rp.ini file & if [Bird Stud] then start reading 2nd DB otherwise
	    DBPath$ = Find_Value("Bird Stud", "DataPath")
	    DBFileName$ = Find_Value("Bird Stud", "Databasename")
	    GLBackup$ = Find_Value("Bird Stud", "Backup")
	    GLSystem$ = Find_Value("Bird Stud", "System")
	    DBFullFileName$ = DBPath$ & DBFileName$
	    If GLBackup$ = "" Then
		GLBackup$ = "A:\"
	    End If
	    If GLSystem$ = "" Then
		GLSystem$ = "Bird Stud"
	    End If
	    If DBPath$ = "" Then
		SaveTitle$ = App.Title
		App.Title = "... duplicate instance."
		fmmain.Caption = "... duplicate instance."
		AppActivate SaveTitle$
		SendKeys "% R", True
		End
	    End If
	Else
	    DBPath$ = Find_Value("Racing Pigeons", "DataPath")
	    DBFileName$ = Find_Value("Racing Pigeons", "Databasename")
	    GLBackup$ = Find_Value("Racing Pigeons", "Backup")
	    GLSystem$ = Find_Value("Racing Pigeons", "System")
	    If GLBackup$ = "" Then
		GLBackup$ = "A:\"
	    End If
	    If GLSystem$ = "" Then
		GLSystem$ = "RP"
	    End If
	    DBFullFileName$ = DBPath$ & DBFileName$
	End If
    End If
    
    'Determine if rp.mdb exists
    x$ = Dir$(DBFullFileName$)
    If x$ = "" Then 'rp.mdb doesn't exist
	MsgBox "Database file '" & DBFullFileName$ & "' not found. Hint: check '" & FileNameini$ & "' settings also check if database exists.", 48, "FATAL ERROR Database missing"
	End
    End If

    'Determine if rp.dat exists if not then run register screen
    x$ = Dir$(DBPath$ & "\rp.dat")
    If x$ = "" Then 'rp.dat doesn't exist
	frmRegister.Show 1
    Else
	GLDemo = 0
    End If
    
    If GLDemo = -1 Then End
    
    'Get User Details
    GetLoftDets

    x$ = "Read RP.DAT"
    Open DBPath$ & "RP.DAT" For Input As #1   ' Open file.
    U$ = ""
    U1$ = ""
    For ii% = 1 To 24 'read user name
	char$ = Input$(1, #1)
	C1% = Asc(char$)
	U$ = U$ + Chr$(C1% Xor &HFF)
    Next ii%
    For ii% = 1 To 8 'read lic
	char$ = Input$(1, #1)
	C1% = Asc(char$)
	U1$ = U1$ + Chr$(C1% Xor &HFF)
    Next ii%
    Close #1  ' Close
    
    'Check loft details against RP.DAT Security file
    If Trim(U$) <> Trim(GLOwner$) Then
	MsgBox "Report to Supplier - quote Warning #99706", 16, "Restart program to reregister software"
	GLOwner$ = "Demonstration Copy"
	GLLoft$ = "Loft name"
	GLAddress1$ = "Demo Address line 1"
	GLAddress2$ = "Demo Address line 2"
	MsgBox "Report to Supplier - quote Warning #99706", 16, "Restart program to reregister software"
	Kill (DBPath$ & "rp.dat")
	End
    End If
    If Trim(U1$) <> Trim(GLSerialNo$) Then
	GLOwner$ = "Demonstration Copy"
	GLLoft$ = "Loft name"
	GLAddress1$ = "Demo Address line 1"
	GLAddress2$ = "Demo Address line 2"
	Kill (DBPath$ & "rp.dat")
	MsgBox "Report to Supplier - quote Warning #99707", 16, "Restart program to reregister software"
	End
    End If
    If LicNo(Trim(GLOwner$)) <> Trim(GLSerialNo$) Then
	U1$ = LicNo(Trim(GLOwner$))
	GLOwner$ = "Demonstration Copy"
	U1$ = LicNo(Trim(GLOwner$))
	GLLoft$ = "Loft name"
	GLAddress1$ = "Demo Address line 1"
	GLAddress2$ = "Demo Address line 2"
	Kill (DBPath$ & "rp.dat")
	MsgBox "Report to Supplier - quote Warning #99708", 16, "Changed to Demonstration mode"
	End
    End If
    
    'If Demo, check that no of birds in Master file
    'Doesn't exceed 100
    If GLDemo = 9 Or GLOwner = "Demonstration Copy" Then
	Set DBFile = OpenDatabase(DBFullFileName$)
	Set TBMaster = DBFile.OpenTable("master")'Open Table
	TBMaster.MoveLast
	Recs& = TBMaster.RecordCount
	TBMaster.Close
	DBFile.Close
	Set TBMaster = Nothing
	If Recs& > 50 And Recs& < 60 Then
	    MsgBox "Demonstration mode allows approx 45 Racing Pigeon records." & Chr(13) & "This is now exceeded. All data entered can be retained, however you must obtain a Licence No." & Chr(13) & "Order software for a registration code" & Chr(13) & "May be done quickly by phone 07 46309009.", 16, "WARNING - Adding further records without a full licence may lock system"
	    FMabout.Show 1
	    End
	End If
	If Recs& > 59 Then
	    MsgBox "Demonstration mode allows approx 45 Racing Pigeon records." & Chr(13) & "This no. is now exceeded. All data entered can be retained, however you must obtain a Licence No. to continue." & Chr(13) & "Order software for a registration code" & Chr(13) & "May be done quickly by phone 07 46309009.", 16, "Demo mode maximum records is now exceeded"
	    FMabout.Show 1
	    End
	End If
    End If
    
Exit Sub

syserr:
    Select Case Err
    Case 0
	Resume Next
    Case 76
	MsgBox "'" & FileNameini$ & "' points to Database filename - '" & DBFullFileName$ & "' which doesn't exist. Hint: modify rp.ini.", 48, "ERROR: File doesn't exist"
	End
    Case 3011
	MsgBox "Database corrupted - or incorrect version - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Attempt Database Repair(System_Init)"
	MsgBox "Field '" & x$ & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
	End
    Case 3018
	MsgBox "Field '" & x$ & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
	End
    Case 3021
	'MsgBox "No records", 0, "Error"
	Resume Next
    Case 3022
	MsgBox "record already selected", 16, "Cannot add record"
	End
    Case 3024
	MsgBox "CANNOT find Database filename - '" & DBFullFileName$ & "'. To fix EITHER restore database file OR modify file '" & FileNameini$ & "' to look to correct directory (ie., to where your rp.mdb database is located.)", 48, "ERROR: Cannot start program"
	End
    Case 3026
	MsgBox "Disk is full no more additions possible.", 48, "Add Error"
	End
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
	End
    Case 3049
	MsgBox "Database filename - '" & DBFullFileName$ & "' " & FileNameini$ & "'. To fix Click OK and software will attempt repair.", 48, "ERROR: Database needs repair"
	reply% = DB_Repair("REPAIR")
	If reply% = -1 Then
	    MsgBox "Attempted repair appears successful", 16, "Repair completed"
	    Resume
	Else
	    MsgBox "Repair not completed", 16, "Program Ending"
	    End
	End If
    Case 48
	MsgBox "Corrupt Windows Dynamic Link Library, Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system(System_Init)"
	End
    Case 62 'Can't read rp.dat
	MsgBox "Error - " & Err & "'" & Error$(Err) & "', Corruption in registration details - You will need to reregister your software - refer manual or video on how to register software.", 16, "Error - Can't start system"
	Close #1  ' Close
	Kill DBPath$ & "RP.DAT"    ' Remove security file.
	End
    Case Else
	If x$ = "Open Database" Then
	    MsgBox Err & " " & Error$(Err) & " '" & x$ & "' failed - " & DBFullFileName$ & "'" & Chr(10) & "System finds Database is already open" & Chr(10) & "Suggestion: Ensure program not already running, if not then try" & Chr(10) & "Doing normal shutdown of computer & power off then power on & retry" & Chr(10) & "If this doesn't work, repair database", 16, "Can't start system"
	Else
	    MsgBox "Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
	End If
	End
    End Select
    Resume Next
End Sub

Function WrtRPDat (ByVal usr, ByVal Lic) As Integer
'create rp.dat
On Error GoTo Err2:
    screen.MousePointer = 11
    Dim LenTestString As Integer
    Dim A As String
    Dim program As Integer
    Dim FirstPosition As Long
    Dim Ctr As Integer
    Dim I As Long
    Dim C As String
    Dim J As Integer
    Dim T As String
    Dim U As String
    Dim U1 As String
    Dim Z As Integer
    Dim C1 As Integer
    Dim ByteCnt As Long
    Dim DBFile As Database
    Dim TbLoft As Table
    
    ByteCnt& = 0
    LenTestString% = 32     'Store length of test string
    program% = 1            'Use a descriptive name for file number
    Open DBPath$ & "rp.dat" For Binary Access Read Write As #program%
    FirstPosition& = 1

    A$ = String$(1, 0)
    
    'Encrypt User string
    T$ = Left$(usr, 24) 'Text1 contains rego info
    T$ = T$ & Space(24 - Len(T$)) 'Pad out with spaces
    U$ = ""         'U$ will be the encrypted string
    U1$ = "" 'U1$ will be encrupted License No.
    For Z% = 1 To 24 'For the length of the string
	C1% = Asc(Mid$(T$, Z%, 1))'Get the ASC code of chr
	U$ = U$ + Chr$(C1% Xor &HFF)'Encrypt chr & append
    Next Z%
		
    'Encrypt the Licence Code
    Lic = Lic & Space(8 - Len(Lic)) 'Pad out with spaces
    For Z% = 1 To 8 'For the length of the string
	C1% = Asc(Mid$(Lic, Z%, 1))'Get the ASC code of chr
	U1$ = U1$ + Chr$(C1% Xor &HFF)'Encrypt chr & append
    Next Z%
		
    U$ = U$ & U1$
		
    C$ = Left$(U$, 1)   'Get the first char
    Put #program%, FirstPosition&, C$   'Output List
    For J% = 2 To 32 'Write other chars
	C$ = Mid$(U$, J%, 1)
	If C$ = "" Then 'If out of chars in name
	    C$ = Chr$(Asc(" ") Xor &HFF)    'encrypt space
	End If
	Put #program%, , C$
    Next J%
		
    Close #program%
    
    'Update Loft table
    Set DBFile = OpenDatabase(DBFullFileName$)
    Set TbLoft = DBFile.OpenTable("loft") 'Open Table
    TbLoft.MoveFirst
    TbLoft.Edit
    TbLoft("ownername") = Trim(usr)
    TbLoft("License") = Trim(Lic)
    TbLoft.Update
    TbLoft.Close
    DBFile.Close
    Set TbLoft = Nothing
    Set DBFile = Nothing

    WrtRPDat = -1
    screen.MousePointer = 0
    Exit Function

Err2:
    MsgBox "Error is " & Err & " " & Error
    Resume Next

End Function

