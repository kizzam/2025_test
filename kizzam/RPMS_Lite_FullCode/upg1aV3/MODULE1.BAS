Option Explicit
Global WinSysDir As String
Global DBFileName As String
Global DBPath As String
Global DestPath As String
Global SourcePath As String
Global FileNameini As String
Global DBFullFileName As String
Global db As database
Global td As TableDef
Global Fld As field

'vb3
Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer)
'vb5
'Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function GetWindowsDirectory Lib "Kernel" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Function AppendDeleteField (tblName As String, strCommand As String, fldName As String, fldType1, fldSize) As Integer
        AppendDeleteField = False
        ' Check first to see if the TableDef object is
        ' updatable. If it isn't, control is passed back to
        ' the calling procedure.
        If db.TableDefs(tblName).Updatable = False Then
            MsgBox tblName & "TableDef not Updatable! Unable to complete task."
            AppendDeleteField = False
            Exit Function
        End If
        ' Depending on the passed data, append or delete a
        ' field to the Fields collection of the specified
        ' TableDef object.
        If strCommand = "APPEND" Then
            '.Fields.Append .CreateField(strName, varType1, varSize)
            Dim f1 As New field
            f1.Name = fldName
            f1.Type = fldType1
            f1.Size = fldSize
            db.TableDefs(tblName).Fields.Append f1
        ElseIf strCommand = "DELETE" Then
            db.TableDefs(tblName).Fields.Delete fldName
        End If
    
        AppendDeleteField = True

End Function

Sub CreateFile2 ()
    Dim LineA As String
    Dim Dr As String
    Dim Pth As String
    Dr$ = Mid$(DestPath$, 1, 2)
    Pth$ = Mid$(DestPath$, 3, Len(Trim(DestPath$)) - 2)
    ChDrive Mid$(Dr$, 1, 1)
    ChDir Pth$
    
    Open "Yes" For Output As #1
    LineA$ = "Y"
    Print #1, LineA$
    Close #1
    
    Open "Upg.bat" For Output As #1
    LineA$ = "echo Upgrade Phase 2 - Started"
    Print #1, LineA$
    LineA$ = Dr$
    Print #1, LineA$
    LineA$ = "cd " & Pth$
    Print #1, LineA$
    LineA$ = "echo Backup old files"
    Print #1, LineA$
    LineA$ = "md upg_v32z"
    Print #1, LineA$
    LineA$ = "cd upg_v32z"
    Print #1, LineA$
    LineA$ = "copy ..\rp.exe rpold.exe"
    Print #1, LineA$
    LineA$ = "copy ..\rp.mdb rpold.mdb"
    Print #1, LineA$
    LineA$ = "copy ..\*.rpt"
    Print #1, LineA$
    LineA$ = "copy ..\rpextra.mdb"
    Print #1, LineA$
    LineA$ = "cd .."
    Print #1, LineA$
    LineA$ = "del *.rpt < Yes"
    Print #1, LineA$
    LineA$ = "del rp.exe < Yes"
    Print #1, LineA$
    LineA$ = "RPMS < Yes"
    Print #1, LineA$
    LineA$ = "echo Upgrade Phase 4 - Completed"
    Print #1, LineA$
    LineA$ = "echo Click on 'X' top right to close window"
    Print #1, LineA$
    Close #1

End Sub

Function createLoft1 ()
Dim Newtd As New TableDef
Dim NewFld As New field

NewFld.Name = ""
NewFld.Type = db_Integer
NewFld.Size = 0
Newtd.Fields.Append NewFld

db.TableDefs.Append Newtd

End Function

Function CreateTable (ByVal DatabaseName As String, ByVal TableName As String) As Integer
'CREATE TABLE

'DataBaseName is the file/path name of the database
'TableName is the name of the table you want to create
'Returns true if successful, false otherwise
'Project must include reference to DAO

On Error GoTo errorhandler
Dim oDB As database
Dim Newtd As New TableDef
Dim NewFld As New field

On Error GoTo errorhandler

CreateTable = False

Set oDB = OpenDatabase(DatabaseName)
If TableExists(TableName) Then GoTo errorhandler

'Create table object
'Set td = oDB.CreateTableDef(TableName)

'Must add a field
'This adds an auto-incremented ID field

'Set f = td.CreateField("ID", dbLong)
'f.Attributes = dbAutoIncrField
NewFld.Name = "ID"
NewFld.Type = 4  'DB_LONG
NewFld.Attributes = 16 'dbAutoIncrField

'Append field to table
td.Fields.Append NewFld

'Append Table to Database
oDB.TableDefs.Append Newtd

oDB.Close
Set Newtd = Nothing
Set NewFld = Nothing
CreateTable = True
Exit Function

errorhandler:
If Not oDB Is Nothing Then oDB.Close
Set Newtd = Nothing
Set NewFld = Nothing
CreateTable = False
End Function

'Adding a Field to a Table
'You can add a field to an existing table by appending a field to an existing Fields collection. The following code adds two new fields, Address and Phone, to the Authors table.
Function FieldExists (tblName As String, fldName As String) As Integer
        Dim i As Integer
        FieldExists = False
        ' Check first to see if the TableDef object is
        ' updatable. If it isn't, control is passed back to
        ' the calling procedure.
        If db.TableDefs(tblName).Updatable = False Then
            MsgBox tblName$ & "TableDef not Updatable! Unable to complete task."
            Exit Function
        End If
        ' Enumerate the Fields collection to show the new fields.
        For i% = 0 To db.TableDefs(tblName).Fields.Count - 1
            If UCase(db.TableDefs(tblName).Fields(i%).Name) = UCase(fldName) Then
                FieldExists = True
                Exit Function
            End If
        Next i%
End Function

Function FileExists (ByVal strPathName As String) As Integer
'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0
End Function

Function Find_Value (ByVal fvSection As String, ByVal fvKey As String) As String
    Dim lpBuffer As String, lpDefault As String
    Dim lpSize As Long, lpValidL As Long
    Dim lpReturnString As String
    
    On Error GoTo Conv_error

    ' Set up some of the variables needed to lookup the INI file.
    lpDefault$ = ""
    lpReturnString$ = Space$(128)
    lpSize& = Len(lpReturnString$)
    lpValidL& = GetPrivateProfileString(fvSection$, fvKey$, lpDefault$, lpReturnString$, lpSize&, FileNameini$)
    lpBuffer$ = Left$(lpReturnString$, lpValidL&)
    Find_Value = lpBuffer$
    Exit Function

Conv_error:
    Find_Value = ""

    Exit Function
End Function

Sub Fix_relDate ()
Dim db As database
Dim td As Table
Dim x As String

On Error GoTo err1

x = ""

'MsgBox "Opening DB..."
Set db = OpenDatabase(DBFullFileName$)
Set td = db.OpenTable("racetrans")

td.MoveLast
If td.RecordCount = 0 Then Exit Sub
td.MoveFirst
Do While Not td.EOF
    If Year(td!reldate) = 1900 Then
        td.Edit
        td!reldate = DateAdd("yyyy", 100, td!reldate)
        td.Update
    End If
    td.MoveNext
    Loop
td.Close
'Now do Race table
Set td = db.OpenTable("race")
td.MoveLast
If td.RecordCount = 0 Then Exit Sub
td.MoveFirst
Do While Not td.EOF
    If Year(td!reldate) = 1900 Then
        td.Edit
        td!reldate = DateAdd("yyyy", 100, td!reldate)
        td.Update
    End If
    td.MoveNext
    Loop
td.Close
db.Close
Set td = Nothing
Set db = Nothing

Exit Sub

err1:
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
        MsgBox "Field '" & x$ & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
        End
    Case 3021
        Resume Next
    Case 3024
        MsgBox "CANNOT find Database filename - '" & DBFullFileName$ & "'. To fix EITHER restore database file OR modify file '" & FileNameini$ & "' to look to correct directory (ie., to where your rp.mdb database is located.)", 48, "ERROR: Cannot start program"
        End
    Case 3026
        MsgBox "Disk is full no more additions possible.", 48, "Add Error"
        End
    Case 3049
        MsgBox "Database filename - '" & DBFullFileName$ & "' " & FileNameini$ & "'. To fix Click OK and software will attempt repair.", 48, "ERROR: Database needs repair"
        Resume
    Case 48
        MsgBox "Corrupt Windows Dynamic Link Library, Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
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

Sub Main ()
    On Error GoTo syserr:
    
    Dim lpBuffer As String
    Dim lpSize As Integer, lpValid As Integer
    Dim ii As Integer
    Dim x As String
    Dim RetVal As Double
    
    'Find the windows directory.
    lpBuffer$ = String$(145, 0)
    lpSize% = GetWindowsDirectory(lpBuffer$, 145)
    WinSysDir$ = Left$(lpBuffer$, lpSize%)
    FileNameini$ = WinSysDir$ & "\RP.INI"

    'Determine if rp.ini exists, if not then advise user
    x$ = Dir$(FileNameini$)
    If x$ = "" Then
        MsgBox "File doesn't exist"
    Else
        DBPath$ = Find_Value("Racing Pigeons", "DataPath")
        DBFileName$ = Find_Value("Racing Pigeons", "Databasename")
        DBFullFileName$ = DBPath$ & DBFileName$
        DestPath$ = DBPath$
    End If
    
    'Determine if rp.mdb exists
    x$ = Dir$(DBFullFileName$)
    If x$ = "" Then 'rp.mdb doesn't exist
        MsgBox "Database file '" & DBFullFileName$ & "' not found. Hint: check '" & FileNameini$ & "' settings also check if database exists.", 48, "FATAL ERROR Database missing"
        End
    End If

    'Determine current path & Directory (Where is upgrade being installed from?)
    SourcePath$ = CurDir
    SourcePath$ = App.Path
    
    'Determine if rp.dat exists if not then run register screen
    x$ = Dir$(DBPath$ & "\rp.dat")
    If x$ = "" Then 'rp.dat doesn't exist
        MsgBox "Need to register Software prior to upgrade"
        End
    Else
        Set db = OpenDatabase(DBFullFileName$)
        frmMain.Show 1
        db.Close
        Set td = Nothing
        Set db = Nothing
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
        MsgBox "Database corrupted - or incorrect version - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "To Fix run Database Repair"
        End
    Case 3018
        MsgBox "Field '" & x$ & "' missing, suspect incorrect version Database - '" & DBFullFileName$ & "' " & Err & " " & Error$(Err), 48, "Fatal Error - cannot continue"
        End
    Case 3024
        MsgBox "CANNOT find Database filename - '" & DBFullFileName$ & "'. To fix EITHER restore database file OR modify file '" & FileNameini$ & "' to look to correct directory (ie., to where your rp.mdb database is located.)", 48, "ERROR: Cannot start program"
        End
    Case 3026
        MsgBox "Disk is full no more additions possible.", 48, "Add Error"
        End
    Case 3049
        MsgBox "Database filename - '" & DBFullFileName$ & "' " & FileNameini$ & "'. To fix Click OK and software will attempt repair.", 48, "ERROR: Database needs repair"
        Resume
    Case 48
        MsgBox "Corrupt Windows Dynamic Link Library, Error - " & Err & " " & Error$(Err) & " '" & x$ & "' - " & DBFullFileName$ & "'", 16, "Error - Can't start system"
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

Function ModifyTable (ByVal DatabaseName As String, ByVal TableName As String) As Integer

'MODIFY TABLE
'DataBaseName is the file/path name of the database
'TableName is the name of the table you want to create
'Returns true if successful, false otherwise
'Project must include reference to DAO

On Error GoTo errorhandler
Dim oDB As database
Dim Newtd As New TableDef
Dim NewFld As New field

On Error GoTo errorhandler

ModifyTable = False

Set oDB = OpenDatabase(DatabaseName)
If Not TableExists(TableName) Then GoTo errorhandlerN

'Create table object
'Set td = oDB.CreateTableDef(TableName)

'Must add a field
'This adds an auto-incremented ID field

'Set f = td.CreateField("ID", dbLong)
'f.Attributes = dbAutoIncrField
'NewFld.Name = "ID"
'NewFld.Type = 4  'DB_LONG
'NewFld.Attributes = 16 'dbAutoIncrField

'This adds integer field
Dim x As Integer

For x = 1 To 31
    'Set NewFld = Newtd.CreateField(Trim("PrtColour" & Trim(CStr(x))), db_Integer)
    NewFld.Name = Trim("PrtColour" & Trim(CStr(x)))
    NewFld.Type = db_Integer
    'NewFld.Attributes = 16 'dbAutoIncrField
    'NewFld.Required = False
    NewFld.Size = 0
    'NewFld.AllowZeroLength = True
    'NewFld.DefaultValue = 0
'Append field to table
    Newtd.Fields.Append NewFld
 Next x

'Append Table to Database
'oDB.tableDefs.Append NewTd

oDB.Close
Set Newtd = Nothing
Set NewFld = Nothing
ModifyTable = True
Exit Function

errorhandlerN:
If Not oDB Is Nothing Then oDB.Close
Set Newtd = Nothing
Set NewFld = Nothing
ModifyTable = False
End Function

Function TableExists (tblName As String) As Integer
    Dim i As Integer
    On Error Resume Next
    TableExists = True
    For i% = 0 To db.TableDefs.Count
        If UCase(db.TableDefs(i%).Name) = UCase(tblName) Then
            TableExists = True
            Exit Function
        End If
    Next i%
End Function

