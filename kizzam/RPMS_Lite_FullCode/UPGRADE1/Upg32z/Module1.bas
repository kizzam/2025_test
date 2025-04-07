Attribute VB_Name = "Module1"
Option Explicit
Global WinSysDir As String
Global DBFileName As String
Global DBPath As String
Global DestPath As String
Global SourcePath As String
Global FileNameini As String
Global DBFullFileName As String
Global Cnt As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'CREATE TABLE
Public Function CreateTable(ByVal DatabaseName As String, _
   ByVal TableName As String) As Boolean

'DataBaseName is the file/path name of the database
'TableName is the name of the table you want to create
'Returns true if successful, false otherwise
'Project must include reference to DAO

On Error GoTo errorhandler
Dim oDB As DAO.Database
Dim td As DAO.TableDef
Dim f As DAO.Field

Set oDB = Workspaces(0).OpenDatabase(DatabaseName)
On Error GoTo errorhandler
If TableExists(oDB, TableName) Then GoTo errorhandler
'Create table object
Set td = oDB.CreateTableDef(TableName)

'Must add a field
'This adds an auto-incremented ID field

Set f = td.CreateField("ID", dbLong)
f.Attributes = dbAutoIncrField

'Append field to table
td.Fields.Append f

'Append Table to Database
oDB.TableDefs.Append td

oDB.Close
CreateTable = True
Exit Function
errorhandler:
If Not oDB Is Nothing Then oDB.Close
Set td = Nothing
Set f = Nothing

End Function
Public Function TableExists(oDB As Database, TableName As String) As Boolean
    
    Dim td As DAO.TableDef
    On Error Resume Next

    Set td = oDB.TableDefs(TableName)
    TableExists = Err.Number = 0
End Function
Sub CreateFile1()
    Dim Line As String
    Dim Dr As String
    Dim Pth As String
    Dr$ = Mid$(DestPath$, 1, 2)
    Pth$ = Mid$(DestPath$, 3, Len(Trim(DestPath$)) - 2)
    ChDrive Mid$(Dr$, 1, 1)
    ChDir Pth$
    Open "Upg.bat" For Output As #1
    'Line$ = "Copy a:\rpms.32x " & DestPath$ & "rpms.exe"
    'Print #1, Line$
    Line$ = Dr$
    Print #1, Line$
    'Line$ = "echo Backup old files"
    'Print #1, Line$
    Line$ = "cd " & Pth$
    Print #1, Line$
    'Line$ = "del *.rpt /Y"
    'Print #1, Line$
    'Line$ = "pause"
    'Print #1, Line$
    'Line$ = "del rp.exe /Y"
    'Print #1, Line$
    'Line$ = "pause"
    'Print #1, Line$
    'Line$ = "md v32y"
    'Print #1, Line$
    'Line$ = "cd v32y"
    'Print #1, Line$
    'Line$ = "copy ..\*.rpt"
    'Print #1, Line$
    'Line$ = "copy ..\rp.exe"
    'Print #1, Line$
    'Line$ = "copy ..\rp.mdb"
    'Print #1, Line$
    'Line$ = "cd .."
    'Print #1, Line$
    'Line$ = "echo Upgrade reports"
    'Print #1, Line$
    Line$ = "RPMS < Yes"
    Print #1, Line$
    'Line$ = "RPMS1 /Y"
    'Print #1, Line$
    Line$ = "echo Upgrade Complete"
    Print #1, Line$
    'Line$ = "pause"
    'Print #1, Line$
    Close #1
End Sub
Sub CreateFile2()
    Dim Line As String
    Dim Dr As String
    Dim Pth As String
    Dr$ = Mid$(DestPath$, 1, 2)
    Pth$ = Mid$(DestPath$, 3, Len(Trim(DestPath$)) - 2)
    ChDrive Mid$(Dr$, 1, 1)
    ChDir Pth$
    
    Open "Yes" For Output As #1
    Line$ = "Y"
    Print #1, Line$
    Close #1
    
    Open "Upg.bat" For Output As #1
    Line$ = "echo Upgrade Phase 2 - Started"
    Print #1, Line$
    Line$ = Dr$
    Print #1, Line$
    Line$ = "cd " & Pth$
    Print #1, Line$
    Line$ = "echo Backup old files"
    Print #1, Line$
    Line$ = "md upg_v32z"
    Print #1, Line$
    Line$ = "cd upg_v32z"
    Print #1, Line$
    Line$ = "copy ..\rp.exe rpold.exe"
    Print #1, Line$
    Line$ = "copy ..\rp.mdb rpold.mdb"
    Print #1, Line$
    Line$ = "copy ..\*.rpt"
    Print #1, Line$
    Line$ = "copy ..\rpextra.mdb"
    Print #1, Line$
    Line$ = "cd .."
    Print #1, Line$
    Line$ = "del *.rpt < Yes"
    Print #1, Line$
    Line$ = "del rp.exe < Yes"
    Print #1, Line$
    Line$ = "RPMS < Yes"
    Print #1, Line$
    Line$ = "echo Upgrade Phase 4 - Completed"
    Print #1, Line$
    Line$ = "echo Click on 'X' top right to close window"
    Print #1, Line$
    Close #1

End Sub

'Adding a Field to a Table
'You can add a field to an existing table by appending a field to an existing Fields collection. The following code adds two new fields, Address and Phone, to the Authors table.
Function FieldExists(tdfTemp As TableDef, strName As String)
    Dim fldLoop As Field
    FieldExists = False
    With tdfTemp
        ' Check first to see if the TableDef object is
        ' updatable. If it isn't, control is passed back to
        ' the calling procedure.
        If .Updatable = False Then
            MsgBox "TableDef not Updatable! " & _
                "Unable to complete task."
            Exit Function
        End If
        ' Enumerate the Fields collection to show the new fields.
        For Each fldLoop In tdfTemp.Fields
            If UCase(fldLoop.Name) = UCase(strName) Then
                FieldExists = True
                Exit Function
            End If
        Next fldLoop
    End With
End Function
Function AppendDeleteField(tdfTemp As TableDef, _
    strCommand As String, strName As String, _
    Optional varType, Optional varSize)
    
    AppendDeleteField = False
    With tdfTemp
        ' Check first to see if the TableDef object is
        ' updatable. If it isn't, control is passed back to
        ' the calling procedure.
        If .Updatable = False Then
            MsgBox "TableDef not Updatable! " & _
                "Unable to complete task."
            Exit Function
        End If
        ' Depending on the passed data, append or delete a
        ' field to the Fields collection of the specified
        ' TableDef object.
        If strCommand = "APPEND" Then
            .Fields.Append .CreateField(strName, _
                varType, varSize)
        Else
            If strCommand = "DELETE" Then .Fields.Delete strName
        End If
    End With
    
    AppendDeleteField = True

End Function
'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Public Function FileExists(ByVal strPathName As String) As Integer
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
Sub Main()
    Dim pas As Integer
    On Error GoTo syserr:
    
    Dim lpBuffer As String
    Dim lpSize As Integer, lpValid As Integer
    Dim ii As Integer
    Dim x As String
    Dim RetVal As Double
    Dim Pos As Integer
    
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
    Pos = InStr(Len(SourcePath$), SourcePath$, "\", 1)
    If Pos <> Len(SourcePath$) Then
        SourcePath$ = SourcePath$ & "\"
    End If
    Pos = InStr(Len(DestPath$), DestPath$, "\", 1)
    If Pos <> Len(DestPath$) Then
        DestPath$ = DestPath$ & "\"
    End If
    
    
    
    'Determine if rp.dat exists if not then run register screen
    x$ = Dir$(DBPath$ & "\rp.dat")
    If x$ = "" Then 'rp.dat doesn't exist
        MsgBox "Need to register Software prior to upgrade"
        End
    Else
        frmMain.Show 1
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

Function Find_Value(ByVal fvSection As String, ByVal fvKey As String) As String
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
Sub Fix_relDate()
Dim wrkJet As Workspace
Dim db As Database
Dim td As Recordset
Dim tr As Recordset
Dim x As String

On Error GoTo err1

x = ""

'Create Microsoft Jet Workspace object.
Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)

'MsgBox "Opening DB..."
Set db = wrkJet.OpenDatabase(DBFullFileName$, True)

'Now do Race table
Set tr = db.OpenRecordset("race", dbOpenTable)
tr.MoveLast
If tr.RecordCount = 0 Then Exit Sub
tr.MoveFirst
Do While Not tr.EOF
    If Year(tr!reldate) < 1903 Then
        tr.Edit
        tr!reldate = DateAdd("yyyy", 100, tr!reldate)
        tr.Update
    End If
    tr.MoveNext
Loop
tr.Close

'Now do Race Trans table
Set td = db.OpenRecordset("racetrans", dbOpenTable)
Set tr = db.OpenRecordset("race", dbOpenTable)
tr.Index = "RaceCode"

td.MoveLast
If td.RecordCount = 0 Then Exit Sub
td.MoveFirst
Do While Not td.EOF
    If IsNull(td!reldate) Then
        tr.Seek "=", td!RaceCode
        If tr.NoMatch Then
            td.Edit
            td!reldate = Now()
            td.Update
        Else
            td.Edit
            td!reldate = tr!reldate
            td.Update
        End If
    End If
    If Year(td!reldate) < 1903 Then
        td.Edit
        td!reldate = DateAdd("yyyy", 100, td!reldate)
        td.Update
    End If
    td.MoveNext
Loop

td.Close
tr.Close
db.Close
wrkJet.Close
Set td = Nothing
Set tr = Nothing
Set db = Nothing
Set wrkJet = Nothing

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
