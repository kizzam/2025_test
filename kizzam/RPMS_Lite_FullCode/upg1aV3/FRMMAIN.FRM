VERSION 2.00
Begin Form frmMain 
   Caption         =   "Upgrade Software"
   ClientHeight    =   4305
   ClientLeft      =   690
   ClientTop       =   1620
   ClientWidth     =   5310
   Height          =   4815
   Left            =   630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5310
   Top             =   1170
   Width           =   5430
   Begin SSCommand butt3PedgUpg 
      Caption         =   "Update Pedg"
      Font3D          =   0  'None
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin SSCommand butt3UpdRPExtra 
      Caption         =   "Update RPExtra"
      Font3D          =   0  'None
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin SSCommand cmdUpgrade 
      Caption         =   "&Start Upgrade"
      Font3D          =   0  'None
      Height          =   975
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin SSCommand cmdFixRTrans 
      Caption         =   "Fix y2K Problem"
      Font3D          =   0  'None
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin SSCommand cmdCheck 
      Caption         =   "&Check"
      Font3D          =   0  'None
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   4575
   End
End
Option Explicit

Sub butt3PedgUpg_Click ()

updatePedg

'Global db As database
'Global td As TableDef
'Global Fld As field
'Set db = OpenDatabase("C:\rpigeon\rpextra.mdb")

'If Not TableExists("Pedg") Then GoTo errorhandler

'Set db = CurrentDb()
'> For Each tdf In db.TableDefs
'> If (tdf.Attributes And dbSystemObject) = 0 Then
'> If (tdf.Name <> "Switchboard Items") And (tdf.Connect = "")
'> Then
'> For Each fld In tdf.Fields
'> If fld.Properties("AllowZeroLength") = False Then
'> fld.Properties("AllowZeroLength") = True
'> End If
'> Next
'> End If
'> End If
'> Next

'Create table object
'Set Td = db.CreateTableDef("Pedg")

'Must add a field
'This adds an auto-incremented ID field

'Set fld = Td.CreateField("PrtColour1", db_Integer)
'f.Attributes = dbAutoIncrField
'NewFld.Name = "ID"
'NewFld.Type = 4  'DB_LONG
'NewFld.Attributes = 16 'dbAutoIncrField

'Append field to table
'Td.Fields.Append fld

'Append Table to Database
'db.TableDefs.Append Td

'db.Close
'Set Td = Nothing
'Set fld = Nothing


Exit Sub

'errorhandler:
'If Not oDB Is Nothing Then oDB.Close
'Set NewTd = Nothing
'Set NewFld = Nothing

End Sub

Sub butt3UpdRPExtra_Click ()
Dim response As Integer
response = ModifyTable(RPextra.mdb, "Pedg")

End Sub

Private Sub cmdCheck_Click ()
    'Dim db As Database
    'Dim td As TableDef
    'Dim Fld As Field
    'Dim FldName As String
    'Dim x As Integer
    'AppName = "Upgrade V3.2 to V3.2z"
    'Screen.MousePointer = 11
        
    'Update Table LOFT
    'Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
    'Set td = db.TableDefs("Loft")
    'FldName$ = "ownerfed"
    'For x = 1 To 3
        ' Delete the newly added fields.
        'FldName$ = "OwnerFed"
        'FldName$ = InputBox("Field", "Input field", 1)

        'If FieldExists(td, FldName$) Then
            'MsgBox "Field Exists"
            'AppendDeleteField Td, "DELETE", FldName$
            'AppendDeleteField Td, "APPEND", FldName$, dbText, 39
        'Else
            'MsgBox "Field Doesn't Exists"
            'AppendDeleteField Td, "APPEND", FldName$, dbText, 39
        'End If
        'Label1.caption = "Upgrade Completed"
        'cmdUpgrade.Enabled = False
        'Screen.MousePointer = 0
    'Next x
End Sub

Private Sub cmdExit_Click ()
    Unload Me
End Sub

'
Private Sub cmdFixRTrans_Click ()
    Fix_relDate
End Sub

'
Sub cmdUpgrade_Click ()
    Dim db As Database
    Dim Td As TableDef
    Dim fld As field
    Dim FldName As String
    Dim x As Integer
    Dim FldValue As String
    Dim CmdLine As String
    Dim result As Integer
    Dim retval As Double
        
    cmdUpgrade.Enabled = False
    
    result = MsgBox("Have you completed two(2) separate backups on separate diskettes?", 36, "MUST do BACKUP prior to this operation!")
    If result <> 6 Then
        Label2.Caption = "Upgrade ABORTED - MUST do backup first"
        MsgBox "See Manual re Backing up RPMS, after backup complete you can attempt upgrade!", 48, "BACKUP MUST BE RUN PRIOR THIS UPGRADE"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Label2.Caption = "PHASE 1 Copying Files - Please Wait..."
    frmCopy.Show 1
    
    Label2.Caption = "PHASE 2 Modifying Database - Please Wait..."
    frmDBupg.Show 1
    DoEvents
    
    Label2.Caption = "PHASE 3 Creating Batch file - Please Wait..."
    CreateFile2
    
    Label2.Caption = "PHASE 4 - Install upgraded files" & Chr(10)
    retval = Shell(DestPath$ & "upg.bat", 1)
    
    Label2.Caption = "INSTALLATION COMPLETED" & Chr(10)
    Screen.MousePointer = 0

Exit Sub

ErrorSetup:
    MsgBox "Error in upgrade", 48, "CANNOT CONTINUE"
    End
    Exit Sub
End Sub

Sub Form_Load ()
Label2.Caption = "Ensure you have backup of your data first"
End Sub

Sub updatePedg ()
 
 Dim dbsRPMS As Database
 Dim tdfNew As TableDef
 Dim prpLoop As Property
 
 Set dbsRPMS = OpenDatabase("RPExtra.mdb")
 
 ' Create a new TableDef object.
 Set tdfNew = dbsRPMS.CreateTableDef("TEST")
 
 'With tdfNew
 ' Create fields and append them to the new TableDef
 ' object. This must be done before appending the
 ' TableDef object to the TableDefs collection of the
 ' RPMS database.
 tdfNew.Fields.Append tdfNew.CreateField("Prt_Colour1", dbInteger)
tdfNew.Fields.Append tdfNew.CreateField("LastName", dbText)
tdfNew.Fields.Append tdfNew.CreateField("Phone", dbText)
'  .Fields.Append .CreateField("Notes", dbMemo)
 
' Debug.Print "Properties of new TableDef object " & _
 '"before appending to collection:"
 
 ' Enumerate Properties collection of new TableDef
 ' object.
 For Each prpLoop In tdfNew.Properties
 On Error Resume Next
 If prpLoop <> "" Then Debug.Print " " & _
 prpLoop.Name & " = " & prpLoop
 On Error GoTo 0
 Next prpLoop
 
 ' Append the new TableDef object to the RPMS
 ' database.
 dbsRPMS.TableDefs.Append tdfNew
 
 Debug.Print "Properties of new TableDef object " & _
 "after appending to collection:"
 
 ' Enumerate Properties collection of new TableDef
 ' object.
 For Each prpLoop In .Properties
 On Error Resume Next
 If prpLoop <> "" Then Debug.Print " " & _
 prpLoop.Name & " = " & prpLoop
 On Error GoTo 0
 Next prpLoop
 
 End With
 
 ' Delete new TableDef object since this is a
 ' demonstration.
 dbsRPMS.TableDefs.Delete "TEST"
 
 dbsRPMS.Close
 
End Sub

