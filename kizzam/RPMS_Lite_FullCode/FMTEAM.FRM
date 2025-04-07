VERSION 2.00
Begin Form FMteam 
   Caption         =   "Team Code Maintenance"
   ClientHeight    =   4215
   ClientLeft      =   450
   ClientTop       =   1590
   ClientWidth     =   8505
   Height          =   4725
   Left            =   390
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   8505
   Top             =   1140
   Width           =   8625
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "teams"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0FFC0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   8535
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0FFC0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      Begin SSCommand ButtDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7080
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin SSCommand ButtAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7080
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin SSCommand ButtUpdate 
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7080
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7080
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin SSCommand ButtOK 
         Caption         =   "&Ok"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H00FFFF80&
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   2  'Entire grid
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeadingHeight   =   1.5
         Height          =   2775
         HorzLines       =   0  'None
         Layout          =   FMTEAM.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   0
         Top             =   840
         UseBookmarks    =   -1  'True
         Width           =   6615
         WrapCellPointer =   0   'False
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Team Codes"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
End
Option Explicit

Sub ButtAdd_Click ()
    FMaddteam.Show 1
End Sub

Sub ButtCancel_Click ()
    If ButtUpdate.Enabled Then
        'Exit form
        Unload FMteam
    Else
        'Disable Update
        ButtUpdate.Enabled = True
        ButtAdd.Enabled = False
        ButtDelete.Enabled = False
        Table1.Editable = False
        Table1.Refresh
    End If

End Sub

Sub ButtDelete_Click ()

    Const IDYES = 6, IDNO = 7
    Dim Key As String
    Dim Key1 As String
    Dim Msg As String
    Dim Counter As Integer
    Dim Response As Integer
    Dim ActionQry As String
    Dim DB As Database
    Dim Tbl As Table
    
    Table1.ColumnIndex = 1
    Key$ = Table1.ColumnText(1)
    Key1$ = Table1.ColumnText(2)

    'Find Record to Delete, Make current
    Data1.Recordset.FindFirst "code = '" & Key$ & "'"
    If Data1.Recordset.NoMatch Then
        MsgBox "Cannot find record to delete", 16, "Delete Error"
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Ensure NOT system record, can't be changed
    If Data1.Recordset("status") = "S" Then
        MsgBox "Cannot DELETE this code, It is set by the system", 16, "ERROR CANNOT DELETE"
        Screen.MousePointer = 0
        Exit Sub
    Else
    End If
    
    Msg$ = "Are you sure you want to DELETE " & Key$ & " " & Key1$ & " ?"
    Response = MsgBox(Msg$, 36, "Delete Record")
    If Response = IDYES Then
        Screen.MousePointer = 11
        'Ensure NOT used by a bird record, retain db integrity
        Set DB = OpenDatabase(DBFullFileName$)    'Open Database
        Set Tbl = DB.OpenTable("teambird")     'Open Table
        Tbl.Index = "code"                 'Open Index
        Tbl.Seek "=", Key$ 'Ensure code to be deleted NOT in use
        If Tbl.NoMatch Then
            Data1.Recordset.Delete
            Data1.Refresh
            MsgBox "Record Deleted", 16, "Notification Only"
            Tbl.Close
            DB.Close
        Else
            MsgBox "Code in use on bird " & Tbl("yr") & " " & Trim(Tbl("mark")) & " " & Trim(Tbl("ringno")), 16, "Notification CANNOT DELETE"
            Tbl.Close
            DB.Close
            End If
    End If

    Screen.MousePointer = 0

End Sub

Sub ButtOK_Click ()
    If Table1.Editable = True Then
        'Update table
        Table1.Refresh
    Else
        'Exit form
        Unload FMteam
    End If
End Sub

Sub ButtUpdate_Click ()
    Table1.Enabled = True
    Table1.Editable = True
    ButtAdd.Enabled = True
    ButtDelete.Enabled = True
    ButtUpdate.Enabled = False
End Sub

Sub Form_Load ()
    'Set DBFile = OpenDatabase(DBFullFileName$)
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "teams"
    Data1.Refresh
    ButtAdd.Enabled = False
    ButtDelete.Enabled = False
    ButtUpdate.Enabled = True
    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True
End Sub

