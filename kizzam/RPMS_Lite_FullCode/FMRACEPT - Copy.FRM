VERSION 2.00
Begin Form FMracepnt 
   Caption         =   "Race Point Maintenance/Enquiry"
   ClientHeight    =   4680
   ClientLeft      =   240
   ClientTop       =   1485
   ClientWidth     =   9135
   Height          =   5085
   Left            =   180
   LinkTopic       =   "Form3"
   ScaleHeight     =   4680
   ScaleWidth      =   9135
   Top             =   1140
   Width           =   9255
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   9135
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9135
      Begin SSCommand ButtDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   3240
         Width           =   975
      End
      Begin SSCommand ButtAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7920
         TabIndex        =   2
         Top             =   2520
         Width           =   975
      End
      Begin SSCommand ButtUpdate 
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7920
         TabIndex        =   1
         Top             =   1800
         Width           =   975
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7920
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin SSCommand ButtOK 
         Caption         =   "&Ok"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7920
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "points"
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
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
         HeadingHeight   =   1
         Height          =   3015
         HorzLines       =   0  'None
         Layout          =   FMRACEPT.FRX:0000
         LayoutIndex     =   1
         Left            =   120
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1080
         UseBookmarks    =   -1  'True
         Width           =   7695
         WrapCellPointer =   0   'False
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Race Point Details"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   3135
      End
   End
End
Option Explicit

Sub ButtAdd_Click ()
    FMaddracepnt.Show 1
End Sub

Sub ButtCancel_Click ()
    If ButtUpdate.Enabled Then
        'Exit form
        Unload FMracepnt
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
    Data1.Recordset.FindFirst "code_pnt = '" & Key$ & "'"
    If Data1.Recordset.NoMatch Then
        MsgBox "Cannot find record to delete", 16, "Delete Error"
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Ensure NOT system record, can't be changed
    If Data1.Recordset("system") Then
        MsgBox "Cannot DELETE this code, It is set by the system", 16, "ERROR CANNOT DELETE"
        Screen.MousePointer = 0
        Exit Sub
    Else
    End If
    
    Msg$ = "Are you sure you want to DELETE record " & Key$ & " " & Key1$ & " ?"
    Response = MsgBox(Msg$, 36, "Delete Record")
    If Response = IDYES Then
        Screen.MousePointer = 11
        'Ensure NOT used by a race record, retain db integrity
        Set DB = OpenDatabase(DBFullFileName$)    'Open Database
        Set Tbl = DB.OpenTable("race")     'Open Table
        Tbl.Index = "race_pnt"                 'Open Index
        Tbl.Seek "=", Key$ 'Ensure code to be deleted NOT in use
        If Tbl.NoMatch Then
            Data1.Recordset.Delete
            Data1.Refresh
            MsgBox "Record Deleted", 16, "Notification Only"
            Tbl.Close
            DB.Close
        Else
            MsgBox "Code in use on race code " & Tbl("code_race") & " " & Tbl("reldate"), 16, "Cannot DELETE"
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
        Unload FMracepnt
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
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "points"
    Data1.Refresh
    
    ButtAdd.Enabled = False
    ButtDelete.Enabled = False
    ButtUpdate.Enabled = True

    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True
End Sub

