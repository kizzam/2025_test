VERSION 2.00
Begin Form FMeyesgn 
   Caption         =   "Eye Sign Maintenance/Enquiry"
   ClientHeight    =   4950
   ClientLeft      =   1320
   ClientTop       =   1350
   ClientWidth     =   7815
   Height          =   5355
   Left            =   1260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7815
   Top             =   1005
   Width           =   7935
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4560
      Width           =   7815
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      Begin SSCommand ButtDelete 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         RoundedCorners  =   0   'False
         TabIndex        =   5
         Top             =   3240
         Width           =   975
      End
      Begin SSCommand ButtAdd 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         RoundedCorners  =   0   'False
         TabIndex        =   2
         Top             =   2400
         Width           =   975
      End
      Begin SSCommand ButtUpdate 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         RoundedCorners  =   0   'False
         TabIndex        =   1
         Top             =   1680
         Width           =   975
      End
      Begin SSCommand ButtCancel 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         RoundedCorners  =   0   'False
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin SSCommand ButtOK 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Ok"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         RoundedCorners  =   0   'False
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "eyesign"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H00FFFF80&
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeadingHeight   =   1.5
         Height          =   2775
         HorzLines       =   0  'None
         Layout          =   FMEYESGN.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   0
         Top             =   1200
         UseBookmarks    =   -1  'True
         Width           =   6255
         WrapCellPointer =   0   'False
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eye Sign"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
End
Option Explicit

Sub ButtAdd_Click ()
    FMaddeyesgn.Show 1
End Sub

Sub ButtCancel_Click ()
    If ButtUpdate.Enabled Then
        'Exit form
        Unload FMeyesgn
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
        'Ensure NOT used by a bird record, retain db integrity
        Set DB = OpenDatabase(DBFullFileName$)    'Open Database
        Set Tbl = DB.OpenTable("master")     'Open Table
        Tbl.Index = "code_eye"                 'Open Index
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
        Unload FMeyesgn
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
    Data1.RecordSource = "eyesign"
    Data1.Refresh
    
    ButtAdd.Enabled = False
    ButtDelete.Enabled = False
    ButtUpdate.Enabled = True

    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True

End Sub

Sub Form_Unload (Cancel As Integer)
    Unload FMaddeyesgn
    'DBFile.Close
End Sub

