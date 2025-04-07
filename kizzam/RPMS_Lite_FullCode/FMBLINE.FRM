VERSION 2.00
Begin Form FMBLine 
   Caption         =   "List Bloodlines"
   ClientHeight    =   4410
   ClientLeft      =   915
   ClientTop       =   1695
   ClientWidth     =   8040
   Height          =   4815
   Left            =   855
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   8040
   Top             =   1350
   Width           =   8160
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   8055
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Top             =   3600
         Width           =   975
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Screen Selections - Bloodlines"
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data1"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   240
         MarginBottom    =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginTop       =   0
         PrinterCollation=   2  'default
         PrinterCopies   =   1
         PrinterDriver   =   ""
         PrinterName     =   ""
         PrinterPort     =   ""
         PrinterStartPage=   0
         PrinterStopPage =   0
         PrintFileCharSepQuote=   ""
         PrintFileCharSepSeparator=   ""
         PrintFileName   =   ""
         PrintFileType   =   2  'Text
         PrintFileUseRptDateFmt=   0   'False
         PrintFileUseRptNumberFmt=   0   'False
         ReportFileName  =   ""
         ReportSource    =   1  'Bound TrueGrid Control
         ReportTitle     =   ""
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   3600
         UserName        =   ""
         WindowBorderStyle=   2  'Sizable
         WindowControlBox=   -1  'True
         WindowControls  =   -1  'True
         WindowHeight    =   300
         WindowLeft      =   100
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowParentHandle=   0
         WindowState     =   0  'normal
         WindowTitle     =   ""
         WindowTop       =   100
         WindowWidth     =   480
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
         HeadingHeight   =   1
         Height          =   2895
         HorzLines       =   0  'None
         Layout          =   FMBLINE.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   7
         Top             =   840
         UseBookmarks    =   -1  'True
         Width           =   6495
         WrapCellPointer =   0   'False
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "BloodLine"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin SSCommand CmdAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6840
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin SSCommand CmdUpdate 
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6840
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin SSCommand CmdDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   3000
         Width           =   975
      End
      Begin SSCommand CmdCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6840
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin SSCommand CmdOk 
         Caption         =   "&Ok"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6840
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bloodlines"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
   End
End
Option Explicit

Sub CmdAdd_Click ()
    FMaddBLine.Show 1
End Sub

Sub CmdCancel_Click ()
    If CmdUpdate.Enabled Then
        'Exit form
        Unload FMBLine
    Else
        'Disable Update
        CmdUpdate.Enabled = True
        CmdAdd.Enabled = False
        CmdDelete.Enabled = False
        Table1.Editable = False
        Table1.Refresh
    End If


End Sub

Sub CmdDelete_Click ()

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
    
    Msg$ = "Are you sure you want to DELETE " & Key$ & " " & Key1$ & " ?"
    Response = MsgBox(Msg$, 36, "Delete Record")
    If Response = IDYES Then
        Screen.MousePointer = 11
        'Ensure NOT used by a bird record, retain db integrity
        Set DB = OpenDatabase(DBFullFileName$)    'Open Database
        Set Tbl = DB.OpenTable("master")     'Open Table
        Tbl.Index = "bloodline"              'Open Index
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
    'CmdCancel.Default = True

End Sub

Sub cmdOK_Click ()
    If Table1.Editable = True Then
        'Update table
        Table1.Refresh
    Else
        'Exit form
        Unload FMBLine
    End If

End Sub

Sub cmdPrint_Click ()
Report1.Action = 1
DoEvents
End Sub

Sub CmdUpdate_click ()
    Table1.Enabled = True
    Table1.Editable = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = True
    CmdUpdate.Enabled = False
End Sub

Sub Form_Load ()
    'Set DBFile = OpenDatabase(DBFullFileName$)
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "bloodline"
    Data1.Refresh
    CmdAdd.Enabled = False
    CmdDelete.Enabled = False
    CmdUpdate.Enabled = True
    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True

End Sub

