VERSION 2.00
Begin Form FMcolor 
   Caption         =   "Bird Colour Maintenance/Enquiry"
   ClientHeight    =   4800
   ClientLeft      =   1290
   ClientTop       =   1650
   ClientWidth     =   7905
   Height          =   5205
   Left            =   1230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7905
   Top             =   1305
   Width           =   8025
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4440
      Width           =   7935
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7935
      Begin SSCommand ButtDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   9
         Top             =   3120
         Width           =   975
      End
      Begin SSCommand ButtAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6720
         TabIndex        =   2
         Top             =   2400
         Width           =   975
      End
      Begin SSCommand ButtUpdate 
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6720
         TabIndex        =   1
         Top             =   1680
         Width           =   975
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6720
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin SSCommand ButtOK 
         Caption         =   "&Ok"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6720
         TabIndex        =   4
         Top             =   3720
         Width           =   975
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Screen selections - Colours"
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
         Left            =   120
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
         Top             =   480
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
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "colour"
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H00FFFF80&
         DataSource      =   "Data1"
         Editable        =   0   'False
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
         Height          =   3495
         HorzLines       =   0  'None
         Layout          =   FMCOLOR.FRX:0000
         LayoutIndex     =   1
         Left            =   600
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   0
         Top             =   720
         UseBookmarks    =   -1  'True
         Width           =   5775
         WrapCellPointer =   0   'False
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Colour Codes"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
End
Option Explicit

Sub ButtAdd_Click ()
    FMaddcolour.Show 1
End Sub

Sub ButtCancel_Click ()
    If ButtUpdate.Enabled Then
        'Exit form
        Unload FMcolor
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
    
    Msg$ = "Are you sure you want to delete Colour " & Key$ & " " & Key1$ & " ?"
    Response = MsgBox(Msg$, 36, "Delete Colour")
    If Response = IDYES Then
        Screen.MousePointer = 11
        'Ensure NOT used by a bird record, retain db integrity
        Set DB = OpenDatabase(DBFullFileName$)    'Open Database
        Set Tbl = DB.OpenTable("master")     'Open Table
        Tbl.Index = "code_colour"                 'Open Index
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
        Unload FMcolor
    End If
End Sub

Sub ButtUpdate_Click ()
    Table1.Enabled = True
    Table1.Editable = True
    ButtAdd.Enabled = True
    ButtDelete.Enabled = True
    ButtUpdate.Enabled = False
End Sub

Sub cmdPrint_Click ()
Report1.Action = 1
DoEvents
End Sub

Sub Form_Load ()
    'Set DBFile = OpenDatabase(DBFullFileName$)
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "colour"
    Data1.Refresh
    
    ButtAdd.Enabled = False
    ButtDelete.Enabled = False
    ButtUpdate.Enabled = True

    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True
    
End Sub

Sub Form_Unload (Cancel As Integer)
    Unload FMaddcolour
End Sub

