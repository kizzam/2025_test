VERSION 2.00
Begin Form FMlstratings 
   Caption         =   "Rating Maintenance"
   ClientHeight    =   6000
   ClientLeft      =   1035
   ClientTop       =   585
   ClientWidth     =   7545
   Height          =   6405
   Left            =   975
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   7545
   Top             =   240
   Width           =   7665
   Begin SSPanel MsgBox2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin SSPanel MsgBox1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5640
      Width           =   2655
   End
   Begin SSPanel MsgBox 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   4935
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin CrystalReport Report1 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Screen Selections - Ratings"
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data1"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   -1  'True
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   3960
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
         Top             =   5160
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
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   4560
         Width           =   1095
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
         HeadingHeight   =   1
         Height          =   4215
         HorzLines       =   0  'None
         Layout          =   FMLSTRAT.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   9
         Top             =   720
         UseBookmarks    =   -1  'True
         Width           =   5775
         WrapCellPointer =   0   'False
      End
      Begin SSCommand CmdDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6240
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin SSCommand CmdUpdate 
         Caption         =   "&Update"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6240
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin SSCommand CmdAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6240
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin SSCommand CmdCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6240
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin SSCommand CmdOK 
         Caption         =   "&OK"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
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
         RecordSource    =   "Ratings"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ratings"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
   End
End
Option Explicit

Sub CmdAdd_Click ()
    FMaddrating.Show 1
End Sub

Sub CmdCancel_Click ()
    If CmdUpdate.Enabled Then
        'Exit form
        Unload FMlstratings
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
    Dim tbl1 As snapshot
    
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
        'check to see if bird is used in current Race Results
        Screen.MousePointer = 11
        Set DB = OpenDatabase(DBFullFileName$)
        Set tbl1 = DB.CreateSnapshot("racetrans")     'Open Table
        tbl1.FindFirst "CodeRating = '" & Key$ & "'"
        If tbl1.NoMatch Then
            Data1.Recordset.Delete
            Data1.Refresh
            MsgBox "Record Deleted", 16, "Notification Only"
        Else
            MsgBox "Cannot delete while RATING is used in results for race code - '" & Str(tbl1("RaceCode")) & "' Bird '" & tbl1("ringno") & "'", 16, "Warning - Cannot Delete record"
        End If
        
        tbl1.Close
        DB.Close
    End If

    Screen.MousePointer = 0
    'CmdCancel.Default = True
    Set tbl1 = Nothing
    Set DB = Nothing

End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdOK_Click ()
    If Table1.Editable = True Then
        'Update table
        Table1.Refresh
    Else
        'Exit form
        Unload FMlstratings
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
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "Ratings"
    Data1.Refresh
    
    CmdAdd.Enabled = False
    CmdDelete.Enabled = False
    CmdUpdate.Enabled = True

    Table1.Enabled = True
    Table1.Editable = False
    Table1.Active = True
    
   'data1.Recordset.MoveFirst
End Sub

