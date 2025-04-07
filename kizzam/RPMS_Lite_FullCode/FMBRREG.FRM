VERSION 2.00
Begin Form FMbrreg 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Breeding Register"
   ClientHeight    =   8370
   ClientLeft      =   420
   ClientTop       =   1605
   ClientWidth     =   11220
   Height          =   8880
   Left            =   360
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   11220
   Top             =   1155
   Width           =   11340
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11175
      Begin SSCommand butt3dProdgAll 
         Caption         =   "Prodgeny ALL"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   37
         Top             =   7200
         Width           =   1695
      End
      Begin SSCommand ButtPtsFlwn 
         Caption         =   "&Points Flown"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1200
         TabIndex        =   26
         Top             =   7200
         Width           =   1695
      End
      Begin SSCommand ButtFind 
         Caption         =   "&Find"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9720
         TabIndex        =   22
         Top             =   4200
         Width           =   1215
      End
      Begin SSCommand ButtCopy 
         Caption         =   "&Copy"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9720
         TabIndex        =   21
         Top             =   3480
         Width           =   1215
      End
      Begin SSCommand ButtAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9720
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin SSCommand ButtProcess 
         Caption         =   "&Process"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9720
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   9720
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSCommand Butt3DPedg 
         Caption         =   "&Pedigree"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   7200
         TabIndex        =   24
         Top             =   7200
         Width           =   1695
      End
      Begin TextBox Task 
         Height          =   285
         Left            =   10440
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   6960
         Visible         =   0   'False
         Width           =   615
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Order by"
         Font3D          =   3  'Inset w/light shading
         Height          =   495
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   35
         Top             =   120
         Width           =   9375
         Begin SSOption ByYrMarkNo 
            Caption         =   "Yr-Mark-No"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin SSOption ByMarkNoYr 
            Caption         =   "Mark-No-Yr"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2040
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
         Begin SSOption ByNoYrMark 
            Caption         =   "No-Yr-Mark"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3480
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
         Begin SSOption BySire 
            Caption         =   "Sire"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4920
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin SSOption ByDam 
            Caption         =   "Dam"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5880
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin SSOption ByDOB 
            Caption         =   "DOB"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6840
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Show only"
         Font3D          =   3  'Inset w/light shading
         Height          =   1095
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   34
         Top             =   720
         Width           =   9375
         Begin SSCheck ShowAll1 
            Caption         =   "Select All"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin SSCheck ShowCurrent 
            Caption         =   "Current"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin SSCheck ShowRace 
            Caption         =   "Race"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   9
            Top             =   480
            Width           =   735
         End
         Begin SSCheck ShowBP1 
            Caption         =   "B/Plate A"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2280
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin SSCheck ShowStock 
            Caption         =   "Stock"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   10
            Top             =   720
            Width           =   855
         End
         Begin SSCheck ShowRef 
            Caption         =   "Reference"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin SSCheck ShowBP2 
            Caption         =   "B/Plate B"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2280
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin SSCheck ShowSPb 
            Caption         =   "Special 2"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3480
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
         Begin SSCheck ShowSPa 
            Caption         =   "Special 1"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3480
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin SSOption ShowCocks 
            Caption         =   "Cocks"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5160
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin SSOption ShowHens 
            Caption         =   "Hens"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5160
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Width           =   735
         End
         Begin SSOption ShowUnSexed 
            Caption         =   "UnSexed"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5160
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   600
            Width           =   1095
         End
         Begin SSOption ShowAllSex 
            Caption         =   "ALL Sexes"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5160
            TabIndex        =   16
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMBRREG.FRX:0000
            Contents        =   FMBRREG.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   6360
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   17
            Text            =   "ALL"
            TextDisplay     =   0  'first visible
            Top             =   240
            ValueCol        =   -1
            Width           =   1095
         End
      End
      Begin TextBox RefreshData 
         Height          =   285
         Left            =   9840
         TabIndex        =   33
         Text            =   "N"
         Top             =   7080
         Visible         =   0   'False
         Width           =   495
      End
      Begin SSCommand But3Prodg 
         Caption         =   "Prodgeny"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3240
         TabIndex        =   25
         Top             =   7200
         Width           =   1695
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
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
         Height          =   5055
         HorzLines       =   0  'None
         Layout          =   FMBRREG.FRX:0184
         LayoutIndex     =   1
         Left            =   120
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   18
         Top             =   1920
         UseBookmarks    =   -1  'True
         Width           =   9495
         WrapCellPointer =   0   'False
      End
      Begin SSCommand CmdPrint 
         Caption         =   "P&rint"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   9720
         TabIndex        =   23
         Top             =   5040
         Width           =   1215
      End
      Begin CommonDialog CMDialog1 
         Left            =   9840
         Top             =   5880
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Report - Screen Selections"
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
         Left            =   10440
         MarginBottom    =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginTop       =   0
         PrinterCollation=   1  'collated
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
         ReportTitle     =   "Report from screen selection"
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   5880
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
         WindowTitle     =   "Report - Selected records"
         WindowTop       =   100
         WindowWidth     =   480
      End
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   31
      Top             =   7920
      Width           =   2535
   End
   Begin SSPanel MsgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   7920
      Width           =   8655
   End
End
Option Explicit
Dim fgCntBirds As Long

Sub But3Prodg_Click ()
    Dim ii As Integer
    'if no record selected display message
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Bird(s) selected", 0, "ERROR - Unable to Process"
    Else
	' Process each selected row.
	For ii% = 0 To Table1.BookmarkCount - 1
	    data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	    FMprodg.FYr.Text = data1.Recordset("yr")
	    FMprodg.FMark.Text = data1.Recordset("mark")
	    FMprodg.Fring.Text = data1.Recordset("ringno")
	    FMprodg.Fsex.Text = data1.Recordset("codesex")
	    If data1.Recordset("codesex") = "U" Then
		MsgBox "Birds Sex has not been defined", 0, "ERROR - Unable to Process Prodgeny"
	    Else
		FMprodg.Show 1
	    End If
	Next ii%
	Table1.BookmarkCount = 0
    End If
End Sub

Sub Butt3DPedg_Click ()

    Dim ii As Integer

    'if no record selected display message
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Bird(s) selected", 0, "ERROR - Unable to Process"
    Else
	' Process each selected row.
	For ii% = 0 To Table1.BookmarkCount - 1
	    data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	    FMpedg.FYr.Text = data1.Recordset("yr")
	    FMpedg.FMark.Text = data1.Recordset("mark")
	    FMpedg.Fring.Text = data1.Recordset("ringno")
	    FMpedg.Show 1
	Next ii%
	Table1.BookmarkCount = 0
	'Data1.Refresh
    End If
End Sub

Sub butt3dProdgAll_Click ()
    Dim ii As Integer
    'if no record selected display message
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Bird(s) selected", 0, "ERROR - Unable to Process"
    Else
	' Process each selected row.
	For ii% = 0 To Table1.BookmarkCount - 1
	    data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	    FMprodgALL.FYr.Text = data1.Recordset("yr")
	    FMprodgALL.FMark.Text = data1.Recordset("mark")
	    FMprodgALL.FRingNo.Text = data1.Recordset("ringno")
	    FMprodgALL.Fsex.Text = data1.Recordset("codesex")
	    If data1.Recordset("codesex") = "U" Then
		MsgBox "Birds Sex has not been defined", 0, "ERROR - Unable to Process Prodgeny"
	    Else
		FMprodgALL.Show 1
	    End If
	Next ii%
	Table1.BookmarkCount = 0
    End If

End Sub

Sub ButtAdd_Click ()
On Error GoTo FMErr2a

'Check if in Demo mode, limit no of birds allow entry.
If LicNo(Trim(GLOwner$)) = "2223190N" Then '"262469J"
    If fgCntBirds& > 45 Then 'data1.Recordset.RecordCount
	MsgBox "Demo mode limits the number of bird entries to 45 birds, this gives you enough to get a feel for the software." & Chr(10) & "To unlock software, print order form from the Main menu." & Chr(10) & "and send with remittance to supplier. Manuals,disks and key to unlock software will be sent to you.", 64, "This software is running in Demonstration mode for evaluation purposes"
	Exit Sub
    End If
End If

' Advise the next form we are in add mode.
    If GLLfmbird% = False Then 'Load form if not loaded
	Load FMbird
    End If

    FMbird.Task.Text = "A"
    FMbird.Show 1
    data1.Refresh '29/8
    Exit Sub

FMErr2a:
    Select Case Err
    Case 3021
	MsgBox "Database Empty", 16, "INFORMATION ONLY"
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

End Sub

Sub ButtCancel_Click ()
    Unload FMbrreg
    If GLLfmbird% = True Then 'only unload if loaded
	Unload FMbird
    End If
    If GLLfmerror% = True Then 'only unload if loaded
	Unload FMerror
    End If
End Sub

Sub ButtCopy_Click ()
    ' Advise the next form we are in add mode.
    If GLLfmbird% = False Then 'Load form if not loaded
	Load FMbird
    End If
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Bird selected to copy details from", 0, "ERROR - Unable to Create new bird record"
    ElseIf Table1.BookmarkCount > 1 Then
	MsgBox "More than one bird selected to copy details from", 0, "ERROR - Unable to Create new bird record"
    Else
	' Process selected row.
	DoEvents
	data1.Recordset.Bookmark = Table1.BookmarkList(0)
	FMbird.Key_yr.Text = data1.Recordset("yr")
	FMbird.Key_mark.Text = data1.Recordset("mark")
	FMbird.Key_ring_no.Text = data1.Recordset("ringno")
	FMbird.Task.Text = "C" 'Copy
	FMbird.Show 1
	DoEvents
	data1.Refresh
    End If
    'Table1.BookmarkCount = 0
    
End Sub

Sub ButtFind_Click ()
  Dim Key As String
  Dim Myr As Integer
  Dim Mmark As String
  Dim Mring As String
  Dim Mcnt As Integer
  Dim Ma As Integer
  Dim Mkey1 As String

    'bit of work to do here, composite key Yr/Mark/RingNo.
    Key$ = InputBox("Enter Yr/Mark/Ring Number to find :", "Find Pigeon")
    Key$ = UCase(Key$)
    Mcnt% = Len(Key$)
    If Mcnt% < 1 Then  'Nothing entered
	Exit Sub
    End If
    Ma% = InStr(Key$, "/") ' get position of /
    
    If Ma% = 0 Then 'must have entered year only
	Myr% = Val(Key$)
	data1.Recordset.FindFirst "yr >= " & Myr
	Exit Sub
    End If

    Mkey1$ = Right(Key$, Len(Key$) - (Ma%))
    Myr% = Val(Left(Key$, Ma% - 1))
    Ma% = InStr(Mkey1$, "/") ' get position of /
    If Ma% = 0 Then 'must have entered year & Mark only
	Mmark$ = Mkey1$
	data1.Recordset.FindFirst "yr >= " & Myr & "mark >= '" & Mmark & "'"
	Exit Sub
    End If
    Mmark$ = Left(Mkey1$, Ma% - 1)
    Mring$ = Right(Mkey1$, Len(Mkey1$) - (Ma%))

    If Trim(Key$) <> "" Then
	data1.Recordset.FindFirst "yr = " & Myr% & " AND mark = '" & Mmark$ & "' AND ringno = '" & Mring$ & "'"
	If data1.Recordset.NoMatch Then  ' Check if record is found.
	    MsgBox "Bird Record Not in currently selected 'show only' list", 0, "Not Found"
	Else
	    MsgBox "Bird Record on File", 0, "FOUND"
	End If
    End If

    'If Trim(Key$) <> "" Then
	'If ByYrMarkNo Then
	    'OrderBy$ = "yr DESC, mark DESC, Val(ringno) DESC"
	'ElseIf ByMarkNoYr Then
	    'OrderBy$ = "mark DESC, Val(ringno) DESC,yr DESC"
	'ElseIf ByNoYrMark Then
	    'OrderBy$ = "Val(ringno) DESC, yr DESC, mark DESC"
	'ElseIf BySire Then
	    'OrderBy$ = "sireyr, siremark, sireringno"
	'ElseIf ByDam Then
	    'OrderBy$ = "damyr, dammark, damringno"
	'ElseIf ByDOB Then
	    'OrderBy$ = "Val(dob_li) DESC, yr DESC, mark DESC, Val(ringno) DESC"
	'Else
	    'OrderBy$ = "yr DESC, mark DESC, ringno DESC"
	'End If
    'End If

End Sub

Sub ButtProcess_Click ()
On Error GoTo FMerr4a
    Dim ii As Integer

    'if no record selected display message
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Bird(s) selected", 0, "ERROR - Unable to Process"
    Else
	' Process each selected row.
	For ii% = 0 To Table1.BookmarkCount - 1
	    DoEvents
	    data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	    FMbird.Task.Text = Task.Text
	    FMbird.Key_yr.Text = data1.Recordset("yr")
	    FMbird.Key_mark.Text = data1.Recordset("mark")
	    FMbird.Key_ring_no.Text = data1.Recordset("ringno")
	    FMbird.Show 1
	Next ii%
    End If
    Table1.BookmarkCount = 0
    
    DoEvents
    
    'refresh data1 if records deleted
    If RefreshData.Text = "Y" Then
	data1.Refresh
	DoEvents
	RefreshData.Text = "N"
    End If
    Exit Sub
FMerr4a:
    Select Case Err
    Case 3021
	MsgBox "Database Low", 16, "Continue to add birds"
	Resume Next
	'Unload Me
	Exit Sub
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

End Sub

Sub ButtPtsFlwn_Click ()
Dim ii As Integer
    For ii% = 0 To Table1.BookmarkCount - 1
	data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	FMpfbb.scrYr.Text = data1.Recordset("yr")
	FMpfbb.scrMark.Text = data1.Recordset("mark")
	FMpfbb.scrRingno.Text = data1.Recordset("ringno")
	FMpfbb.Show 1
    Next ii%
    Table1.BookmarkCount = 0
End Sub

Sub ByDam_Click (Value As Integer)
    setselections
End Sub

Sub byDOB_Click (Value As Integer)
    setselections
End Sub

Sub ByMarkNoYr_Click (Value As Integer)
setselections
End Sub

Sub ByNoYrMark_Click (Value As Integer)
setselections
End Sub

Sub BySire_Click (Value As Integer)
setselections
End Sub

Sub byYrMarkNo_Click (Value As Integer)
    setselections
End Sub

Sub cmdPrint_Click ()
screen.MousePointer = 11
Report1.Action = 1
DoEvents
screen.MousePointer = 0
End Sub

Sub CSCombo1_Click ()
    setselections
End Sub

Sub Form_Activate ()
On Error GoTo FMerr1a
    ' Set up the form for the current mode.
    'FMbird.Task.Text = Task.Text
    If Task.Text = "E" Then
	ButtAdd.Visible = False
    Else
	ButtAdd.Visible = True
    End If
Exit Sub
FMerr1a:
    Select Case Err
    Case 3021
	MsgBox "Database Empty", 16, "INFORMATION ONLY"
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

End Sub

Sub Form_Load ()
    On Error GoTo FMErr1:
    screen.MousePointer = 11
    data1.DatabaseName = DBFullFileName$
    data1.RecordSource = "master"
    'Data1.RecordSource = "SELECT * FROM master WHERE 'indsoh = True' ORDER BY 'yr DESC, mark DESC, Val(ringno) DESC'"
    setselections
    RefreshData.Text = "N" 'Default to no refresh of Data
    screen.MousePointer = 0
    Exit Sub

FMErr1:
    Select Case Err
    Case 3021
	MsgBox "Database Empty", 16, "INFORMATION ONLY"
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
 

End Sub

Function OrderBy () As String
    If ByYrMarkNo Then
       OrderBy$ = "yr DESC, mark DESC, Val(ringno) DESC"
    ElseIf ByMarkNoYr Then
	OrderBy$ = "mark DESC, Val(ringno) DESC,yr DESC"
    ElseIf ByNoYrMark Then
	OrderBy$ = "Val(ringno) DESC, yr DESC, mark DESC"
    ElseIf BySire Then
	OrderBy$ = "sireyr, siremark, sireringno"
    ElseIf ByDam Then
	OrderBy$ = "damyr, dammark, damringno"
    ElseIf ByDOB Then
	OrderBy$ = "Val(dob_li) DESC, yr DESC, mark DESC, Val(ringno) DESC"
    Else
       OrderBy$ = "yr DESC, mark DESC, ringno DESC"
    End If
End Function

Sub setselections ()
    Dim s As String
    Dim o As String

    On Error GoTo ByErr9

    s$ = ShowWhat()
    o$ = OrderBy()

    If s$ = "" Then
	data1.RecordSource = "SELECT * FROM master ORDER BY " & o$
    Else
	data1.RecordSource = "SELECT * FROM master WHERE " & s$ & " ORDER BY " & o$
    End If
    data1.Refresh
    data1.Recordset.MoveLast
    fgCntBirds& = data1.Recordset.RecordCount
    data1.Recordset.MoveFirst
    MsgPanel.Caption = "Count of birds listed = " & Str(fgCntBirds&)
    Exit Sub
    
ByErr9:
    If Err = 380 Then
	Resume Next
    ElseIf Err = 3021 Then
	MsgBox "No Birds on file meet this criteria"
	Exit Sub
    Else
	MsgBox "Error :" & Err & " " & Error$, 16, "Error"
	Exit Sub
    End If
End Sub

Sub ShowAll1_Click (Value As Integer)
If ShowAll1 Then
    ShowCurrent.Visible = False
    ShowRef.Visible = False
    ShowRace.Visible = False
    ShowStock.Visible = False
    ShowBP1.Visible = False
    ShowBP2.Visible = False
    ShowSPa.Visible = False
    ShowSPb.Visible = False
Else
    ShowCurrent.Visible = True
    ShowRef.Visible = True
    ShowRace.Visible = True
    ShowStock.Visible = True
    ShowBP1.Visible = True
    ShowBP2.Visible = True
    ShowSPa.Visible = True
    ShowSPb.Visible = True
End If
setselections
End Sub

Sub ShowAllSex_Click (Value As Integer)
    setselections
End Sub

Sub ShowBP1_Click (Value As Integer)
    setselections
End Sub

Sub ShowBP2_Click (Value As Integer)
    setselections
End Sub

Sub ShowCocks_Click (Value As Integer)
    setselections
End Sub

Sub showCurrent_Click (Value As Integer)
    setselections
End Sub

Sub ShowHens_Click (Value As Integer)
    setselections
End Sub

Sub showRace_Click (Value As Integer)
    setselections
End Sub

Sub showRef_Click (Value As Integer)
    setselections
End Sub

Sub ShowSPa_Click (Value As Integer)
    setselections
End Sub

Sub ShowSPb_Click (Value As Integer)
    setselections
End Sub

Sub showStock_Click (Value As Integer)
    setselections
End Sub

Sub ShowUnSexed_Click (Value As Integer)
    setselections
End Sub

Function ShowWhat () As String

    Dim s As String

    If ShowAll1 Then
	s$ = ""
    Else
	If ShowCurrent Then
	    s$ = "indsoh = True"
	Else
	    s$ = "indsoh = False"
	End If

	If ShowRef Then
	    If s$ = "" Then
		s$ = "indref = True"
	    Else
		s$ = s$ & " AND indref = True"
	    End If
	End If
	If ShowRace Then
	    If s$ = "" Then
		s$ = "indrace = True"
	    Else
		s$ = s$ & " AND indrace = True"
	    End If
	End If
	If ShowStock Then
	    If s$ = "" Then
		s$ = "indstock = True"
	    Else
		s$ = s$ & " AND indstock = True"
	    End If
	'Else
	'   If s$ = "" Then
	'    s$ = "indstock = False"
	'   Else
	'    s$ = s$ & " AND indstock = False"
	'   End If
	End If
	If ShowSPb Then
	    If s$ = "" Then
		s$ = "indsp2 = True"
	    Else
		s$ = s$ & " AND indsp2 = True"
	    End If
	End If
	
	If ShowSPa Then
	    If s$ = "" Then
		s$ = "indsp1 = True"
	    Else
		s$ = s$ & " AND indsp1 = True"
	    End If
	End If
	If ShowBP2 Then
	    If s$ = "" Then
		s$ = "indbpb = True"
	    Else
		s$ = s$ & " AND indbpb = True"
	    End If
	End If
	If ShowBP1 Then
	    If s$ = "" Then
		s$ = "indbpa = True"
	    Else
		s$ = s$ & " AND indbpa = True"
	    End If
	End If
    End If

    If ShowHens Then
       If s$ = "" Then
	s$ = "codesex = 'H'"
       Else
	s$ = s$ & " AND codesex = 'H'"
       End If
    ElseIf ShowCocks Then
       If s$ = "" Then
	s$ = "codesex = 'C'"
       Else
	s$ = s$ & " AND codesex = 'C'"
       End If
    ElseIf ShowUnSexed Then
       If s$ = "" Then
	s$ = "codesex = 'U'"
       Else
	s$ = s$ & " AND codesex = 'U'"
       End If
    End If
    If CSCombo1.Text <> "ALL" Then
	If CSCombo1.Value = "PRE80" Then
	    If s$ = "" Then
		s$ = "yr < 1980"
	    Else
		s$ = s$ & " AND yr < 1980"
	    End If
	Else
	    If s$ = "" Then
		s$ = "yr = " & Val(CSCombo1.Value)
	    Else
		s$ = s$ & " AND yr = " & Val(CSCombo1.Value)
	    End If
	End If
    End If

    ShowWhat$ = s$
    
End Function

