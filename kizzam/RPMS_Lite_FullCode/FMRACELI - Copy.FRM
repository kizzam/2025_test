VERSION 2.00
Begin Form FMracelist 
   Caption         =   "Race Nominations"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9465
   Height          =   7140
   Left            =   0
   LinkTopic       =   "Form6"
   ScaleHeight     =   6735
   ScaleWidth      =   9465
   Top             =   15
   Width           =   9585
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6360
      Width           =   9495
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9495
      Begin SSCommand ButtInclude 
         Caption         =   ">>>"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   615
         Left            =   4680
         TabIndex        =   6
         Top             =   2640
         Width           =   615
      End
      Begin SSCommand ButtExclude 
         Caption         =   "<<<"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   615
         Left            =   4680
         TabIndex        =   8
         Top             =   3480
         Width           =   615
      End
      Begin SSCommand ButtDisplayFlown 
         Caption         =   "Display Points &Flown"
         Font3D          =   3  'Inset w/light shading
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   5160
         Width           =   2055
      End
      Begin SSCommand ButtDisplayPdg 
         Caption         =   "&Display Pedigree"
         Font3D          =   3  'Inset w/light shading
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   5760
         Width           =   2055
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         Height          =   975
         Left            =   8040
         TabIndex        =   12
         Top             =   5160
         Width           =   1095
      End
      Begin SSCommand cmdTEAMS 
         Caption         =   "Select entries from &TEAMS"
         Font3D          =   3  'Inset w/light shading
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   5160
         Width           =   2535
      End
      Begin SSFrame FrameSelect 
         Caption         =   "Group Select of birds"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   31
         Top             =   4320
         Width           =   4335
         Begin SSOption indNotLost 
            Caption         =   "Status <> ""L"""
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2640
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin SSOption indNone 
            Caption         =   "None"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin SSOption indAll 
            Caption         =   "All Birds Listed"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
      End
      Begin TextBox mRaceStatus 
         Height          =   375
         Left            =   7680
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin SSCommand CmdIncludeTeam 
         Caption         =   ">>>"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4680
         TabIndex        =   29
         Top             =   2640
         Width           =   615
      End
      Begin SSCommand CmdMaster 
         Caption         =   "Select from &MASTER list"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   5760
         Width           =   2535
      End
      Begin SSPanel TeamBirds 
         BevelOuter      =   0  'None
         Caption         =   "Team Birds"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   4215
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMRACELI.FRX:0000
         Contents        =   FMRACELI.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   1320
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "description,code"
         ListDataSource  =   "Data4"
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   2
         Text            =   "CSCombo1"
         TextDisplay     =   0  'first visible
         Top             =   1560
         ValueCol        =   1
         Width           =   3255
      End
      Begin TrueGrid Table3 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data3"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   2175
         HorzLines       =   0  'None
         Layout          =   FMRACELI.FRX:0042
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   3
         Top             =   2160
         UseBookmarks    =   -1  'True
         Width           =   4335
         WrapCellPointer =   0   'False
      End
      Begin Data Data4 
         Caption         =   "Data4"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "teams"
         Top             =   960
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Data Data3 
         Caption         =   "Data3"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "teambird"
         Top             =   720
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSCommand ButtPrint 
         Caption         =   "&Print Race Sheet"
         Font3D          =   3  'Inset w/light shading
         Height          =   495
         Left            =   5640
         TabIndex        =   11
         Top             =   5160
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         Caption         =   "Code:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin TextBox mRaceCode 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin TextBox mPntName 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   720
         Width           =   3135
      End
      Begin SSPanel RaceEntries 
         BevelOuter      =   0  'None
         Caption         =   "Race Entries"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   1440
         Width           =   3735
      End
      Begin SSPanel RaceBirds 
         BevelOuter      =   0  'None
         Caption         =   "Current Race Birds"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   4215
      End
      Begin SSPanel Panel3D5 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         Caption         =   "Race Point:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin SSPanel Panel3D4 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         Caption         =   "Race Type:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin SSPanel Panel3D3 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         Caption         =   "Date of Race:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin TrueGrid Table2 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data2"
         Editable        =   0   'False
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   3375
         HorzLines       =   0  'None
         Layout          =   FMRACELI.FRX:01CA
         LayoutIndex     =   1
         Left            =   5400
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   7
         Top             =   1680
         UseBookmarks    =   -1  'True
         Width           =   3855
         WrapCellPointer =   0   'False
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data1"
         Editable        =   0   'False
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   3375
         HorzLines       =   0  'None
         Layout          =   FMRACELI.FRX:05A0
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   18
         Top             =   1680
         UseBookmarks    =   -1  'True
         Width           =   4215
         WrapCellPointer =   0   'False
      End
      Begin Data Data2 
         Caption         =   "Data2"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7920
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "racetrans"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7920
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TextBox Mcoderace 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   240
         Width           =   1215
      End
      Begin TextBox Mcodepnt 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   3240
         MaxLength       =   5
         MousePointer    =   8  'Size NW SE
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   720
         Width           =   975
      End
      Begin TextBox Mracedate 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Team:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
   End
End
Option Explicit
Dim tblPoint As Table
Dim DB As Database
Dim GLdist As Integer
Dim nCntEnt As Integer
Dim nCntteam As Long

Sub ButtCancel_Click ()
    Unload FMracelist
End Sub

Sub ButtCancel_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click to Exit birds nominations/entry for race"
End Sub

Sub ButtDisplayFlown_Click ()
Dim ii As Integer
    If Table3.Visible = True Then
        For ii% = 0 To Table3.BookmarkCount - 1
            If ii% > 0 Then
                Exit For
            End If
            Data3.Recordset.Bookmark = Table3.BookmarkList(ii%)
            FMpfbb.scrYr.Text = Data3.Recordset("yr")
            FMpfbb.scrMark.Text = Data3.Recordset("mark")
            FMpfbb.scrRingno.Text = Data3.Recordset("ringno")
            FMpfbb.Show 1
        Next ii%
        Table3.BookmarkCount = 0
    ElseIf Table1.Visible = True Then
        For ii% = 0 To Table1.BookmarkCount - 1
            If ii% > 0 Then
                Exit For
            End If
            data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
            FMpfbb.scrYr.Text = data1.Recordset("yr")
            FMpfbb.scrMark.Text = data1.Recordset("mark")
            FMpfbb.scrRingno.Text = data1.Recordset("ringno")
            FMpfbb.Show 1
        Next ii%
        Table1.BookmarkCount = 0
    End If

    ' Process each selected row.
    For ii% = 0 To table2.BookmarkCount - 1
        'Workaround - only allow 2 selections
        If ii% > 0 Then
            Exit For
        End If
        data2.Recordset.Bookmark = table2.BookmarkList(ii%)
        FMpfbb.scrYr.Text = data2.Recordset("yr")
        FMpfbb.scrMark.Text = data2.Recordset("mark")
        FMpfbb.scrRingno.Text = data2.Recordset("ringno")
        FMpfbb.Show 1
    Next ii%
    table2.BookmarkCount = 0
End Sub

Sub ButtDisplayFlown_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Which races has this bird flown - select bird first"
End Sub

Sub ButtDisplayPdg_Click ()

    Dim ii As Integer

    If Table1.Visible = True Then
    ' Process each selected row.
        For ii% = 0 To Table1.BookmarkCount - 1
            'Workaround - only allow 2 selections
            If ii% > 0 Then
                Exit For
            End If
            data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
            FMpedg.Fyr.Text = data1.Recordset("yr")
            FMpedg.Fmark.Text = data1.Recordset("mark")
            FMpedg.Fring.Text = data1.Recordset("ringno")
            FMpedg.Show 1
        Next ii%
        Table1.BookmarkCount = 0
    ElseIf Table3.Visible = True Then
    ' Process each selected row.
        For ii% = 0 To Table3.BookmarkCount - 1
            'Workaround - only allow 2 selections
            If ii% > 0 Then
                Exit For
            End If
            Data3.Recordset.Bookmark = Table3.BookmarkList(ii%)
            FMpedg.Fyr.Text = Data3.Recordset("yr")
            FMpedg.Fmark.Text = Data3.Recordset("mark")
            FMpedg.Fring.Text = Data3.Recordset("ringno")
            FMpedg.Show 1
        Next ii%
        Table3.BookmarkCount = 0
    End If

    ' Process each selected row.
    For ii% = 0 To table2.BookmarkCount - 1
        'Workaround - only allow 2 selections
        If ii% > 0 Then
            Exit For
        End If
        data2.Recordset.Bookmark = table2.BookmarkList(ii%)
        FMpedg.Fyr.Text = data2.Recordset("yr")
        FMpedg.Fmark.Text = data2.Recordset("mark")
        FMpedg.Fring.Text = data2.Recordset("ringno")
        FMpedg.Show 1
    Next ii%
    table2.BookmarkCount = 0
End Sub

Sub ButtDisplayPdg_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Select bird then click for Pedigree"
End Sub

Sub ButtExclude_Click ()

    On Error GoTo FMErr2:
    Dim NL As String
    Dim ii As Integer

'Stop
    NL$ = Chr(13)
        For ii% = 0 To table2.BookmarkCount - 1
            Screen.MousePointer = 11
            data2.Recordset.Bookmark = table2.BookmarkList(ii%)
            data2.Recordset.Delete
            nCntEnt% = nCntEnt% - 1
            RaceEntries.Caption = "Race Entries " & Trim(Str(nCntEnt%))
        Next ii%
    table2.BookmarkCount = 0
    If Err = 3021 Then Resume
    Screen.MousePointer = 0
    Exit Sub
FMErr2:
    Select Case Err
    Case 3021  'No Current Record
        MsgBox "No Bird Records"
        Resume Next
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next

End Sub

Sub ButtExclude_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err3
MsgPanel.Caption = "Click to remove selected birds from race (on right)"
Exit Sub
Err3:
    Select Case Err
    Case 3021  'No Current Record
        Resume Next
    Case 3022
        MsgBox "Bird ALREADY selected for race" & Chr(13) & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & Chr(13) & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next
End Sub

Sub ButtInclude_Click ()
    On Error GoTo FMErr11:
    Dim ii As Integer
    Dim NL As String
    NL$ = Chr$(10)

    ' Process each selected row.
    DoEvents
    For ii% = 0 To Table1.BookmarkCount - 1
        data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
        data2.Recordset.AddNew
        data2.Recordset("RaceCode") = CLng(Trim(mRaceCode.Text))
        data2.Recordset("code_race") = Trim(mCodeRace.Text)
        If mRaceStatus.Text = "Race" Then
            data2.Recordset("RaceStatus") = "R"
        Else
            data2.Recordset("RaceStatus") = "T"
        End If
        data2.Recordset("reldate") = Format(mRaceDate.Text, "ddddd")
        data2.Recordset("code_pnt") = Trim(mCodePnt.Text)
        data2.Recordset("yr") = data1.Recordset("yr")
        data2.Recordset("mark") = data1.Recordset("mark")
        data2.Recordset("ringno") = data1.Recordset("ringno")
        data2.Recordset("coderating") = "HO" 'indicate Homed
        data2.Recordset("rating") = "HO" 'indicate Homed
        data2.Recordset.Update
        'Table1.BookmarkList(ii) = ""
        nCntEnt% = nCntEnt% + 1
        RaceEntries.Caption = "Race Entries " & Trim(Str(nCntEnt%))

    Next ii%

    Table1.BookmarkCount = 0
    data2.Refresh
    Exit Sub

FMErr11:
    Select Case Err
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Table1.BookmarkList(ii) = ""
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
        Resume Next
    End Select
    Resume Next
End Sub

Sub ButtInclude_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Click to Nominate selected birds for race (on right)"
End Sub

Sub ButtPrint_Click ()
FMPrtRaceSheet.mRaceCode.Text = mRaceCode.Text
FMPrtRaceSheet.mRaceDate.Text = mRaceDate.Text
MsgPanel.Caption = ""
FMPrtRaceSheet.Show 1
End Sub

Sub ButtPrint_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click to Print RACE entry Sheet"
End Sub

Sub CmdIncludeTeam_Click ()
    On Error GoTo FMErrX1:
    Dim ii As Integer
    Dim NL As String
    NL$ = Chr$(10)

    ' Process each selected row.
    DoEvents
    For ii% = 0 To Table3.BookmarkCount - 1
        Screen.MousePointer = 11
        Data3.Recordset.Bookmark = Table3.BookmarkList(ii%)
        data2.Recordset.AddNew
        data2.Recordset("RaceCode") = CLng(Trim(mRaceCode.Text))
        data2.Recordset("code_race") = Trim(mCodeRace.Text)
        If mRaceStatus.Text = "Race" Then
            data2.Recordset("RaceStatus") = "R"
        Else
            data2.Recordset("RaceStatus") = "T"
        End If
        data2.Recordset("reldate") = Format(mRaceDate.Text, "ddddd")
        data2.Recordset("code_pnt") = Trim(mCodePnt.Text)
        data2.Recordset("yr") = Data3.Recordset("yr")
        data2.Recordset("mark") = Data3.Recordset("mark")
        data2.Recordset("ringno") = Data3.Recordset("ringno")
        data2.Recordset("coderating") = "HO" 'indicate Homed
        data2.Recordset("rating") = "HO" 'indicate Homed
        data2.Recordset.Update
        'Table3.BookmarkList(ii) = ""
        nCntEnt% = nCntEnt% + 1
        RaceEntries.Caption = "Race Entries " & Trim(Str(nCntEnt%))

    Next ii%

    Table3.BookmarkCount = 0
    data2.Refresh
    table2.Refresh
    Screen.MousePointer = 0
    Exit Sub

FMErrX1:
    Select Case Err
    Case 3021
        Resume
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data3.Recordset("yr"), "") & "-" & Data3.Recordset("mark") & "-" & Data3.Recordset("ringno"), 16, "Cannot add bird"
        Table3.BookmarkList(ii) = ""
        Screen.MousePointer = 0
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
        Resume Next
    End Select
    Resume Next
End Sub

Sub CmdIncludeTeam_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click to move selected birds (left) to Race entry (right)"
End Sub

Sub CmdMaster_Click ()
'Turn stuff for Team selections
TeamBirds.Visible = False
Label1.Visible = False
CSCombo1.Visible = False
CSCombo1.Enabled = False
Table3.Visible = False
Table3.Enabled = False
CmdIncludeTeam.Visible = False
CmdIncludeTeam.Enabled = False
FrameSelect.Visible = False

'Turn off stuff for Master selection
RaceBirds.Visible = True
Table1.Visible = True
Table1.Enabled = True
ButtInclude.Visible = True
ButtInclude.Enabled = True
ButtExclude.Visible = True
ButtExclude.Enabled = True
End Sub

Sub CmdMaster_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click to Display ALL current birds on file"
End Sub

Sub CmdTeams_Click ()

'Turn stuff for Team selections
TeamBirds.Visible = True
Label1.Visible = True
CSCombo1.Visible = True
CSCombo1.Enabled = True
Table3.Visible = True
Table3.Enabled = True
CmdIncludeTeam.Visible = True
CmdIncludeTeam.Enabled = True
ButtExclude.Visible = True
ButtExclude.Enabled = True
FrameSelect.Visible = True

'Turn off stuff for Master selection
RaceBirds.Visible = False
Table1.Visible = False
Table1.Enabled = False
ButtInclude.Visible = False
ButtInclude.Enabled = False

'setfocus
CSCombo1.SetFocus

End Sub

Sub CmdTeams_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click to Display list of Team birds"
End Sub

Sub CSCombo1_Change ()
Dim strSelection As String
Dim NL As String

On Error GoTo FMLstTeamERR1:

strSelection$ = ""
NL$ = Chr(13)

Data3.Refresh
strSelection$ = "SELECT * FROM teambird WHERE code = '" & CSCombo1.Value & "'"
strSelection$ = strSelection$ & " ORDER BY val(ringno), ringno, mark, yr"
Data3.RecordSource = strSelection$
Data3.Refresh
DoEvents
Data3.Recordset.MoveLast
nCntteam& = Data3.Recordset.RecordCount
TeamBirds.Caption = Trim(Str(nCntteam&)) & " BIRDS IN TEAM"
Data3.Recordset.MoveFirst
DoEvents
'Table3.Refresh
Exit Sub

FMLstTeamERR1:
    Select Case Err
    Case 3021
        'MsgBox "No Current record" & NL$ & "so there!", 16, "ERROR"
        Resume Next
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next
End Sub

Sub Form_Activate ()
On Error GoTo FMerr1
    Dim StrCriteria1 As String
    Dim StrCriteria2 As String
    FreeLocks
    DoEvents
    StrCriteria1$ = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY val(ringno), ringno, yr, mark"
    data2.RecordSource = StrCriteria1$ ' & StrCriteria2$ 'Racetrans
    data2.Refresh
    FreeLocks
    DoEvents
    Table1.Refresh
    DoEvents
    FreeLocks
    table2.Refresh
    DoEvents
    'Find Point Description
    tblPoint.Seek "=", mCodePnt.Text
    If tblPoint.NoMatch Then
        MsgBox "Race Point details have been removed or corrupted", 16, "Race Point not on file"
        GLracepoint = "NOT ON FILE"
    Else
        GLracepoint$ = tblPoint("pntname")
        mCodePnt.Text = tblPoint("code_pnt")
        mPntName.Text = GLracepoint$
    End If
    data2.Recordset.MoveFirst
    ButtPrint.Visible = True
    'ButtAbort.Visible = False
    data1.Recordset.MoveLast
    RaceBirds.Caption = "Current Race Birds " & Trim(Str(data1.Recordset.RecordCount))
    data1.Recordset.MoveFirst
    FreeLocks
    data2.Recordset.MoveLast
    FreeLocks
    nCntEnt% = data2.Recordset.RecordCount
    RaceEntries.Caption = "Race Entries " & Trim(Str(nCntEnt%))
    data2.Recordset.MoveFirst
    FreeLocks
    CSCombo1.Text = Data4.Recordset("description")
    CSCombo1.Value = Data4.Recordset("code")
    DoEvents
    FreeLocks
    'CSCombo1_change
    Exit Sub

FMerr1:
    Select Case Err
    Case 3021 'No Current Record
        Resume Next
    'Case 3022
        'MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        'Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next

End Sub

Sub Form_Load ()
    Dim StrCriteria As String
    Dim StrCriteria1 As String
    data1.DatabaseName = DBFullFileName$
    'Select birds which are Race, Current Birds and not marked as Stock
    StrCriteria$ = "SELECT * FROM master WHERE indsoh = -1 and indrace = -1 and indstock = 0 ORDER BY val(ringno), ringno, yr, mark;" 'or Primary key
    data1.RecordSource = StrCriteria$
    data1.Refresh
    FreeLocks
    DoEvents

    'StrCriteria1$ = "SELECT * FROM racetrans WHERE code_race = '" & mCodeRace.Text & "' ORDER BY yr ASC, mark, ringno;" 'or Primary key
    data2.DatabaseName = DBFullFileName$
'    Data2.RecordSource = StrCriteria1$
'    Data2.Refresh
    Data3.DatabaseName = DBFullFileName$
    Data4.DatabaseName = DBFullFileName$
    Data4.RecordSource = "SELECT * FROM Teams order by description DESC"
    FreeLocks
    DoEvents
    
    'Open Database
    Set DB = OpenDatabase(DBFullFileName$, True, False)  'Open Database
    Set tblPoint = DB.OpenTable("Points", 4)    'Open Table
    tblPoint.Index = "Points"              'Open Index
    
    'Make Select List Invisible
'Turn stuff for Team selections
TeamBirds.Visible = False
Label1.Visible = False
CSCombo1.Visible = False
CSCombo1.Enabled = False
Table3.Visible = False
Table3.Enabled = False
CmdIncludeTeam.Visible = False
CmdIncludeTeam.Enabled = False

'Turn off stuff for Master selection
RaceBirds.Visible = False
Table1.Visible = False
Table1.Enabled = False
ButtInclude.Visible = False
ButtInclude.Enabled = False
ButtExclude.Visible = False
ButtExclude.Enabled = False
FrameSelect.Visible = False

End Sub

Sub Form_Unload (Cancel As Integer)
On Error GoTo FMerr9

DoEvents
FreeLocks
'tblPoint.Close
Set tblPoint = Nothing
DB.Close
Set DB = Nothing
Exit Sub

FMerr9:
    Select Case Err
    Case 3021 'No Current Record
        'MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        Resume Next
    'Case 3022
        'MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        'Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & Error$(Err), 16, "Error in FMrace.frm"
        Resume Next
    End Select
    Resume Next
End Sub

Sub indAll_Click (Value As Integer)
setselections
End Sub

Sub indNone_Click (Value As Integer)
setselections
End Sub

Sub indNotLost_Click (Value As Integer)
setselections
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = ""
End Sub

Sub setselections ()
On Error GoTo SetSelErr:
If indNone Then
    Table3.BookmarkCount = 0
    Data3.Recordset.MoveFirst
    Exit Sub
ElseIf indAll Then
    Data3.Recordset.MoveFirst
    While Not Data3.Recordset.EOF
        Table3.BookmarkList(Table3.BookmarkCount) = Data3.Recordset.Bookmark
        Data3.Recordset.MoveNext
    Wend
    Data3.Recordset.MoveFirst
ElseIf indNotLost Then
    Table3.BookmarkCount = 0
    Data3.Recordset.MoveFirst
    While Not Data3.Recordset.EOF
        If Data3.Recordset("status") <> "L" Then
            Table3.BookmarkList(Table3.BookmarkCount) = Data3.Recordset.Bookmark
        End If
        Data3.Recordset.MoveNext
    Wend
    Data3.Recordset.MoveFirst
End If

Exit Sub

SetSelErr:
    Select Case Err
    Case 3021
        Resume
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & Chr$(13) & Error$(Err), 16, "Error in FMrace.frm"
        Resume Next
    End Select
    Resume Next
End Sub

Sub Table1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "To nominate any bird, select (highlight) then click |>>>| button"
End Sub

Sub Table2_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "The birds here are nominated for the race && will be on Race entry sheet"
End Sub

Sub Table3_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "To nominate team bird, select (highlight) then click |>>>| button"
End Sub

