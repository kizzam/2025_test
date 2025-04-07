VERSION 2.00
Begin Form FMdsprpt 
   Caption         =   "Report - Display"
   ClientHeight    =   6600
   ClientLeft      =   420
   ClientTop       =   1380
   ClientWidth     =   9120
   Height          =   7110
   Left            =   360
   LinkTopic       =   "Form5"
   ScaleHeight     =   6600
   ScaleWidth      =   9120
   Top             =   930
   Width           =   9240
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6240
      Width           =   9135
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9135
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   5520
         Width           =   2175
      End
      Begin SSCommand cmdCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5040
         TabIndex        =   0
         Top             =   5520
         Width           =   2175
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Report Format"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   7560
         ShadowStyle     =   1  'Raised
         TabIndex        =   19
         Top             =   720
         Width           =   1455
         Begin SSOption ShortRing 
            Caption         =   "Short Ring"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin SSOption FullRing 
            Caption         =   "Full Ring"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   480
            Width           =   1095
         End
      End
      Begin SSCommand cmdAbort 
         Caption         =   "&Abort"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   18
         Top             =   5520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Show Only"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   3720
         ShadowStyle     =   1  'Raised
         TabIndex        =   17
         Top             =   720
         Width           =   3735
         Begin SSCheck BredByMe 
            Caption         =   "Bred by Me"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
         Begin SSOption ShowCurrent 
            Caption         =   "Current"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin SSOption ShowAll 
            Caption         =   "All"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Sort by"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1920
         ShadowStyle     =   1  'Raised
         TabIndex        =   16
         Top             =   720
         Width           =   1695
         Begin SSOption byDam 
            Caption         =   "Dam"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin SSOption bySire 
            Caption         =   "Sire"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H0080FFFF&
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         ForeColor       =   &H00000000&
         HeadingHeight   =   1
         Height          =   3495
         HorzLines       =   0  'None
         Layout          =   FMDSPRP.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   9
         Top             =   1560
         UseBookmarks    =   -1  'True
         Width           =   8655
         WrapCellPointer =   0   'False
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMDSPRP.FRX:0B7A
         Contents        =   FMDSPRP.FRX:0BBA
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   360
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   1
         Text            =   "CSCombo1"
         TextDisplay     =   0  'first visible
         Top             =   960
         ValueCol        =   -1
         Width           =   1455
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7560
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Note: The above list includes every bird record on file (every record for selected year)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5160
         Width           =   8655
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Year:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Breeding by Matings"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   8535
      End
   End
End
Option Explicit
Dim FgCntBirds As Long
Dim flag As Integer

Sub BredByMe_Click (Value As Integer)
setselections
End Sub

Sub ButtCancel_Click ()
    Unload Me
End Sub

Sub ByDam_Click (Value As Integer)
setselections

End Sub

Sub BySire_Click (Value As Integer)
setselections
End Sub

Sub cmdAbort_Click ()
    flag% = -1
End Sub

Sub CmdCancel_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
    Dim Offset As Integer, Trigger As Integer
    Dim Next_Line As Integer
    Dim Line_Count As Integer, item_count As Integer
    Dim CounterR As Integer, CounterC As Integer
    Dim Line_Desc As String
    Dim Detail As String
    Dim NoMatings As Integer
    
    Dim reply As Integer
    Dim flag As Integer
    Dim Msgstmt As String
    Dim MlMargin As Single
    Dim MrMargin As Single
    Dim RecCnt As Integer
    Dim Sire As String, Sire1 As String, Dam As String, dam1 As String
    Dim Prodgeny As String
    Dim CapMsg As String
    
    On Error GoTo DB_err2:

    NoMatings% = 0

    'Print the Report
    Screen.MousePointer = 11
    Msgpanel.Caption = "Please wait - Report being generated & printed"

    CmdAbort.Visible = True
    CmdPrint.Visible = False
    CmdCancel.Visible = False

    Msgstmt$ = "Ensure PRINTER is turned ON and On-Line" & Chr(10) & "Race Entry Sheet is about to be printed"
    reply% = MsgBox(Msgstmt$, 33, "Click OK when ready")
    If reply% <> 1 Then
        flag% = -1
        CmdAbort.Visible = False
        CmdPrint.Visible = True
        CmdCancel.Visible = True
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'read on system init from LOFT table, set Page length
    If FullRing Then
        Line_Count% = GLRSLpp% - 6
    Else
        Line_Count% = GLRSLpp%
    End If
    
    Data1.Recordset.MoveLast
    RecCnt% = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst

    'Can't print race sheet with no birds
    If RecCnt% < 1 Then
        MsgBox "No Birds in selection list", 16, "NOTHING TO PRINT"
        Screen.MousePointer = 0
        CmdAbort.Visible = False
        CmdPrint.Visible = True
        CmdCancel.Visible = True
        Exit Sub
    End If
    
    Data1.Recordset.MoveFirst

    ' Offset for left hand edge of duplicate.
    Offset% = 108
    MlMargin! = 5
    MrMargin! = 200

    Printer.ScaleMode = 6  ' millimeters
    Next_Line% = 6

    'Print Heading
    GoSub Heading

    Sire1$ = ""
    dam1$ = ""

    Printer.FontName = "Times New Roman"
    Printer.FontSize = 12
    Printer.FontBold = True
    item_count% = 0

    Do While Not Data1.Recordset.EOF
       DoEvents
       If flag% = -1 Then
            MsgBox "Print Aborted by User", 16, "Aborted"
            Printer.CurrentX = MlMargin!
            Printer.Print " "
            Printer.FontSize = 14
            Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
            GoSub footer
            Printer.EndDoc
            'ButtPrint.Visible = True
            CmdAbort.Visible = False
            CmdPrint.Visible = True
            CmdCancel.Visible = True
            Screen.MousePointer = 0
            Msgpanel.Caption = ""
            Exit Sub
       End If
       
       Printer.FontName = "Times New Roman"
       Printer.FontSize = 14
       Printer.FontBold = True
       Sire$ = Trim(Trim(Data1.Recordset("sireyr") & "/") & Trim(Data1.Recordset("siremark") & "/") & Trim(Data1.Recordset("sireringno") & ""))
       Dam$ = Trim(Trim(Data1.Recordset("damyr") & "/") & Trim(Data1.Recordset("dammark") & "/") & Trim(Data1.Recordset("damringno") & ""))
       
       If FullRing Then
        Prodgeny$ = Trim(Trim(Data1.Recordset("yr") & "/") & Trim(Data1.Recordset("mark") & "/") & Trim(Data1.Recordset("ringno") & ""))
       Else
        Prodgeny$ = Trim(Data1.Recordset("ringno") & "")
       End If

       If Sire$ <> Sire1$ Or Dam$ <> dam1$ Then 'different mating
        NoMatings% = NoMatings% + 1
        Sire1$ = Sire$
        dam1$ = Dam$
        GoSub HeadingMating
       End If
       
       If FullRing Then
        If Printer.CurrentX > 159 Then
            Printer.CurrentY = Printer.CurrentY + Next_Line%
            Printer.CurrentX = MlMargin! + 40
        End If
       Else
        If Printer.CurrentX > 182 Then
            Printer.CurrentY = Printer.CurrentY + Next_Line%
            Printer.CurrentX = MlMargin! + 40
        End If
       End If
       
       If Printer.CurrentY > 250 Then
       'print footer then new header
        GoSub footer
        GoSub Heading
        GoSub HeadingMating
       End If
       
       Printer.FontSize = 14
       Printer.Print Prodgeny$ & "   ";
    
       item_count% = item_count% + 1
       Data1.Recordset.MoveNext

    Loop
    
    ' Print the number of items.
    'Printer.CurrentY = Printer.CurrentY + Next_Line%
    Printer.CurrentY = Printer.CurrentY + Next_Line%
    DrawWidth = 50
    Printer.Line (MlMargin!, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = Printer.CurrentY + Next_Line%
    Printer.Print "Number of Matings = ";
    Printer.Print NoMatings% + 1;
    
    Printer.Print "     Number of Birds bred = ";
    Printer.Print item_count%
    Printer.CurrentX = MlMargin! + 50

    GoSub footer
    Printer.EndDoc

    ' Reset mouse pointer.
    Screen.MousePointer = 0
    CmdAbort.Visible = False
    CmdPrint.Visible = True
    CmdCancel.Visible = True
    Exit Sub

Heading:
    ' Print Heading
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = 8
    Printer.Print "Prodgeny by Matings of '" & CSCombo1 & "' for " & GLOwner$
    'Print Horizontal Line
    DrawWidth = 30
    Printer.CurrentY = Printer.CurrentY
    'Printer.Line (MlMargin!, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
Return

footer:
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = 269
    Printer.Print "FLYERS NAME: ";
    Printer.FontBold = True
    Printer.Print GLOwner$; " " & GLPhoneNo
    Printer.FontBold = False
    Printer.FontSize = 6
    Printer.CurrentX = MlMargin!
    Printer.Print "Registered to " & GLOwner$
    Printer.CurrentX = MlMargin!
    Printer.Print "Produced by Racing Pigeon Management Software V3.2 (Copyright K Malone 1988-2001, All Rights reserved)"
    Printer.NewPage
    DoEvents
Return

HeadingMating:
    If Printer.CurrentY > 260 Then
       'print footer then new header
        GoSub footer
        GoSub Heading
    End If
        Printer.CurrentY = Printer.CurrentY + Next_Line% + 2
        Printer.Line (MlMargin!, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
        Printer.CurrentX = MlMargin!
        Printer.FontSize = 12
        Printer.Print "MATING - Sire: ";
        Printer.FontUnderline = True
        Printer.FontSize = 14
        Printer.Print Sire1$;
        Printer.FontUnderline = False
        Printer.FontSize = 12
        Printer.Print " X Dam: ";
        Printer.FontUnderline = True
        Printer.FontSize = 14
        Printer.Print dam1$;
        Printer.FontUnderline = False
        Printer.CurrentX = MlMargin! + 20
        Printer.CurrentY = Printer.CurrentY + Next_Line%
        Printer.FontSize = 12
        Printer.Print "Prodgeny:";
        Printer.CurrentX = MlMargin! + 40
Return


DB_err2:
    Select Case Err
    Case 482
        MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
        Exit Sub
    Case 3021
        MsgBox "Error " & Err & " " & Val(Error) & Chr(10) & "(Database Empty)", 16, "INFORMATION ONLY"
        Resume Next
    Case 3022
        MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "Error " & Err & ", at line " & Erl, 16, "Error Code " & Val(Error) & " " & Error$
        Resume Next
    End Select
    'what about closing files?
    MsgBox Error$, 16, "Program Error"
    Printer.EndDoc
    Screen.MousePointer = 0
    CmdAbort.Visible = False
    CmdPrint.Visible = True
    CmdCancel.Visible = True
    Msgpanel.Caption = ""
End Sub

Sub CSCombo1_Change ()
setselections
End Sub

Sub Form_Load ()
On Error GoTo FMErr1:
    Data1.DatabaseName = DBFullFileName$
    DoEvents
    CSCombo1.Text = Trim(Str(Year(Now) - 1))
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
    If BySire Then
        OrderBy$ = "sireyr, siremark, sireringno,damyr, dammark, damringno, yr,ringno"
    ElseIf ByDam Then
        OrderBy$ = "damyr, dammark, damringno,sireyr, siremark, sireringno, yr,ringno"
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
        Data1.RecordSource = "SELECT * FROM master ORDER BY " & o$
    Else
        Data1.RecordSource = "SELECT * FROM master WHERE " & s$ & " ORDER BY " & o$
    End If
    Data1.Refresh
    Data1.Recordset.MoveLast
    FgCntBirds& = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    Msgpanel.Caption = "Count of birds listed = " & Str(FgCntBirds&)
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

Sub ShowAll_Click (Value As Integer)
setselections

End Sub

Sub showCurrent_Click (Value As Integer)
setselections

End Sub

Function ShowWhat () As String

    Dim s As String

    If ShowAll Then
        s$ = ""
    Else
       s$ = "indsoh = True"
    'Else
    '   s$ = "indsoh = false"
    End If

    If BredByMe Then
        If s$ = "" Then
            s$ = "bredby = '" & GLOwner & "'"
        Else
            s$ = s$ & " AND bredby = '" & GLOwner & "'"
        End If
    End If
    
    If CSCombo1.Text <> "ALL" Then
        If CSCombo1.Value = "Pre 80" Then
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
    
'    If CSCombo1 = "ALL" Then
'        data1.RecordSource = "SELECT * FROM master ORDER BY sireyr, siremark, sireringno, damyr, dammark, damringno, YR, ringno"
'    ElseIf CSCombo1 = "Pre 1980" Then
'        data1.RecordSource = "SELECT * FROM master WHERE yr < 1980 ORDER BY sireyr, siremark, sireringno, damyr, dammark, damringno, ringno"
'    Else
'        data1.RecordSource = "SELECT * FROM master WHERE yr = " & CSCombo1 & " ORDER BY sireyr, siremark, sireringno, damyr, dammark, damringno, ringno"
'    End If
End Function

