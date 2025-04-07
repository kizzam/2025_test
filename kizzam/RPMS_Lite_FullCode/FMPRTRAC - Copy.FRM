VERSION 2.00
Begin Form FMPrtRaceSheet 
   Caption         =   "Printing Race Sheet"
   ClientHeight    =   5505
   ClientLeft      =   930
   ClientTop       =   1155
   ClientWidth     =   6960
   Height          =   5910
   Left            =   870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6960
   Top             =   810
   Width           =   7080
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   6975
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   735
         Left            =   3840
         TabIndex        =   19
         Top             =   4080
         Width           =   1935
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Sort Order of Birds on Race Entry Sheet"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   1560
         ShadowStyle     =   1  'Raised
         TabIndex        =   15
         Top             =   2520
         Width           =   3975
         Begin SSOption OptMarkRingNo 
            Caption         =   "by Mark then Ring No."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2175
         End
         Begin SSOption OptYrRingNo 
            Caption         =   "By Year, then Ring No."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   840
            Width           =   2295
         End
         Begin SSOption OptSexRingNo 
            Caption         =   "by Sex, then Ring No."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   600
            Width           =   2175
         End
         Begin SSOption OptRingNo 
            Caption         =   "by Ring Nos (Normal)"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin SSCommand cmdGo 
         Caption         =   "&Continue"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1320
         TabIndex        =   14
         Top             =   4080
         Width           =   1935
      End
      Begin TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   2040
         Width           =   735
      End
      Begin TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2040
         Width           =   735
      End
      Begin TextBox mRaceDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin TextBox mRacePoint 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin TextBox mRaceCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin SSPanel MsgPanel1 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Please wait race sheet being prepared and printed"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
      Begin SSCommand CmdAbort 
         Caption         =   "&Abort Print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   2520
         TabIndex        =   0
         Top             =   4080
         Width           =   1935
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lines per page:"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Sheet Style:"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   1935
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Point:"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Date:"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Code:"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
End
Option Explicit
Dim Flag As Integer
Dim DB As Database
Dim tblMaster As Table
Dim tblColour As Table
Dim DSRTrans As Dynaset
Dim SQLStmt As String
Dim Msgstmt As String
Dim Line_Count As Integer
Dim RecCnt As Integer
Dim X As Integer
Dim Y As Integer
Dim TpCM As Integer 'Twips per CM
Dim NoPages As Integer
Dim CntBirds As Integer
Dim CapMsg As String
'Dim mRaceC As Long

Dim MlMargin As Single
Dim MtMargin As Single
Dim MrMargin As Single
Dim MlOffSet As Single
Dim MtOffSet As Single
Dim Mline As Single
Dim LastHoriz As Single

'Box Corners
Dim TLX As Single, TLY As Single

Dim Col_1 As Single
Dim Col_2 As Single
Dim Col_3 As Single
Dim Col_4 As Single
Dim Col_5 As Single
Dim Col_6 As Single
Dim Col_7 As Single
Dim Col_8 As Single
Dim Col_9 As Single
Dim Col_10 As Single
Dim Col_11 As Single
Dim Col_12 As Single
Dim Col_13 As Single
Dim Col_14 As Single
Dim Col_15 As Single
Dim Col_16 As Single

Dim Row_1 As Single
Dim Row_2 As Single
Dim Row_3 As Single
Dim Row_4 As Single
Dim Row_5 As Single
Dim Row_6 As Single
Dim RowFt1 As Single
Dim RowFt2 As Single
Dim RowFt3 As Single
Dim RowFt4 As Single
Dim reply As Integer
Dim RSLPP As Integer

Sub cmdAbort_Click ()
    Flag% = -1
    MsgPanel.Caption = "Print ABORTED by user"
    MsgPanel1.Caption = "Printing aborted by user"
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdGo_Click ()

On Error GoTo RS_Err0

Dim PrtOk As Integer
Dim MyQuery As QueryDef
Dim SortOrder As String

CmdGo.Visible = False
CmdExit.Visible = False
CapMsg$ = "Please WAIT Race Sheet Printing - Page: "

'Print the Race Sheet
Screen.MousePointer = 11
MsgPanel.Caption = "Please WAIT - Race Sheet being prepared"
TpCM% = 567 'No of Twips per CM

'Set Page length according to value in Loft setup (28 or 30)
Line_Count% = GLRSLpp%    'read on system init from LOFT table

'Open DB & Files
Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
Set tblMaster = DB.OpenTable("Master", 4)    'Open Table
tblMaster.Index = "bird_ymr"              'Open Index
Set tblColour = DB.OpenTable("Colour", 4)    'Open Table
tblColour.Index = "colour"              'Open Index

' Create the new Dynaset.
' mRaceC& = CLng(Val(Trim(mRaceCode.Text)))
If OptRingNo Then
    SQLStmt$ = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY val(ringno), ringno, mark, yr"
    Set DSRTrans = DB.CreateDynaset(SQLStmt$, 4)
    SortOrder$ = "by Ring No."
ElseIf OptSexRingNo Then
    ' Create a QueryDef, from trans and master to include Sex of bird.
    SQLStmt$ = "SELECT DISTINCTROW racetrans.yr, master.codesex, racetrans.RaceCode, racetrans.mark, racetrans.ringno FROM racetrans LEFT JOIN master ON (racetrans.ringno = master.ringno) AND (racetrans.mark = master.mark) AND (racetrans.yr = master.yr) WHERE (racetrans.RaceCode=" & Val(mRaceCode.Text) & ") ORDER BY master.codesex, val(racetrans.ringno), racetrans.ringno, racetrans.mark;"
    Set MyQuery = DB.CreateQueryDef("RTS", SQLStmt$)
    Set DSRTrans = DB.CreateDynaset("RTS")  ' Create dynaset from query.
    DoEvents
    FreeLocks
    MyQuery.Close           ' Close QueryDef object.
    DB.DeleteQueryDef ("RTS")' Delete QueryDef.
    SortOrder$ = "by Sex then Ring No."
ElseIf OptYrRingNo Then
    SQLStmt$ = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY yr, val(ringno), ringno, mark"
    Set DSRTrans = DB.CreateDynaset(SQLStmt$, 4)
    SortOrder$ = "by Year then Ring No."
ElseIf OptMarkRingNo Then
    SQLStmt$ = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY mark, val(ringno), ringno, yr"
    Set DSRTrans = DB.CreateDynaset(SQLStmt$, 4)
    SortOrder$ = "by Mark then Ring No."
End If
DSRTrans.MoveLast
RecCnt% = DSRTrans.RecordCount
'Can't print race sheet with no birds
If RecCnt% < 1 Then
    Screen.MousePointer = 0
    MsgBox "No Birds selected for race", 16, "CANNOT PRINT RACE SHEET"
    GoTo RSShutDown
End If
DSRTrans.MoveFirst

Msgstmt$ = "Ensure PRINTER is turned ON and On-Line" & Chr(10) & "Race Entry Sheet is about to be printed" & Chr(10) & "STYLE = " & GLRSRaceFormType% & Chr(10) & "Lines per page = " & GLRSLpp% & Chr(10) & "Sorted by " & SortOrder$
Screen.MousePointer = 0
reply% = MsgBox(Msgstmt$, 33, "Click OK when ready")
If reply% <> 1 Then
    Flag% = -1
    GoTo RSShutDown
    Exit Sub
End If
Screen.MousePointer = 11
cmdAbort.Visible = True

'Calc no of pages
CntBirds% = 0
NoPages% = Int((RecCnt% / Line_Count%) + .99)

If GLRSRaceFormType% = 1 Or GLRSRaceFormType% = 2 Then
    PrtOk% = RaceSheet1("0")
ElseIf GLRSRaceFormType% = 3 Or GLRSRaceFormType% = 4 Then
    PrtOk% = RaceSheet2()
ElseIf GLRSRaceFormType% = 5 Or GLRSRaceFormType% = 6 Then 'QPF New Race sheet required
    PrtOk% = RaceSheet3()
ElseIf GLRSRaceFormType% = 7 Or GLRSRaceFormType% = 8 Then 'exact new QPF race sheet required
    PrtOk% = RaceSheet4()
ElseIf GLRSRaceFormType% = 9 Or GLRSRaceFormType% = 10 Then 'TASMANIA race sheet required
    PrtOk% = RaceSheet5()
ElseIf GLRSRaceFormType% = 11 Then
    PrtOk% = RaceSheet11()
ElseIf GLRSRaceFormType% = 12 Then 'VHA race sheet required
    PrtOk% = RaceSheet12("12")
ElseIf GLRSRaceFormType% = 13 Then 'Lismore race sheet required
    PrtOk% = RaceSheet1("1")
ElseIf GLRSRaceFormType% = 14 Then 'Lismore race sheet required
    PrtOk% = RaceSheet1("1")
ElseIf GLRSRaceFormType% = 15 Then 'Central Highlands race sheet required (Similar to 12)
    PrtOk% = RaceSheet12("15")
End If

cmdAbort.Visible = False
If PrtOk Then
    MsgPanel.Caption = CapMsg$ & Str$(NoPages%) & " of " & Str$(NoPages%) & " - COMPLETE"
Else
    MsgPanel.Caption = CapMsg$ & Str$(NoPages%) & " of " & Str$(NoPages%) & " - COMPLETE"
End If

GoTo RSShutDown

RS_Err0:
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Sub
    Case 3012
	DB.DeleteQueryDef ("RTS")' Delete QueryDef.
	Resume
    Case 3021
	MsgBox "Error " & Err & " " & Val(Error) & Chr(10) & "(Database Empty)", 16, "INFORMATION ONLY"
	Resume Next
    Case 3022
	MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
	MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case 3061
	MsgBox "Error " & Err & ", at line " & Erl, 16, "Error Code " & Val(Error) & " " & Error$
	Resume Next
    Case 3075
	MsgBox "Error " & Err & ", at line " & Erl, 16, "Error Code " & Val(Error) & " " & Error$
    Case Else
	MsgBox "Error " & Err & ", at line " & Erl, 16, "Error Code " & Val(Error) & " " & Error$
    End Select

RSShutDown:
    cmdAbort.Visible = False
    CmdGo.Visible = True
    CmdExit.Visible = True
    MsgPanel.Caption = ""
    tblMaster.Close
    tblColour.Close
    DSRTrans.Close
    DB.Close
    Set DSRTrans = Nothing
    Set tblColour = Nothing
    Set tblMaster = Nothing
    Set DB = Nothing
    Screen.MousePointer = 0
    Unload Me
End Sub

Sub Form_Activate ()
Screen.MousePointer = 11
Screen.MousePointer = 0
End Sub

Sub Form_Load ()
Screen.MousePointer = 11
cmdAbort.Visible = False
FMPrtRaceSheet.mRacePoint.Text = GLracepoint$
FMPrtRaceSheet.MsgPanel.Caption = "Check to ensure all is correct then click 'Continue' to print or 'EXIT' to Quit"
FMPrtRaceSheet.MsgPanel1.Caption = "Ensure Printer is ON LINE and Ready"
Text1.Text = Trim(Str(GLRSRaceFormType%))
'Select Case GLRSRaceFormType%
'    Case 1 Or 3 Or 5 Or 7 Or 9 Or 13
'        GLRSLpp% = 30
'    Case 2 Or 4 Or 6 Or 8 Or 10 Or 14
'        GLRSLpp% = 28
'    Case 11 Or 12
'        GLRSLpp% = 25
'End Select
Text2.Text = Str(GLRSLpp%)
Screen.MousePointer = 0
End Sub

Function RaceSheet1 (ByVal Type13 As String)
' Style 1 - TDPF Stye race sheet (No Breeder, Allows for two rubbers).
'GLRSRaceFormType% = 1

On Error GoTo RS1_err

RaceSheet1 = False

'Set Start of each column from left of page
Col_1! = 0 * TpCM%
Col_2! = .7 * TpCM%
Col_3! = 3.7 * TpCM%
Col_4! = 4.3 * TpCM%
Col_5! = 6 * TpCM%
Col_6! = 8.45 * TpCM%
Col_7! = 9.3 * TpCM%
Col_8! = 10.95 * TpCM%
Col_9! = 12.5 * TpCM%
Col_10! = 13.2 * TpCM%
Col_11! = 14.1 * TpCM%
Col_12! = 14.9 * TpCM%
Col_13! = 15.7 * TpCM%
Col_14! = 16.5 * TpCM%
Col_15! = 17.5 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = .9 * TpCM%
Row_3! = 1.5 * TpCM%
Row_4! = 2 * TpCM%
Row_5! = 2.5 * TpCM%
Row_6! = 3.7 * TpCM%

If GLRSLpp% > 28 Then
    RowFt1! = 24.3 * TpCM%
    RowFt2! = 25 * TpCM%
    RowFt3! = 25.7 * TpCM%
    RowFt4! = 26.5 * TpCM%
    LastHoriz! = 23.5
    MtMargin! = .63 * TpCM%
Else
    RowFt1! = 22.9 * TpCM%
    RowFt2! = 23.6 * TpCM%
    RowFt3! = 24.3 * TpCM%
    RowFt4! = 25.1 * TpCM%
    LastHoriz! = 22.1
    MtMargin! = .9 * TpCM%
End If

MlMargin! = 1 * TpCM%
MrMargin! = MlMargin! + (18.5 * TpCM%)
MlOffSet! = .09 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'do new page - format etc, DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.Print GLFed$
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_2! 'was 890
    Printer.Print "Please enter the following birds for - Race Point: ";
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.Print " Date: ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.Print " Page: ";
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.Print Str(Y%) & " of " & Str(NoPages%)

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 40
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + (LastHoriz! * TpCM%)))-(MrMargin, (MtMargin! + (LastHoriz! * TpCM%)))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_5!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentY = 1190
    Printer.FontBold = False
    Printer.Print "No";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "Colour";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Sex";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Ring No";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Mark";

    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Year";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Rubber Nos";

    'Col8

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    If Type13 = "1" Then
	Printer.Print "T/B";
    Else
	Printer.Print "50c";
    End If

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "COL";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    If Type13 = "1" Then
	Printer.Print "L/R";
    Else
	Printer.Print "$1";
    End If

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$2";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "$5";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Lofts"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 30
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    DrawStyle = 0
    DrawWidth = 30
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + RowFt1!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 12
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	    
	    Else
	    
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col2
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If

	     'Col3
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If

	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col5
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));

	     'Horizontal line between birds are printed previously
	     'Data2.recordset.MoveNext
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Draw Static Loft Auction/Pools etc.
	    
	    'Please Mark like this Box
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (4.5 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    'Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY!)'Right Vert
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    ' Cross X lines corner to corner
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))
	    'Printer.Line (TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))-(TLX!, TLY!)
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY!)
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (3.75 * TpCM%)
	    Printer.Print "PLEASE"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "MARK"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "THUS"
	
	    '50c Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (7.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (6.6 * TpCM%)
	    Printer.Print "50c"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$1 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (9.8 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (9.2 * TpCM%)
	    Printer.Print "$1.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$2 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (12.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (12.3 * TpCM%)
	    Printer.Print "$2.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"

	    '$5 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (15.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (15.3 * TpCM%)
	    Printer.Print "$5.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	
	    'Print line under Loft Auction
	    DrawWidth = 30
	    'Printer.Line (Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    Printer.Line (MlMargin! + Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    
    'Print rest
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.CurrentY = MtMargin! + (18.3 * TpCM%)
    Printer.Print "Any"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Other"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Pools"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "or"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Jackpots"
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    Printer.FontItalic = False
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    Printer.Print "Total No of Birds entered : ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print Str$(RecCnt%)
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.Print "Cost : ........................";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.FontBold = True
    Printer.Print "TOTAL"
	    
    'CLUB Box
    TLX! = Col_1! + MlMargin!
    TLY! = RowFt2! + MtMargin! - (.5 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + (.9 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt2! - MtOffSet!
    Printer.Print "CLUB"
    'Rest of Line
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt2! - MtOffSet!
    Printer.Print "NAME: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.Print GLOwner

    'FED Box
    TLX! = Col_1! + MlMargin!
    TLY! = RowFt3! + MtMargin! - (.5 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + (.9 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "FEDERATION"
    'Rest of line
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "SIGNATURE: ";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet! + 120
    Printer.Print "x....................................................";
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt4! - MtOffSet!
    Printer.Print "CLUB: ";
    Printer.FontBold = True
    If Len(Trim(GLOwnerClub)) > 26 Then
	Printer.FontSize = 9
    ElseIf Len(Trim(GLOwnerClub)) > 14 Then
	Printer.FontSize = 10
    Else
	Printer.FontSize = 11
    End If
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.Print GLOwnerClub
    Printer.FontBold = False

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + ((LastHoriz! + 2.8) * TpCM%)
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2z Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage

Next Y%
Printer.EndDoc
DoEvents
RaceSheet1 = True
Exit Function

RS1_err:
    RaceSheet1 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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
End Function

Function RaceSheet11 ()
' Style 11 - Similar to VHA Style race sheet (No Breeder, Allows for two rubbers).
'GLRSRaceFormType% = 11

On Error GoTo RS11_err

RaceSheet11 = False

'Set Start of each column from left of page
Col_1! = 0 * TpCM%
Col_2! = .7 * TpCM%
Col_3! = 1.4 * TpCM%
Col_4! = 2.25 * TpCM%
Col_5! = 4.7 * TpCM%
Col_6! = 6.4 * TpCM%
Col_7! = 9.45 * TpCM%
Col_8! = 10.05 * TpCM%
Col_9! = 12.5 * TpCM%
Col_10! = 13.3 * TpCM%
Col_11! = 14.1 * TpCM%
Col_12! = 14.9 * TpCM%
Col_13! = 15.7 * TpCM%
Col_14! = 16.5 * TpCM%
Col_15! = 17.6 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = 3 * TpCM%
Row_3! = 3.9 * TpCM%
Row_4! = 4.4 * TpCM%
Row_5! = 4.9 * TpCM%
Row_6! = 6.1 * TpCM%

RowFt1! = 23 * TpCM%
RowFt2! = 24.5 * TpCM%
RowFt3! = 27 * TpCM%
LastHoriz! = 22.4 * TpCM%

MlMargin! = 1 * TpCM%
MtMargin! = .75 * TpCM%
MrMargin! = MlMargin! + (18.5 * TpCM%)
MlOffSet! = .09 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'do new page - format etc, DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    'NEW STUFF
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.FontBold = True
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.FontSize = 8
    Printer.Print "This Entry is subject to the Rules & Regulations of"
    Printer.FontSize = 20
    Printer.CurrentX = MlMargin!
    Printer.Print GLFed$
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.FontSize = 9
    Printer.Print "ENTRY FORM & POOLS"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.Print "RACE POINT: ";
    Printer.FontBold = True
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.Print "DATE: ";
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.Print " Page: ";
    Printer.FontBold = True
    Printer.Print Str(Y%) & " of " & Str(NoPages%)
    Printer.CurrentX = MlMargin!
    Printer.FontBold = False
    Printer.Print "FLYERS NAME: ";
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.FontBold = True
    Printer.Print GLOwner$;
    'Printer.Print "Club: " & GLOwnerClub
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.FontBold = False
    Printer.Print "CLUB: ";
    Printer.FontBold = True
    Printer.Print GLMark
    Printer.CurrentX = MlMargin!
    Printer.FontBold = False
    Printer.Print "ADDRESS: ";
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.FontBold = True
    If GLAddress1 <> "" Then Printer.Print GLAddress1
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    If GLAddress2 <> "" Then Printer.Print GLAddress2
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print GLCity & " " & GLState & " " & GLPostcode
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print "Phone: " & GLPhoneNo

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS(Optional, mark with 'X' if desired)";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 140
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + LastHoriz!))-(MrMargin, (MtMargin! + LastHoriz!))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    DrawWidth = 8
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentY = 1190
    Printer.FontBold = False
    Printer.Print "No";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Year";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Ring Marks";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Ring No.";

    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Colour";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Sex";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.Print "Race Ring No";

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "Nom";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "50c";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "$1";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$2";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Lofts"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 30
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    DrawStyle = 0
    DrawWidth = 30
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + RowFt1!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 12
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col3
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	     
	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	    Else
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;

	     'Col3*
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));
	     
	     'Col4*
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print tblMaster("mark");
	     
	     'Col5*
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col6*
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If
	     
	     'Col7*
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If

	     'Horizontal line between birds are printed previously
	     'Data2.recordset.MoveNext
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Draw Static Loft Auction/Pools etc.
	    
	    'Please Mark like this Box
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (7 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    'Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY!)'Right Vert
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    ' Cross X lines corner to corner
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))
	    'Printer.Line (TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))-(TLX!, TLY!)
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY!)
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (6.25 * TpCM%)
	    Printer.Print "PLEASE"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "MARK"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "THUS"
	
	    '50c Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (9.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (9.1 * TpCM%)
	    Printer.Print "50c"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$1 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (12.3 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (11.7 * TpCM%)
	    Printer.Print "$1.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$2 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (15.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (14.8 * TpCM%)
	    Printer.Print "$2.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"

	    '$5 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (18.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (17.8 * TpCM%)
	    Printer.Print "$5.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	
	    'Print line under Loft Auction
	    DrawWidth = 30
	    'Printer.Line (Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    Printer.Line (MlMargin! + Col_15!, MtMargin! + (19.7 * TpCM%))-(MrMargin!, MtMargin! + (19.7 * TpCM%))
	    
    'Print rest
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.CurrentY = MtMargin! + (20.8 * TpCM%)
    Printer.Print "Any"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Other"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Pools"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "or"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Jackpots"
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    Printer.FontItalic = False
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    Printer.Print "Total No of Birds entered : ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print Str$(RecCnt%);
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontBold = True
    Printer.Print "TOTAL $"
    
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print ""
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
    Printer.FontBold = False
    Printer.Print "No of Cocks...........   No of Hens..............";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Entry fee";
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "@";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Cartage";
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "@";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin!
    Printer.Print "I hereby declare that all the above entries are my own property and are registered";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Pools";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin!
    Printer.Print "in my name in above mentioned club, and that all disputes of whatever";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Other Nominations";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print "kind arising shall be decided by the executive body, whose decision shall be final.";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Coloured Bird";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print "";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Other Noms";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
    Printer.Print "SIGNATURE: ";
    Printer.Print "x.............................................................";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Other";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.FontBold = True
    Printer.Print "TOTAL";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    'Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2z Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage

Next Y%
Printer.EndDoc
DoEvents
RaceSheet11 = True
Exit Function

RS11_err:
    RaceSheet11 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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

End Function

Function RaceSheet12 (ByVal parmStyle As String)
' Style 12 - VHA Stye race sheet (No Breeder, Allows for two rubbers).
' Style 15 - Vic Central Highlands Combine Inc.
'GLRSRaceFormType% = 12 & 15

On Error GoTo RS12_err

Dim cntCocks As Integer
Dim cntHens As Integer
Dim txtStyle As String

txtStyle = parmStyle
RaceSheet12 = False

'Set Start of each column from left of page
Col_1! = 0 * TpCM%
Col_2! = .7 * TpCM%
Col_3! = 1.4 * TpCM%
'Col_4! = 2.25 * TpCM%
Col_4! = 3.85 * TpCM%
Col_5! = 4.7 * TpCM%
Col_6! = 6.4 * TpCM%
Col_7! = 9.45 * TpCM%
Col_8! = 10.05 * TpCM%
Col_9! = 12.5 * TpCM%
Col_10! = 13.3 * TpCM%
Col_11! = 14.1 * TpCM%
Col_12! = 14.9 * TpCM%
Col_13! = 15.7 * TpCM%
Col_14! = 16.5 * TpCM%
Col_15! = 17.6 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = 3 * TpCM%
Row_3! = 3.9 * TpCM%
Row_4! = 4.4 * TpCM%
Row_5! = 4.9 * TpCM%
Row_6! = 6.1 * TpCM%

RowFt1! = 23 * TpCM%
RowFt2! = 24.5 * TpCM%
RowFt3! = 27 * TpCM%
LastHoriz! = 22.4 * TpCM%

MlMargin! = 1 * TpCM%
MtMargin! = .75 * TpCM%
MrMargin! = MlMargin! + (18.5 * TpCM%)
MlOffSet! = .09 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'do new page - format etc, DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    'NEW STUFF
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.FontBold = True
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.FontSize = 8
    Printer.Print "This Entry is subject to the Rules & Regulations of"
    Printer.FontSize = 20
    Printer.CurrentX = MlMargin!
    Printer.Print GLFed$
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.FontSize = 9
    Printer.Print "ENTRY FORM & POOLS"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.Print "RACE POINT: ";
    Printer.FontBold = True
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.Print "DATE: ";
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.Print " Page: ";
    Printer.FontBold = True
    Printer.Print Str(Y%) & " of " & Str(NoPages%)
    Printer.CurrentX = MlMargin!
    Printer.FontBold = False
    Printer.Print "FLYERS NAME: ";
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.FontBold = True
    Printer.Print GLOwner$;
    'Printer.Print "Club: " & GLOwnerClub
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.FontBold = False
    Printer.Print "CLUB: ";
    Printer.FontBold = True
    Printer.Print GLMark
    Printer.CurrentX = MlMargin!
    Printer.FontBold = False
    Printer.Print "ADDRESS: ";
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.FontBold = True
    If GLAddress1 <> "" Then Printer.Print GLAddress1
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    If GLAddress2 <> "" Then Printer.Print GLAddress2
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print GLCity & " " & GLState & " " & GLPostcode
    Printer.CurrentX = MlMargin! + Col_4! + (1 * TpCM%)
    Printer.Print "Phone: " & GLPhoneNo

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS(Optional, mark with 'X' if desired)";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 140
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + LastHoriz!))-(MrMargin, (MtMargin! + LastHoriz!))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    DrawWidth = 8
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentY = 1190
    Printer.FontBold = False
    Printer.Print "No";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Ring Marks";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Year";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Ring No.";

    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Colour";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Sex";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    If txtStyle = "12" Then
	Printer.Print "Race Ring No.";
    Else
	Printer.Print "Race Rubber No.";
    End If
    
    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "Nom";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    If txtStyle = "12" Then
	Printer.Print "50c";
    Else
	Printer.Print "10c";
    End If
    
    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    If txtStyle = "12" Then
	Printer.Print "$1";
    Else
	Printer.Print "20c";
    End If

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    If txtStyle = "12" Then
	Printer.Print "$2";
    Else
	Printer.Print "50c";
    End If

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Lofts"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 30
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    DrawStyle = 0
    DrawWidth = 30
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + RowFt1!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 12
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	     
	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	    Else
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col3*
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
	     Printer.Print tblMaster("mark");

	     'Col4*
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));
	     
	     'Col5*
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col6*
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If
	     
	     'Col7*
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If
	     If tblMaster("codeSex") = "H" Then
		cntHens% = cntHens% + 1
	     ElseIf tblMaster("codeSex") = "C" Then
		cntCocks% = cntCocks% + 1
	     End If
	     'Horizontal line between birds are printed previously
	     'Data2.recordset.MoveNext
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Draw Static Loft Auction/Pools etc.
	    
	    'Please Mark like this Box
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (7 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    'Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY!)'Right Vert
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    ' Cross X lines corner to corner
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))
	    'Printer.Line (TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))-(TLX!, TLY!)
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY!)
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (6.25 * TpCM%)
	    Printer.Print "PLEASE"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "MARK"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "THUS"
	
	    '50c Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (9.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (9.1 * TpCM%)
	    Printer.Print "50c"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$1 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (12.3 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (11.7 * TpCM%)
	    Printer.Print "$1.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$2 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (15.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (14.8 * TpCM%)
	    Printer.Print "$2.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"

	    '$5 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (18.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (17.8 * TpCM%)
	    Printer.Print "$5.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	
	    'Print line under Loft Auction
	    DrawWidth = 30
	    'Printer.Line (Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    Printer.Line (MlMargin! + Col_15!, MtMargin! + (19.7 * TpCM%))-(MrMargin!, MtMargin! + (19.7 * TpCM%))
	    
    'Print rest
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.CurrentY = MtMargin! + (20.8 * TpCM%)
    Printer.Print "Any"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Other"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Pools"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "or"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Jackpots"
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    Printer.FontItalic = False
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    Printer.Print "Total No of Birds entered : ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print Str$(RecCnt%);
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontBold = True
    Printer.Print "TOTAL $"
    
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print ""
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
    Printer.FontBold = False
    Printer.Print "No of Cocks...........   No of Hens..............";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Entry fee";
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "@";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "Cartage";
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "@";
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin!
    Printer.Print "I hereby declare that all the above entries are my own property and are registered";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    If txtStyle = "12" Then
	Printer.Print "Pools";
    Else
	Printer.Print "Pools 10c";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin!
    Printer.Print "in my name in above mentioned club, and that all disputes of whatever";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    If txtStyle = "12" Then
	Printer.Print "Other Nominations";
    Else
	Printer.Print "Pools 20c";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print "kind arising shall be decided by the executive body, whose decision shall be final.";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    If txtStyle = "12" Then
	Printer.Print "Coloured Bird";
    Else
	Printer.Print "Pools 50c";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 8
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.Print "";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    If txtStyle = "12" Then
	Printer.Print "Other Noms";
    Else
	Printer.Print "Other";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
    Printer.Print "SIGNATURE: ";
    Printer.Print "x.............................................................";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    If txtStyle = "12" Then
	Printer.Print "Other";
    Else
	Printer.Print "TOTAL";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! + (1.2 * TpCM%)
    Printer.FontBold = True
    If txtStyle = "12" Then
	Printer.Print "TOTAL";
    Else
	Printer.Print "";
    End If
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "=";
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet! + (1.2 * TpCM%)
    Printer.Print "$"

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    'Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2z Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage

Next Y%
Printer.EndDoc
DoEvents
RaceSheet12 = True
Exit Function

RS12_err:
    RaceSheet12 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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
End Function

Function RaceSheet2 ()
' Style 2 - Created for QPF Race sheet in 1997 which includes Breeder Name
' GLRSRaceFormType% = 2

On Error GoTo RS2_err

RaceSheet2 = False

'Set Start of each column from left of page
Col_1! = 0 * TpCM%
Col_2! = .6 * TpCM%
Col_3! = 3.4 * TpCM%
Col_4! = 4 * TpCM%
Col_5! = 5.6 * TpCM%
Col_6! = 8.05 * TpCM%
Col_7! = 8.9 * TpCM%
Col_8! = 10.15 * TpCM%
Col_9! = 13.3 * TpCM%
Col_10! = 14 * TpCM%
Col_11! = 14.8 * TpCM%
Col_12! = 15.4 * TpCM%
Col_13! = 16 * TpCM%
Col_14! = 16.6 * TpCM%
Col_15! = 17.3 * TpCM%
Col_16! = 18 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = .9 * TpCM%
Row_3! = 1.5 * TpCM%
Row_4! = 2 * TpCM%
Row_5! = 2.5 * TpCM%
Row_6! = 3.7 * TpCM%

If GLRSLpp% > 28 Then
    RowFt1! = 24.3 * TpCM%
    RowFt2! = 25 * TpCM%
    RowFt3! = 25.7 * TpCM%
    RowFt4! = 26.5 * TpCM%
    LastHoriz! = 23.5 * TpCM%
    MtMargin! = .5 * TpCM%
Else
    RowFt1! = 22.9 * TpCM%
    RowFt2! = 23.6 * TpCM%
    RowFt3! = 24.3 * TpCM%
    RowFt4! = 25.1 * TpCM%
    LastHoriz! = 22.1 * TpCM%
    MtMargin! = .75 * TpCM%
End If

MlMargin! = .5 * TpCM%
MrMargin! = MlMargin! + (19.5 * TpCM%)
MlOffSet! = .08 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    'do new page - format etc
    'DO HEADING
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.Print GLFed$
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_2! 'was 890
    Printer.Print "Race Point: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.FontSize = 11
    Printer.CurrentX = (MlMargin! + Col_9!) - (.3 * TpCM%)
    Printer.Print " Date: ";
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.FontSize = 11
    'Printer.CurrentX = MlMargin! + Col_14!
    Printer.Print "  Page: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print Str(Y%) & " of " & Str(NoPages%)

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_16!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 60
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + LastHoriz!))-(MrMargin, (MtMargin! + LastHoriz!))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_3
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_16!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentY = 1190
    Printer.FontBold = False
    Printer.Print "No";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "Colour";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Sex";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Ring No";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Mark";

    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Year";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Rubber";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.Print "Breeder";

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "50c";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "AOC";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "$1";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$2";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "$5";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "$10";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "$20"

    'Col16
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_16! + MlOffSet!
    Printer.Print "Prize $"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 60
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    DrawStyle = 0
    DrawWidth = 30
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + RowFt1!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 14
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    'ButtPrint.Visible = True
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	    
	    Else
	    
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col2
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If

	     'Col3
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If

	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col5
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));

	     'Col8
	     Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - 35
	     If Not IsNull(tblMaster("bredby")) Then
		If Len(tblMaster("bredby")) > 20 Then
		    Printer.FontSize = 7
		Else
		    Printer.FontSize = 8
		End If
		Printer.Print Trim(tblMaster("bredby"));
	     End If

	     'Horizontal line between birds are printed previously
	     'Data2.recordset.MoveNext
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    Printer.FontItalic = False
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    Printer.Print "Total No of Birds entered : ";
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print Str$(RecCnt%)
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.Print "Cost : $............................";
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.FontBold = True
    Printer.Print "TOTAL"
	    
' ++++++++++++++++
    'Draw Static Loft Auction/Pools etc.
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (1.35 * TpCM%)
    Printer.Print "LOFT POOLS:";
    Printer.CurrentX = MlMargin! + (3.4 * TpCM%)
    Printer.Print "50c";
    Printer.CurrentX = MlMargin! + (5.55 * TpCM%)
    Printer.Print "$1.00";
    Printer.CurrentX = MlMargin! + (8.1 * TpCM%)
    Printer.Print "$2.00";
    Printer.CurrentX = MlMargin! + (10.6 * TpCM%)
    Printer.Print "$5.00";
    Printer.CurrentX = MlMargin! + (13.15 * TpCM%)
    Printer.Print "$10.00";
    Printer.CurrentX = MlMargin! + (15.85 * TpCM%)
    Printer.Print "$20.00"
    'Printer.Print "   50c                $1.00                  $2.00         ";
    'Printer.Print "        $5.00                 $10.00                 $20.00"

'Draw Boxs for above
    '50c Box
    TLX! = MlMargin! + (4.1 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$1 Box
    TLX! = MlMargin! + (6.5 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    '$2 Box
    TLX! = MlMargin! + (9.05 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$5 Box
    TLX! = MlMargin! + (11.6 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$10 Box
    TLX! = MlMargin! + (14.3 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$20 Box
    TLX! = MlMargin! + (17 * TpCM%)
    TLY! = LastHoriz! + (1.55 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    'Complete footer
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (2.35 * TpCM%)
    Printer.Print "FLYERS NAME: ";
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print GLOwner$;
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print "SIGNATURE: ...................................................";
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (3.1 * TpCM%)
    Printer.Print "CLUB SECRETARY: ...................................................";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print "CLUB: ";
    If Len(Trim(GLOwnerClub)) > 26 Then
	Printer.FontSize = 8
    ElseIf Len(Trim(GLOwnerClub)) > 12 Then
	Printer.FontSize = 9
    Else
	Printer.FontSize = 11
    End If
    Printer.FontBold = True
    Printer.Print Trim(GLOwnerClub)

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + LastHoriz! + (3.6 * TpCM%)
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2z Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage
Next Y%
Printer.EndDoc
DoEvents
RaceSheet2 = True
Exit Function

RS2_err:
    RaceSheet2 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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
End Function

Function RaceSheet3 ()
' Style 3 - NEW QPF Race sheet which includes Breeder Name
' GLRSRaceFormType% = 3
' 9/9/00

On Error GoTo RS3_err

RaceSheet3 = False

'Set Start of each column from left of page
Col_1! = .1 * TpCM%
Col_2! = .7 * TpCM% 'No.
Col_3! = 3.6 * TpCM% 'Colour
Col_4! = 4.2 * TpCM% 'Sex
Col_5! = 5.9 * TpCM%
Col_6! = 8.4 * TpCM%
Col_7! = 9.3 * TpCM%
Col_8! = 10.8 * TpCM%
Col_9! = 14.15 * TpCM% '50
Col_10! = 14.85 * TpCM% 'AOC
Col_11! = 15.65 * TpCM%
Col_12! = 16.25 * TpCM%
Col_13! = 16.95 * TpCM%
Col_14! = 17.65 * TpCM%
Col_15! = 18.45 * TpCM%
Col_16! = 19.25 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = .9 * TpCM%
Row_3! = 1.5 * TpCM%
Row_4! = 2 * TpCM%
Row_5! = 2.5 * TpCM%
Row_6! = 3.7 * TpCM%

If GLRSLpp% > 28 Then
    RowFt1! = 24.3 * TpCM%
    RowFt2! = 25 * TpCM%
    RowFt3! = 25.7 * TpCM%
    RowFt4! = 26.5 * TpCM%
    LastHoriz! = 23.5 * TpCM%
    MtMargin! = .5 * TpCM%
Else
    RowFt1! = 22.9 * TpCM%
    RowFt2! = 23.6 * TpCM%
    RowFt3! = 24.3 * TpCM%
    RowFt4! = 25.1 * TpCM%
    LastHoriz! = 22.1 * TpCM%
    MtMargin! = .75 * TpCM%
End If

MlMargin! = .5 * TpCM%
MrMargin! = MlMargin! + (19.5 * TpCM%)
MlOffSet! = .08 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.Print GLFed$
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_2! 'was 890
    Printer.Print "Please enter the following birds for - Race Point: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.FontSize = 11
    Printer.CurrentX = MlMargin! + Col_9! - (1 * TpCM%)
    Printer.Print " Date: ";
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.FontSize = 11
    'Printer.CurrentX = MlMargin! + Col_14! - (.4 * TpCM%)
    Printer.Print "  Page: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print Str(Y%) & " of " & Str(NoPages%)

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_16!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 60
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + LastHoriz!))-(MrMargin, (MtMargin! + LastHoriz!))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_3
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_16!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    Printer.FontBold = False
    Printer.Print "No";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "Colour";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Sex";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Ring No";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Mark";

    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Year";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Rubber";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.Print "Breeder";

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "50c";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "AOC";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "$1";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$2";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "$5";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "$10";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "$20"

    'Col16
    'Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentX = MlMargin! + Col_16! + MlOffSet!
    'Printer.Print "Prize $"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 60
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    'DrawStyle = 0
    'DrawWidth = 30
    'Printer.CurrentX = MlMargin! + Col_8!
    'Printer.CurrentY = MtMargin! + RowFt1!
    'Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 14
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	    
	    Else
	    
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col2
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If

	     'Col3
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If

	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col5
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));

	     'Col8
	     Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - 35
	     If Not IsNull(tblMaster("bredby")) Then
		If Len(tblMaster("bredby")) > 16 Then
		    Printer.FontSize = 8
		Else
		    Printer.FontSize = 9
		End If
		Printer.Print Trim(tblMaster("bredby"));
	     End If

	     'Horizontal line between birds are printed previously
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    'Printer.FontItalic = False
    'Printer.FontSize = 9
    'Printer.FontBold = False
    'Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    'Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    'Printer.Print "Total No of Birds entered : ";
    'Printer.FontSize = 10
    'Printer.FontBold = True
    'Printer.Print Str$(RecCnt%)
    'Printer.FontSize = 9
    'Printer.FontBold = False
    'Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    'Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    'Printer.Print "Cost : $............................";
    'Printer.FontSize = 9
    'Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    'Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    'Printer.FontBold = True
    'Printer.Print "TOTAL"
    
    'Draw Static Loft Auction/Pools etc.
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (.8 * TpCM%)
    Printer.Print "LOFT POOLS:";
    Printer.CurrentX = MlMargin! + (3.4 * TpCM%)
    Printer.Print "50c";
    Printer.CurrentX = MlMargin! + (5.5 * TpCM%)
    Printer.Print "$1.00";
    Printer.CurrentX = MlMargin! + (8 * TpCM%)
    Printer.Print "$2.00";
    Printer.CurrentX = MlMargin! + (10.5 * TpCM%)
    Printer.Print "$5.00";
    Printer.CurrentX = MlMargin! + (13.1 * TpCM%)
    Printer.Print "$10.00";
    Printer.CurrentX = MlMargin! + (15.8 * TpCM%)
    Printer.Print "$20.00";
    'Printer.Print "   50c                $1.00                  $2.00         ";
    'Printer.Print "        $5.00                 $10.00                 $20.00"

'Draw Boxs for above
    '50c Box
    TLX! = MlMargin! + (4.1 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$1 Box
    TLX! = MlMargin! + (6.5 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    '$2 Box
    TLX! = MlMargin! + (9.05 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$5 Box
    TLX! = MlMargin! + (11.6 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$10 Box
    TLX! = MlMargin! + (14.4 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$20 Box
    TLX! = MlMargin! + (17.1 * TpCM%)
    TLY! = LastHoriz! + (1 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    'Complete footer
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (2 * TpCM%)
    Printer.Print "FLYERS NAME: ";
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print GLOwner$;
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.Print "SIGNATURE: ...................................................";
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (3.1 * TpCM%)
    Printer.Print "CLUB SECRETARY: ...................................................";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.Print "Club: ";
    If Len(Trim(GLOwnerClub)) > 23 Then
	Printer.FontSize = 8
    ElseIf Len(Trim(GLOwnerClub)) > 14 Then
	Printer.FontSize = 10
    Else
	Printer.FontSize = 11
    End If
    Printer.FontBold = True
    Printer.Print GLOwnerClub

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + LastHoriz! + (3.6 * TpCM%)
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2z Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage
Next Y%
Printer.EndDoc
DoEvents
RaceSheet3 = True
Exit Function

RS3_err:
    RaceSheet3 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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
	'Resume Next
    End Select
End Function

Function RaceSheet4 ()
' Style 4 - NEW QPF Race sheet which includes Breeder Name
' GLRSRaceFormType% = 4
' 23/9/00

On Error GoTo RS4_err

RaceSheet4 = False
'Set Start of each column from left of page
Col_1! = .1 * TpCM%
Col_2! = .7 * TpCM% 'No.
Col_3! = 3.6 * TpCM% 'Colour
Col_4! = 4.25 * TpCM% 'Sex
Col_5! = 5.95 * TpCM%
Col_6! = 8.45 * TpCM%
Col_7! = 9.35 * TpCM%
Col_8! = 10.85 * TpCM%
Col_9! = 14.2 * TpCM% '50
Col_10! = 14.9 * TpCM% 'AOC
Col_11! = 15.75 * TpCM%
Col_12! = 16.45 * TpCM%
Col_13! = 17.15 * TpCM%
Col_14! = 17.9 * TpCM%
Col_15! = 18.75 * TpCM%
Col_16! = 19.605 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = .9 * TpCM%
Row_3! = 1.5 * TpCM%
Row_4! = 2 * TpCM%
Row_5! = 2.5 * TpCM%
Row_6! = 3.7 * TpCM%

If GLRSLpp% > 28 Then
    'RowFt1! = 24.3 * TpCM%
    LastHoriz! = 23.5 * TpCM%
Else
    'RowFt1! = 22.9 * TpCM%
    LastHoriz! = 22.1 * TpCM%
End If

MlMargin! = .5 * TpCM%
MtMargin! = .5 * TpCM%
MrMargin! = MlMargin! + (19.605 * TpCM%)
MlOffSet! = .08 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.FontBold = True
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.Print GLFed$
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.FontItalic = True
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_2! 'was 890
    Printer.Print "Please enter the following birds for:- ";
    Printer.FontItalic = False
    Printer.Print "RACE POINT ";
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontSize = 12
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.FontSize = 11
    Printer.CurrentX = MlMargin! + Col_9! - (1 * TpCM%)
    Printer.Print " DATE: ";
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.FontSize = 11
    'Printer.CurrentX = MlMargin! + Col_14! - (.4 * TpCM%)
    Printer.Print "  PAGE: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print Str(Y%) & " of " & Str(NoPages%)

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

'Complete sub heading
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.CurrentY = (MtMargin! + Row_4! - MtOffSet!) + (.1 * TpCM%)
    
    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "RING";
    
    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "RUBBER";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.Print "NAME OF";
    
    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "AOC";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "$";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "$";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "$";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "$"
    
    'Print bottom bold line
    DrawWidth = 100
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + LastHoriz!))-(MrMargin, (MtMargin! + LastHoriz!))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_3
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))
    Printer.CurrentX = MlMargin! + Col_16!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + LastHoriz!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    Printer.FontBold = True
    Printer.Print "No.";

    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "COLOUR";

    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "SEX";

    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "NUMBER";

    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "MARK";

    'Col6
    Printer.CurrentX = (MlMargin! + Col_6! + MlOffSet!) - (.05 * TpCM%)
    Printer.Print "YEAR";

    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "No.";

    'Col8
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.Print "BREEDER";

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "50c";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "$1.00";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "1.00";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "2.00";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "5.00";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "10.00";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "20.00"

    'Col16
    'Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentX = MlMargin! + Col_16! + MlOffSet!
    'Printer.Print "Prize $"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 60
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-(MrMargin!, Printer.CurrentY)
    Next X%
    
    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 14
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print DSRTrans("ringno");

	     'Col5
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print DSRTrans("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	    
	    Else
	    
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     
	     'Col2
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If

	     'Col3
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If

	     'Col4
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     Printer.Print tblMaster("ringno");

	     'Col5
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     Printer.Print tblMaster("mark");

	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));

	     'Col8
	     Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - 35
	     If Not IsNull(tblMaster("bredby")) Then
		If Len(tblMaster("bredby")) > 18 Then
		    Printer.FontSize = 8
		Else
		    Printer.FontSize = 9
		End If
		Printer.Print Trim(tblMaster("bredby"));
	     End If

	     'Horizontal line between birds are printed previously
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    
    'Draw Static Loft Auction/Pools etc.
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (.7 * TpCM%)
    Printer.Print "LOFT POOLS:";
    Printer.CurrentX = MlMargin! + (3.4 * TpCM%)
    Printer.Print "50c";
    Printer.CurrentX = MlMargin! + (5.5 * TpCM%)
    Printer.Print "$1.00";
    Printer.CurrentX = MlMargin! + (8 * TpCM%)
    Printer.Print "$2.00";
    Printer.CurrentX = MlMargin! + (10.5 * TpCM%)
    Printer.Print "$5.00";
    Printer.CurrentX = MlMargin! + (13.1 * TpCM%)
    Printer.Print "$10.00";
    Printer.CurrentX = MlMargin! + (15.8 * TpCM%)
    Printer.Print "$20.00";
    'Printer.Print "   50c                $1.00                  $2.00         ";
    'Printer.Print "        $5.00                 $10.00                 $20.00"

'Draw Boxs for above
    '50c Box
    TLX! = MlMargin! + (4.1 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$1 Box
    TLX! = MlMargin! + (6.5 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    '$2 Box
    TLX! = MlMargin! + (9.05 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$5 Box
    TLX! = MlMargin! + (11.6 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$10 Box
    TLX! = MlMargin! + (14.4 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor
    
    '$20 Box
    TLX! = MlMargin! + (17.1 * TpCM%)
    TLY! = LastHoriz! + (.9 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.9 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.9 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.9 * TpCM), TLY!)-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.9 * TpCM%))-(TLX! + (.9 * TpCM%), TLY! + (.9 * TpCM%))'Bot Hor

    'Complete footer
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (1.8 * TpCM%)
    Printer.Print "Name: ";
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print GLOwner$;
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.Print "SIGNATURE: ...................................................";
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.CurrentY = MtMargin! + LastHoriz! + (2.7 * TpCM%)
    Printer.Print "Club Secretary: ...................................................";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.Print "CLUB: ";
    If Len(Trim(GLOwnerClub)) > 23 Then
	Printer.FontSize = 9
    ElseIf Len(Trim(GLOwnerClub)) > 14 Then
	Printer.FontSize = 10
    Else
	Printer.FontSize = 11
    End If
    Printer.FontBold = True
    Printer.Print GLOwnerClub

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + LastHoriz! + (3.2 * TpCM%)
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2zc Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage
Next Y%
Printer.EndDoc
DoEvents
RaceSheet4 = True
Exit Function

RS4_err:
    RaceSheet4 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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
	'Resume Next
    End Select

End Function

Function RaceSheet5 ()
' Style 5 - TDPF Stye race sheet (No Breeder, Allows for two rubbers).
'GLRSRaceFormType% = 1

On Error GoTo RS5_err

RaceSheet5 = False

'Set Start of each column from left of page
Col_1! = 0 * TpCM%
Col_2! = .7 * TpCM%
Col_3! = 3.1 * TpCM%
Col_4! = 4 * TpCM%
Col_5! = 7 * TpCM%
Col_6! = 7.7 * TpCM%
Col_7! = 9.5 * TpCM%
Col_8! = 10.95 * TpCM%
Col_9! = 12.5 * TpCM%
Col_10! = 13.2 * TpCM%
Col_11! = 14.1 * TpCM%
Col_12! = 14.9 * TpCM%
Col_13! = 15.7 * TpCM%
Col_14! = 16.5 * TpCM%
Col_15! = 17.5 * TpCM%

Row_1! = 0 * TpCM%
Row_2! = .9 * TpCM%
Row_3! = 1.5 * TpCM%
Row_4! = 2 * TpCM%
Row_5! = 2.5 * TpCM%
Row_6! = 3.7 * TpCM%

If GLRSLpp% > 28 Then
    RowFt1! = 24.2 * TpCM%
    RowFt2! = 24.7 * TpCM%
    RowFt3! = 25.4 * TpCM%
    RowFt4! = 26.2 * TpCM%
    LastHoriz! = 23.5
    MtMargin! = .5 * TpCM%
Else
    RowFt1! = 22.9 * TpCM%
    RowFt2! = 23.6 * TpCM%
    RowFt3! = 24.3 * TpCM%
    RowFt4! = 25.1 * TpCM%
    LastHoriz! = 22.1
    MtMargin! = .5 * TpCM%
End If

MlMargin! = 1 * TpCM%
MrMargin! = MlMargin! + (18.5 * TpCM%)
MlOffSet! = .09 * TpCM%
MtOffSet! = .45 * TpCM%
Mline! = .7 * TpCM%
Flag% = 0

For Y% = 1 To NoPages%
    'do new page - format etc, DO HEADING
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Headings"
    
    ' Select font and point size.
    Printer.FontName = "Arial"
    Printer.ScaleMode = 1  ' twips

    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_1! 'was 350
    Printer.Print GLFed$
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + Row_2! 'was 890
    Printer.Print "Please enter the following birds for - Race Point: ";
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.Print mRacePoint.Text; 'Print race point name
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.Print " Date: ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print mRaceDate.Text;
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.Print " Page: ";
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.Print Str(Y%) & " of " & Str(NoPages%)

    'Print Top Bold Line across page
    DrawWidth = 40
    Printer.Line (MlMargin!, (MtMargin! + Row_3!))-(MrMargin!, (MtMargin! + Row_3!))

    'Complete sub heading
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.CurrentY = MtMargin! + Row_4! - MtOffSet!
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "POOLS";
    
    'Draw line under Pools
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_4!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    
    'Print bottom bold line
    DrawWidth = 40
    Printer.Line ((MlMargin! + Col_1!), (MtMargin! + (LastHoriz! * TpCM%)))-(MrMargin, (MtMargin! + (LastHoriz! * TpCM%)))

    'Print all Vertical lines
    'X axis will stay the same for all lines only Y axis will change
    Printer.CurrentX = MlMargin! + Col_2!
    Printer.CurrentY = MtMargin! + Row_3!
    DrawWidth = 8
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_3!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_4!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_5!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_6!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_7!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + (LastHoriz! * TpCM%)))
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + Row_5!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_9!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_10!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_11!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_12!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_13!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_14!
    Printer.CurrentY = MtMargin! + Row_4!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))
    Printer.CurrentX = MlMargin! + Col_15!
    Printer.CurrentY = MtMargin! + Row_3!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX, (MtMargin! + RowFt1!))

    'Col 1
    Printer.FontSize = 9
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin! + Row_5! - MtOffSet!
    'Printer.CurrentY = 1190
    Printer.FontBold = False
    Printer.Print "No";
    'Col2
    Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
    Printer.Print "Mark";
    'Col3
    Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 30
    Printer.Print "Year";
    'Col4
    Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
    Printer.Print "Colour";
    'Col5
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.Print "Sex";
    'Col6
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
    Printer.Print "Ring No";
    'Col7
    Printer.CurrentX = MlMargin! + Col_7! + MlOffSet!
    Printer.Print "Rubber Nos";

    'Col8

    'Col9
    Printer.CurrentX = MlMargin! + Col_9! + MlOffSet!
    Printer.Print "50c";

    'Col10
    Printer.CurrentX = MlMargin! + Col_10! + MlOffSet!
    Printer.Print "COL";

    'Col11
    Printer.CurrentX = MlMargin! + Col_11! + MlOffSet!
    Printer.Print "$1";

    'Col12
    Printer.CurrentX = MlMargin! + Col_12! + MlOffSet!
    Printer.Print "$2";

    'Col13
    Printer.CurrentX = MlMargin! + Col_13! + MlOffSet!
    Printer.Print "$5";

    'Col14
    Printer.CurrentX = MlMargin! + Col_14! + MlOffSet!
    Printer.Print "";

    'Col15
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Lofts"

    'Draw line under above sub heading
    Printer.CurrentX = MlMargin! + Col_1!
    Printer.CurrentY = MtMargin! + Row_5!
    DrawWidth = 30
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)

    'Print Horizontal Row lines under entries for entire page
    DrawStyle = 2
    DrawWidth = 2
    Printer.CurrentY = MtMargin! + Row_5!
    For X% = 1 To (Line_Count% - 1)' was 30
	MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Lines " & Str(X%)
	Printer.CurrentX = MlMargin! + Col_1!
	Printer.CurrentY = Printer.CurrentY + Mline!
	Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MlMargin! + Col_15!), Printer.CurrentY)
    Next X%
    
    'Draw Solid Line under Total
    DrawStyle = 0
    DrawWidth = 30
    Printer.CurrentX = MlMargin! + Col_8!
    Printer.CurrentY = MtMargin! + RowFt1!
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-((MrMargin!), Printer.CurrentY)

    For X% = 1 To Line_Count%
    'Read nominated birds until eof, read next bird & print
	DoEvents
	If Flag% = -1 Then
	    MsgBox "Print Aborted by User", 16, "Aborted"
	    Printer.CurrentX = MlMargin!
	    Printer.Print " "
	    Printer.FontSize = 12
	    Printer.Print "*** WARNING - PRINT ABORTED BY USER ***"
	    Printer.NewPage
	    Printer.EndDoc
	    cmdAbort.Visible = False
	    Screen.MousePointer = 0
	    MsgPanel.Caption = ""
	    Exit Function
	End If
	If DSRTrans.EOF Then
	    X% = Line_Count%
	Else
	    CntBirds% = CntBirds% + 1
	    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Bird " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno") & " - Line: " & Str(X%)
	    If X% = 1 Then 'first time through
		Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
		Printer.CurrentY = MtMargin! + Row_5! + Mline! - MtOffSet!
	    Else
		Printer.CurrentY = Printer.CurrentY + Mline!
	    End If

	    Printer.FontBold = False
	    Printer.FontSize = 10
	    
	    'Seek Master
	    tblMaster.Seek "=", DSRTrans("Yr"), DSRTrans("Mark"), DSRTrans("ringno")
	    
	    '23/6/97 if Master record not on file omit bird
	    If tblMaster.NoMatch Then
	     MsgBox "Bird nominated in race but NO MASTER record - Error #30001 - To Fix Goto Breeding Module and add this bird - Race sheet will complete when you click OK", 0, "Ring ID " & Str(DSRTrans("Yr")) & "/" & DSRTrans("Mark") & "/" & DSRTrans("ringno")
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     'Col2
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     Printer.Print DSRTrans("mark");
	     'Col3
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 35
	     Printer.Print Trim(Str(DSRTrans("yr")));
	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     Printer.Print DSRTrans("ringno");
	    Else
	     'Col 1
	     Printer.FontBold = False
	     Printer.FontSize = 9
	     Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
	     Printer.Print CntBirds%;
	     'Col2
	     If Len(tblMaster("mark")) > 10 Then
		Printer.FontSize = 8
	     ElseIf Len(tblMaster("mark")) > 9 Then
		Printer.FontSize = 9
	     Else
		Printer.FontSize = 10
	     End If
	     Printer.CurrentX = MlMargin! + Col_2! + MlOffSet!
	     Printer.Print tblMaster("mark");
	     'Col3
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_3! + MlOffSet! - 35
	     Printer.Print Trim(Str(tblMaster("yr")));
	     'Col4
	     Printer.CurrentX = MlMargin! + Col_4! + MlOffSet!
	     tblColour.Seek "=", tblMaster("codecolour")
	     If tblColour.NoMatch Then
		Printer.Print tblMaster("codecolour");
	     ElseIf Len(tblColour("desc")) > 14 Then
		Printer.FontSize = 9
		Printer.Print tblColour("desc");
	     Else
		Printer.FontSize = 10
		Printer.Print tblColour("desc");
	     End If
	     'Col5
	     'Get Sex from Master
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
	     If tblMaster("codeSex") = "U" Then
		'Leave Blank
	     Else
		Printer.Print tblMaster("codeSex");
	     End If
	     'Col6
	     Printer.FontSize = 10
	     Printer.CurrentX = MlMargin! + Col_6! + MlOffSet!
	     Printer.Print tblMaster("ringno");
	     'Horizontal line between birds are printed previously
	     'Data2.recordset.MoveNext
	    End If
	    DSRTrans.MoveNext
	    DoEvents
	End If
    Next X%
    
    'Draw Static Loft Auction/Pools etc.
	    
	    'Please Mark like this Box
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (4.5 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    'Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY!)'Right Vert
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    ' Cross X lines corner to corner
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))
	    'Printer.Line (TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))-(TLX!, TLY!)
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY!)
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (3.75 * TpCM%)
	    Printer.Print "PLEASE"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "MARK"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "THUS"
	
	    '50c Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (7.4 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (6.6 * TpCM%)
	    Printer.Print "50c"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$1 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (9.8 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (9.2 * TpCM%)
	    Printer.Print "$1.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	    
	    '$2 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (12.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (12.3 * TpCM%)
	    Printer.Print "$2.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"

	    '$5 Loft
	    TLX! = MlMargin! + Col_15! + MlOffSet!
	    TLY! = MtMargin! + (15.9 * TpCM%)
	    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
	    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
	    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
	    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
	    Printer.FontSize = 6
	    Printer.FontBold = False
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.CurrentY = MtMargin! + (15.3 * TpCM%)
	    Printer.Print "$5.00"
	    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
	    Printer.Print "LOFT"
	
	    'Print line under Loft Auction
	    DrawWidth = 30
	    'Printer.Line (Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    Printer.Line (MlMargin! + Col_15!, MtMargin! + (17.2 * TpCM%))-(MrMargin!, MtMargin! + (17.2 * TpCM%))
	    
    'Print rest
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.CurrentY = MtMargin! + (18.3 * TpCM%)
    Printer.Print "Any"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Other"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Pools"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "or"
    Printer.CurrentX = MlMargin! + Col_15! + MlOffSet!
    Printer.Print "Jackpots"
    
    'Print Page footer
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - Footer"
    Printer.FontItalic = False
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + MlOffSet!
    Printer.CurrentY = MtMargin + RowFt1! - MtOffSet!
    Printer.Print "Total No of Birds entered : ";
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print Str$(RecCnt%)
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_5! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.Print "Cost : ........................";
    Printer.FontSize = 10
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet!
    Printer.CurrentY = MtMargin! + RowFt1! - MtOffSet!
    Printer.FontBold = True
    Printer.Print "TOTAL"
	    
    'CLUB Box
    TLX! = Col_1! + MlMargin!
    TLY! = RowFt2! + MtMargin! - (.5 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.6 * TpCM), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + (.9 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt2! - MtOffSet!
    Printer.Print "CLUB"
    'Rest of Line
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt2! - MtOffSet!
    Printer.Print "NAME: ";
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.Print GLOwner

    'FED Box
    TLX! = Col_1! + MlMargin!
    TLY! = RowFt3! + MtMargin! - (.5 * TpCM%)
    Printer.Line (TLX!, TLY!)-(TLX!, TLY! + (.6 * TpCM%)) 'Left Vert
    Printer.Line (TLX!, TLY!)-(TLX! + (.6 * TpCM%), TLY!) 'Top Hor
    Printer.Line (TLX! + (.6 * TpCM%), TLY!)-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Right Vert
    Printer.Line (TLX!, TLY! + (.6 * TpCM%))-(TLX! + (.6 * TpCM%), TLY! + (.6 * TpCM%))'Bot Hor
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_1! + (.9 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "FEDERATION"
    'Rest of line
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet!
    Printer.Print "SIGNATURE: ";
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt3! - MtOffSet! + 120
    Printer.Print "x....................................................";
    Printer.FontBold = False
    Printer.CurrentX = MlMargin! + Col_6! + MlOffSet! - (.5 * TpCM%)
    Printer.CurrentY = MtMargin! + RowFt4! - MtOffSet!
    Printer.Print "CLUB: ";
    Printer.FontBold = True
    If Len(Trim(GLOwnerClub)) > 26 Then
	Printer.FontSize = 9
    ElseIf Len(Trim(GLOwnerClub)) > 14 Then
	Printer.FontSize = 10
    Else
	Printer.FontSize = 11
    End If
    Printer.CurrentX = MlMargin! + Col_8! + MlOffSet! - (.5 * TpCM%)
    Printer.Print GLOwnerClub
    Printer.FontBold = False

    'Print Software Serial No.
    Printer.FontBold = False
    Printer.FontSize = 4
    Printer.CurrentX = MlMargin!
    Printer.CurrentY = MtMargin! + ((LastHoriz! + 2.7) * TpCM%)
    Printer.Print "Software Registered to " & GLOwner$ & " Serial No. " & LicenseDisplay(GLSerialNo$)
    Printer.CurrentX = MlMargin!
    Printer.Print "Style " & GLRSRaceFormType% & " - Produced by RPMS Software V3.2zc Copyright K Malone (racingpigeons@telstra.com) 1988 - 2001, All Rights reserved"
    MsgPanel.Caption = CapMsg$ & Str$(Y%) & " of " & Str$(NoPages%) & " - End Page"
    DoEvents
    Printer.NewPage
Next Y%
Printer.EndDoc
DoEvents
RaceSheet5 = True
Exit Function

RS5_err:
    RaceSheet5 = False
    Select Case Err
    Case 482
	MsgBox "Error " & Err & "," & Val(Error) & ", at line " & Erl, 16, "NO PRINTER DRIVER INSTALLED"
	Exit Function
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

End Function

