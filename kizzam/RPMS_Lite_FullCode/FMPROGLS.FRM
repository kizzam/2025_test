VERSION 2.00
Begin Form FMprodg 
   Caption         =   "Prodgeny Listing"
   ClientHeight    =   4020
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   7365
   Height          =   4425
   Left            =   1575
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1590
   Width           =   7485
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   7335
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin SSCommand But3Cancel 
         Caption         =   "&Cancel"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin SSCommand But3Prt 
         Caption         =   "Print"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin SSPanel Panel3D5 
         Alignment       =   5  'Right Justify - BOTTOM
         BackColor       =   &H00C0C0C0&
         BevelOuter      =   0  'None
         Caption         =   "Ring:"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin SSPanel Panel3D4 
         Alignment       =   5  'Right Justify - BOTTOM
         BackColor       =   &H00C0C0C0&
         BevelOuter      =   0  'None
         Caption         =   "Mark:"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin SSPanel Panel3D3 
         Alignment       =   5  'Right Justify - BOTTOM
         BackColor       =   &H00C0C0C0&
         BevelOuter      =   0  'None
         Caption         =   "Year:"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin TextBox Text3 
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   240
         Width           =   1455
      End
      Begin TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   240
         Width           =   1215
      End
      Begin TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin Grid Grid1 
         Height          =   1935
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   6375
      End
   End
End
Option Explicit
Dim Myr() As Integer
Dim Mmark() As String
Dim Mring() As String
Dim Mcolour() As String
Dim Mcolcode() As String
Dim Meyesign() As String
Dim Msex() As String
Dim Mbredby() As String
Dim Mbredfor() As String
Dim Mbloodline() As String
Dim Mname() As String
Dim Mnotes() As String
Dim Mstatus() As Integer
Dim TblMaster As Table
Dim TblColour As Table
Dim DB As Database
Dim TblPedg As Table
Dim Mcnt As Integer
Dim Mcnt2 As Integer
Dim Mcnt3 As Integer
Dim Minfo As Integer
Dim MOyr As String
Dim MOmark As String
Dim MOring As String

Sub But3Cancel_Click ()
Unload FMprodg
End Sub

Sub Disp_Prodg ()
ReDim Myr(63)
ReDim Mmark(63)
ReDim Mring(63)
ReDim Mcolour(63)
ReDim Mcolcode(63)
ReDim Meyesign(63)
ReDim Msex(63)
ReDim Mbredby(63)
ReDim Mbredfor(63)
ReDim Mbloodline(63)
ReDim Mname(63)
ReDim Mnotes(63)
ReDim Mstatus(63)

Minfo = -1 'Info is ON, Screen Cluttered

Mcnt% = 0
Do While Mcnt% < 63
    Mcnt% = Mcnt% + 1
    Myr%(Mcnt%) = 0
    Mmark$(Mcnt%) = ""
    Mcolcode$(Mcnt%) = ""
    Mcolour$(Mcnt%) = ""
    Meyesign$(Mcnt%) = ""
    Msex$(Mcnt%) = ""
    Mring$(Mcnt%) = ""
    Mbredby$(Mcnt%) = ""
    Mbredfor$(Mcnt%) = ""
    Mbloodline$(Mcnt%) = ""
    Mnotes$(Mcnt%) = ""
    Mname$(Mcnt%) = ""
    Mstatus%(Mcnt%) = False
Loop

Mcnt% = 0
Do While Mcnt% < 30  'Bird + 30 ancestors Bird(0) to Bird(29)
    Fbirdymr(Mcnt%).Text = ""
    Mcnt% = Mcnt% + 1
Loop

For Mcnt% = 0 To 13 Step 1  'only Bird + 14 ancestors Bird(0) to Bird (13)
    Fbirdbb(Mcnt%).Text = ""
Next Mcnt%

For Mcnt% = 0 To 13 Step 1  'only Bird + 14 ancestors Bird(0) to Bird (13)
    Panel3D2(Mcnt%).Caption = ""
Next Mcnt%

For Mcnt% = 0 To 5 Step 1   'bloodline display
    Panel3D3(Mcnt%).Caption = ""
Next Mcnt%
Fnotes(0) = ""
Fnotes(1) = ""

TblMaster.Seek "=", Fyr.Text, Fmark.Text, Fring.Text 'Find the required record.
If TblMaster.NoMatch Then
    MsgBox "Bird NOT found", 4, "ERROR"
    MousePointer = 0
    Unload FMpedg
    Screen.MousePointer = 0
    Exit Sub
End If

'store info on bird u are getting pedigree for
Mstatus(1) = True
Myr(1) = TblMaster("yr")
Mmark(1) = TblMaster("mark")
Mring(1) = TblMaster("ringno")
Mcolcode$(1) = TblMaster("codecolour")
If IsNull(TblMaster("codecolour")) Then
    Mcolour$(1) = "Colour unknown"
Else
    Mcolour$(1) = TblMaster("codecolour")
    TblColour.Seek "=", TblMaster("codecolour")
    If TblColour.NoMatch Then
    Else
        Mcolour$(1) = TblColour("desc")
    End If
End If
If IsNull(TblMaster("codeeye")) Then
    Meyesign$(1) = ""
Else
    Meyesign$(1) = TblMaster("codeeye")
End If

If TblMaster("codesex") = "H" Then
    Msex$(1) = "Hen"
ElseIf TblMaster("codesex") = "C" Then
    Msex(1) = "Cock"
Else
    Msex(1) = "Unsexed"
End If
If IsNull(TblMaster("bredby")) Then
    Mbredby(1) = ""
Else
    Mbredby(1) = TblMaster("bredby")
End If
If IsNull(TblMaster("bredfor")) Then
    Mbredfor(1) = ""
Else
    Mbredfor(1) = TblMaster("bredfor")
End If
If IsNull(TblMaster("bloodline")) Then
    Mbloodline(1) = ""
Else
    Mbloodline(1) = TblMaster("bloodline")
End If
If IsNull(TblMaster("birdname")) Then
    Mname(1) = ""
Else
    Mname(1) = TblMaster("birdname")
End If
If IsNull(TblMaster("notes")) Then
    Mnotes(1) = ""
Else
    Mnotes(1) = TblMaster("notes")
End If

'Store Sire in bird(2) Screen field (0)
If IsNull(TblMaster("sireyr")) Then
    Myr(2) = Val("")
Else
    Myr(2) = TblMaster("sireyr")
End If
If IsNull(TblMaster("siremark")) Then
    Mmark(2) = ""
Else
    Mmark(2) = TblMaster("siremark")
End If
If IsNull(TblMaster("sireringno")) Then
    Mring(2) = ""
Else
    Mring(2) = TblMaster("sireringno")
End If
Fbirdymr(0).Text = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(Myr(2))), 2) & " " & Trim(Mmark(2)) & " " & Trim(Mring(2)))
ButtSire.Caption = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(Myr(2))), 2) & " " & Trim(Mmark(2)) & " " & Trim(Mring(2)))

'Dam in bird(3) Screen field (1)
If IsNull(TblMaster("damyr")) Then
    Myr(3) = Val("")
Else
    Myr(3) = TblMaster("damyr")
End If
If IsNull(TblMaster("dammark")) Then
    Mmark(3) = ""
Else
    Mmark(3) = TblMaster("dammark")
End If
If IsNull(TblMaster("damringno")) Then
    Mring(3) = ""
Else
    Mring(3) = TblMaster("damringno")
End If
Fbirdymr(1).Text = Trim(Right$(Trim(Str$(Myr(3))), 2) & " " & Trim(Mmark(3)) & " " & Trim(Mring(3)) & " " & Trim(Mcolcode(3)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(Myr(3))), 2) & " " & Trim(Mmark(3)) & " " & Trim(Mring(3)) & " " & Trim(Mcolcode$(3)) & " " & Trim(Mcolcode(3)))

Mcnt% = 3 'Record found
Mcnt2% = 1 'Parents of record found, start at 2
Mcnt3% = 1   'screen fields start at 2, already have sire(0) & dam(1)

'Get all breeding info & put in an array
Do While Mcnt% < 63 'was 31
    Mcnt2% = Mcnt2% + 1
    If Mstatus(Int(Mcnt2% / 2)) = True Then      '?((Mcnt2/2)-1)
        TblMaster.Seek "=", Myr(Mcnt2%), Mmark(Mcnt2%), Mring(Mcnt2%) 'Find the required record.
        If TblMaster.NoMatch Then
            Mcnt% = Mcnt% + 2
            Mcnt3% = Mcnt3% + 2
        Else
            'Key YMR already stored as we used it to find record, from parent details
            Mstatus(Mcnt2%) = True
            Mcolcode(Mcnt2%) = TblMaster("codecolour")
            If IsNull(TblMaster("codecolour")) Then
                Mcolour(Mcnt2%) = "Colour unknown"
            Else
                Mcolour(Mcnt2%) = TblMaster("codecolour")
                TblColour.Seek "=", TblMaster("codecolour")
                If TblColour.NoMatch Then
                Else
                Mcolour(Mcnt2%) = TblColour("desc")
                End If
            End If
            If IsNull(TblMaster("codeeye")) Then
                Meyesign(Mcnt2%) = ""
            Else
                Meyesign(Mcnt2%) = TblMaster("codeeye")
            End If
            If IsNull(TblMaster("codesex")) Then
                Msex(Mcnt2%) = ""
            Else
                Msex(Mcnt2%) = TblMaster("codesex")
            End If
            If IsNull(TblMaster("bredby")) Then
                Mbredby(Mcnt2%) = ""
            Else
                Mbredby(Mcnt2%) = TblMaster("bredby")
            End If
            If IsNull(TblMaster("bredfor")) Then
                Mbredfor(Mcnt2%) = ""
            Else
                Mbredfor(Mcnt2%) = TblMaster("bredfor")
            End If
            If IsNull(TblMaster("bloodline")) Then
                Mbloodline(Mcnt2%) = ""
            Else
                Mbloodline(Mcnt2%) = TblMaster("bloodline")
            End If
            If IsNull(TblMaster("birdname")) Then
                Mname(Mcnt2%) = ""
            Else
                Mname(Mcnt2%) = TblMaster("birdname")
            End If
            If IsNull(TblMaster("notes")) Then
                Mnotes(Mcnt2%) = ""
            Else
                Mnotes(Mcnt2%) = TblMaster("notes")
            End If

            'Record current records SIRE
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("sireyr")) Then
                Myr(Mcnt%) = Val("")
            Else
                Myr(Mcnt%) = TblMaster("sireyr")
            End If
            If IsNull(TblMaster("siremark")) Then
                Mmark(Mcnt%) = ""
            Else
                Mmark(Mcnt%) = TblMaster("siremark")
            End If
            If IsNull(TblMaster("sireringno")) Then
                Mring(Mcnt%) = ""
            Else
                Mring(Mcnt%) = TblMaster("sireringno")
            End If
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                Fbirdymr(Mcnt3%).Text = Trim$(Right$(Trim(Str$(Myr(Mcnt%))), 2) & " " & Trim(Mmark(Mcnt%)) & " " & Trim(Mring(Mcnt%)) & " " & Trim(Mcolcode(Mcnt%)))
                If Mcnt3% < 14 Then  'only Bird + 14 ancestors Bird(0) to Bird (14)
                    Fbirdbb(Mcnt3%).Text = Trim(Mbredby(Mcnt%))
                End If
            End If
            'Record current record DAM
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("damyr")) Then
                Myr(Mcnt%) = Val("")
            Else
                Myr(Mcnt%) = TblMaster("damyr")
            End If
            If IsNull(TblMaster("dammark")) Then
                Mmark(Mcnt%) = ""
            Else
                Mmark(Mcnt%) = TblMaster("dammark")
            End If
            If IsNull(TblMaster("damringno")) Then
                Mring(Mcnt%) = ""
            Else
                Mring(Mcnt%) = TblMaster("damringno")
            End If
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                Fbirdymr(Mcnt3%).Text = Trim(Mcolcode$(Mcnt%)) & " " & Right$(Trim(Str$(Myr(Mcnt%))), 2) & " " & Trim(Mmark(Mcnt%)) & " " & Trim(Mring(Mcnt%))
            End If
        End If
    Else
        Mcnt% = Mcnt% + 2
        Mcnt3% = Mcnt3% + 2
    End If
Loop

'Display Bred By to Screen
Mcnt2% = 2
For Mcnt% = 0 To 13 Step 1
    Fbirdbb(Mcnt%).Text = Mbloodline(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Ring Nos to Screen
ButtSire.Caption = Trim$(Right$(Trim(Str$(Myr(2))), 2) & " " & Trim(Mmark(2)) & " " & Trim(Mring(2)) & " " & Trim(Mcolcode(2)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(Myr(3))), 2) & " " & Trim(Mmark(3)) & " " & Trim(Mring(3)) & " " & Trim(Mcolcode(3)))

Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    Fbirdymr(Mcnt%).Text = Trim$(Right$(Trim(Str$(Myr(Mcnt2%))), 2) & " " & Trim(Mmark(Mcnt2%)) & " " & Trim(Mring(Mcnt2%)) & " " & Trim(Mcolcode(Mcnt2%)))
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Name to Screen
Mcnt2% = 2
For Mcnt% = 0 To 13 Step 1
    Panel3D2(Mcnt%).Caption = Mname(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Mcnt2% = 2
For Mcnt% = 0 To 5 Step 1
    Panel3D3(Mcnt%).Caption = Mbredby(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Fyr.Text = Trim(Myr(1))
Fmark.Text = Trim(Mmark(1))
Fring.Text = Trim(Mring(1))
Fnotes(0) = Mnotes(2)
Fnotes(1) = Mnotes(3)

'Determine if Sire & Dam record exist, if not remove button
TblMaster.Seek "=", Myr(2), Mmark(2), Mring(2) 'Find the required record.
If TblMaster.NoMatch Then
    ButtSire.Visible = False
End If
TblMaster.Seek "=", Myr(3), Mmark(3), Mring(3) 'Find the required record.
If TblMaster.NoMatch Then
    ButtDam.Visible = False
End If

End Sub

Sub Form_Activate ()
    MOyr = Fyr.Text
    MOmark = Fmark.Text
    MOring = Fring.Text
    Disp_Pedg
End Sub

Sub Form_Load ()
Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
Set TblMaster = DB.OpenTable("master")     'Open Table
TblMaster.Index = "bird_ymr"              'Open Index
End Sub

