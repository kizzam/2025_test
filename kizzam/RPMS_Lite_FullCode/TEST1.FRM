VERSION 2.00
Begin Form Test 
   Caption         =   "Test"
   ClientHeight    =   4020
   ClientLeft      =   1680
   ClientTop       =   1590
   ClientWidth     =   7365
   Height          =   4425
   Left            =   1620
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1245
   Width           =   7485
   Begin ListBox List2 
      Height          =   615
      Left            =   4560
      TabIndex        =   13
      Top             =   840
      Width           =   2535
   End
   Begin ListBox List1 
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   2295
   End
   Begin TextBox Text5 
      Height          =   285
      Left            =   5880
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   3000
      Width           =   855
   End
   Begin TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   3000
      Width           =   975
   End
   Begin SSCommand CmdOK 
      Caption         =   "&OK"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin SSCommand CmdQuit 
      Caption         =   "&Quit"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin SSCommand CmdDel 
      Caption         =   "<<"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin SSCommand CmdAdd 
      Caption         =   ">>"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin PictureBox CSVList2 
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1065
      ScaleWidth      =   2505
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin PictureBox CSVList1 
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1065
      ScaleWidth      =   2265
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Text            =   "0600"
      Top             =   240
      Width           =   1215
   End
   Begin TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "DDRPA"
      Top             =   240
      Width           =   1215
   End
   Begin TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "1996"
      Top             =   240
      Width           =   1095
   End
   Begin Data Data2 
      Caption         =   "Data2"
      Connect         =   ""
      DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "teambird"
      Top             =   3720
      Width           =   2055
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "teams"
      Top             =   3720
      Width           =   1935
   End
   Begin Label Label1 
      Caption         =   "Y/M/R"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End

Sub CmdOK_Click ()
Dim teamsin As String
Dim MyDB As Database, MySet As Snapshot
Dim Entry, I, Msg   ' Declare variables
Dim Cnt1 As Integer, Cnt2 As Integer

teamsin$ = "SELECT * FROM teambird WHERE yr = 1996 AND mark = '" & Text2.Text & "' AND ringno = '" & Text3.Text & "'"
Data2.RecordSource = teamsin$
Data2.Refresh
'CSVList2.ListDataSource = teamsin$
Text4.Text = CSVList2


Set MyDB = OpenDatabase("h:\dev_vb\rp\RP.MDB")
Set MySet = MyDB.CreateSnapshot(teamsin$)
'MySet.Sort = "Publisher's Name"
Set MySet = MySet.CreateSnapshot()
MySet.MoveLast
Cnt% = MySet.RecordCount
Text5.Text = Str(Cnt%)

'Fill List Box
MySet.MoveFirst
For I = 1 To Cnt%    ' Count from 1 to 100.
    Data1.Recordset.FindFirst "code = '" & MySet("code") & "'"
    'Validate lookup code found later
    Entry = Data1.Recordset("description")
    List1.AddItem Entry ' Add the entry.
    MySet.MoveNext
Next I
    'Msg = "Choose OK to remove every other entry."
    'MsgBox Msg  ' Display message.
    'For I = 1 To 50 ' Determine how to
    '    List1.RemoveItem I  ' remove every other
    'Next I  ' item.
    'Msg = "Choose OK to remove all items from the list box."
    'MsgBox Msg  ' Display message.
    'List1.Clear ' Clear list box.


End Sub

Sub CmdQuit_Click ()
Unload Me

End Sub

Sub CSVList2_QueryItem (Item As Long, Text As String)
Text4.Text = CSVList2

End Sub

