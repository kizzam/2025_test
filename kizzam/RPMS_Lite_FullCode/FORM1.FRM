VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   1950
   ClientTop       =   2040
   ClientWidth     =   7365
   Height          =   4425
   Left            =   1890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1695
   Width           =   7485
   Begin TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   3120
      Width           =   1455
   End
   Begin TextBox Text5 
      DataField       =   "dob_li"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   480
      Width           =   1455
   End
   Begin TextBox Text4 
      DataField       =   "dob"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   480
      Width           =   1335
   End
   Begin TextBox Text3 
      DataField       =   "ringno"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   120
      Width           =   1575
   End
   Begin TextBox Text2 
      DataField       =   "mark"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Width           =   1095
   End
   Begin TextBox Text1 
      DataField       =   "yr"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "F:\DEV_VB\RP\RP.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "master"
      Top             =   3600
      Width           =   1695
   End
   Begin CommandButton Command1 
      Caption         =   "Convert Date (DOB) to LI"
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
End
Option Explicit

Sub Command1_Click ()
Dim cdob As String
Dim a As Integer
Dim b As Integer
Dim c As String

Data1.Recordset.MoveFirst
a = 0

Do Until Data1.Recordset.EOF    ' Until end of file.
    a% = a% + 1
    Data1.Recordset.Edit
    If Not IsDate(Data1.Recordset("dob")) Then
	cdob$ = InputBox("Enter date (dd/mm/yy) :")  ' Get date
	Data1.Recordset("dob") = Trim(cdob$)
    End If
    
    'Pad out date to dd/mm/yy
    'c$ = Data1.Recordset("dob")
    'b% = InStr(c, "/")
    'If b% < 3 Then
    '    c$ = "0" & c$
    'End If
    'Data1.Recordset("dob_li") = DateValue(c$)
    'Data1.Recordset("dob") = Trim(c$)
    'Convert new field ddob to correct value
    Data1.Recordset("ddob") = DateValue(Data1.Recordset("dob"))
    text6.Text = Data1.Recordset("ddob")
    
    'text6.Text = Data1.Recordset("dob_li")
    Data1.Recordset.Update
    Data1.Recordset.MoveNext    ' Move to next record.
Loop
text6.Text = a%
End Sub

