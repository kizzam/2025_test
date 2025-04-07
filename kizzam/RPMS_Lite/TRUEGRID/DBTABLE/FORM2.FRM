VERSION 2.00
Begin Form Form2 
   Caption         =   "TrueGrid"
   ClientHeight    =   1755
   ClientLeft      =   1830
   ClientTop       =   2310
   ClientWidth     =   4800
   Height          =   2160
   Left            =   1770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4800
   Top             =   1965
   Width           =   4920
   Begin CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   885
   End
   Begin PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   FORM2.FRX:0000
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   150
      Width           =   645
   End
   Begin Label Label2 
      Caption         =   "Copyright (C) 1993  Apex Software Corporation"
      Height          =   255
      Left            =   390
      TabIndex        =   3
      Top             =   780
      Width           =   3975
   End
   Begin Label Label1 
      Caption         =   "Apex TrueGrid dBTable Program"
      Height          =   255
      Left            =   930
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
End

Sub Command1_Click ()
    Unload Form2
End Sub

Sub Form_Load ()

    Form2.Caption = "TrueGrid Version " + Form1.FieldTable.Version

End Sub

