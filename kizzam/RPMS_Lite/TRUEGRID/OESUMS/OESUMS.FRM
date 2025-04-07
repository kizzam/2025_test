VERSION 2.00
Begin Form OESumFrm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Column Summary Sample"
   ClientHeight    =   3345
   ClientLeft      =   1650
   ClientTop       =   1860
   ClientWidth     =   6555
   ForeColor       =   &H00C0C0C0&
   Height          =   4035
   Icon            =   OESUMS.FRX:0000
   Left            =   1590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6555
   Top             =   1230
   Width           =   6675
   Begin CommandButton Command1 
      Caption         =   "&New Order"
      Height          =   435
      Left            =   180
      TabIndex        =   11
      Top             =   2760
      Width           =   1155
   End
   Begin TextBox TaxPercent 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Text            =   "6"
      Top             =   2400
      Width           =   375
   End
   Begin TextBox PieceCount 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   3075
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ";"
      DatabaseName    =   "OESUMS.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "LineItems"
      Top             =   3600
      Width           =   1155
   End
   Begin TrueGrid Table1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data1"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeadingHeight   =   1
      Height          =   1935
      HorzLines       =   2  '3D
      Layout          =   OESUMS.FRX:0302
      LayoutIndex     =   1
      Left            =   180
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   0   'False
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   120
      UseBookmarks    =   -1  'True
      Width           =   6195
      WrapCellPointer =   0   'False
   End
   Begin Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Order Total:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   3660
      TabIndex        =   10
      Top             =   2970
      Width           =   975
   End
   Begin Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tax Amount:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   3540
      TabIndex        =   9
      Top             =   2700
      Width           =   1095
   End
   Begin Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tax Percent:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2430
      Width           =   1155
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sub-total:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   795
   End
   Begin Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   2940
      Width           =   1455
   End
   Begin Label TaxAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   2700
      Width           =   1455
   End
   Begin Label SubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      Height          =   195
      Left            =   6180
      TabIndex        =   3
      Top             =   2460
      Width           =   255
   End
   Begin Menu ExitMenuOption 
      Caption         =   "E&xit"
   End
   Begin Menu HelpMenuOption 
      Caption         =   "&Help"
      Begin Menu HelpOption 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin Menu HelpOption 
         Caption         =   "&Using Help"
         Index           =   2
      End
      Begin Menu HelpOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu HelpOption 
         Caption         =   "&About Order Entry Summaries..."
         Index           =   4
      End
   End
End
Option Explicit

Sub CenterForm (F As Form)

' Center the specified form within the screen

    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub Command1_Click ()
    
    ' Add new record
    Data1.Recordset.AddNew
End Sub

Sub ExitMenuOption_Click ()

    'Unload forms and exit

    Unload About
    Unload OESumFrm
    End

End Sub

Sub Form_Load ()
    
    ' Centers form on screen
    CenterForm OESumFrm
    
    ' Change Drive and Dir to current path
    ChDir App.Path
    ChDrive App.Path

    ' Enable Column Summary for Qty and Total Columns
    Table1.ColumnSumEnable(5) = True
    Table1.ColumnSumEnable(1) = True

End Sub

Sub HelpOption_Click (index As Integer)
    
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for Column Summary Sample
    Select Case index
    Case 1
	HelpContext OESumFrm, HELP_OESUM
    Case 2
	HelpOnHelp OESumFrm
    Case 4
	About.Show 1
    End Select


End Sub

Sub RecalcTotal ()

    Dim tot As Double
    Dim taxtot As Double
    
    ' calculate order total
    tot = Table1.ColumnSum(5)
    SubTotal = Format$(tot, "Currency")
    taxtot = tot * Val(TaxPercent) / 100
    TaxAmount = Format$(taxtot, "Currency")
    Total = Format$(taxtot + tot, "Currency")
End Sub

Sub Table1_Append ()
    
    ' Add new record
    Data1.Recordset.AddNew
End Sub

Sub Table1_ColumnSumChange (Col As Integer)
    
    '  If the called for the Total column call recalculate function
    If Col = 5 Then
	RecalcTotal
    ElseIf Col = 1 Then
	' If called for the qty column change text box that displays totals
	PieceCount = Table1.ColumnSum(1) & " pieces in order"
    End If
End Sub

Sub TaxPercent_Change ()
    
    ' If users changes tax rate call Recalctotal
    RecalcTotal
End Sub

