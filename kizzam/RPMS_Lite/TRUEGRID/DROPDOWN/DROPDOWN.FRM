VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Data Dropdown"
   ClientHeight    =   4815
   ClientLeft      =   1155
   ClientTop       =   1770
   ClientWidth     =   6780
   Height          =   5505
   Icon            =   DROPDOWN.FRX:0000
   Left            =   1095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6780
   Top             =   1140
   Width           =   6900
   Begin TrueGrid Table2 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      DataSource      =   "Data2"
      Editable        =   0   'False
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      HeadingHeight   =   1
      Height          =   1935
      HorzLines       =   0  'None
      Layout          =   DROPDOWN.FRX:0302
      Left            =   3675
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   1
      Top             =   3000
      UseBookmarks    =   -1  'True
      Visible         =   0   'False
      Width           =   3015
      WrapCellPointer =   0   'False
   End
   Begin Data Data2 
      Caption         =   "Data2"
      Connect         =   ""
      DatabaseName    =   "BIBLIO.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   1875
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Authors"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1755
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "BIBLIO.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Titles"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1755
   End
   Begin TrueGrid Table1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      DataSource      =   "Data1"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      HeadingHeight   =   1
      Height          =   2835
      HorzLines       =   0  'None
      Layout          =   DROPDOWN.FRX:03A8
      Left            =   75
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   90
      UseBookmarks    =   -1  'True
      Width           =   6615
      WrapCellPointer =   0   'False
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Move to the Au_ID column, then click on the cell to pop-up the data dropdown for the column."
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3660
      Width           =   4935
   End
   Begin Menu ExitOption 
      Caption         =   "E&xit!"
   End
   Begin Menu HelpOption 
      Caption         =   "&Help"
      Begin Menu mHelpOption 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin Menu mHelpOption 
         Caption         =   "&Using Help"
         Index           =   2
      End
      Begin Menu mHelpOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu mHelpOption 
         Caption         =   "&About Data Dropdown..."
         Index           =   4
      End
   End
End

Sub CenterForm (F As Form)

' Center the specified form within the screen

    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub ExitOption_Click ()
    
    Unload Form1

End Sub

Sub Form_Load ()
    Dim dbdir As String

    'Set the DatabaseName of the Data controls
    dbdir = App.Path
    Data1.DatabaseName = dbdir + "\BIBLIO.MDB"
    Data2.DatabaseName = dbdir + "\BIBLIO.MDB"

    'Center the Form on the screen
    CenterForm Form1

End Sub

Sub mHelpOption_Click (index As Integer)

    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for Dropdown sample
    Select Case index
        Case 1
            HelpContext Form1, HELP_DROPDOWN
        Case 2
            HelpOnHelp Form1
        Case 4
            About.Show 1
    End Select

End Sub

Sub Table1_RequestEdit (Cancel As Integer, KeyAscii As Integer)
    ' If they attempt to type in the cell, just let it happen
    ' and don't pop up a dropdown.
    If KeyAscii Then Exit Sub

    ' Popup only for Au_ID
    If Table1.ColumnName(Table1.ColumnIndex) <> "Au_ID" Then Exit Sub

    ' Position the related data control to the proper row
    Data2.Recordset.FindFirst "Au_ID = " & Table1.Text

    ' Popup the dropdown
    Table2.Top = Table1.CellRectTop + Table1.CellRectHeight
    Table2.Left = Table1.CellRectLeft
    Table2.Visible = True
    Table2.SetFocus

    ' Assure we don't go into edit mode
    Cancel = True
End Sub

Sub Table2_DblClick ()
    Table1.ColumnText(Table1.ColumnIndex) = Table2.ColumnText(1)
    Table2.Visible = False
End Sub

Sub Table2_LostFocus ()
    Table2.Visible = False
End Sub

