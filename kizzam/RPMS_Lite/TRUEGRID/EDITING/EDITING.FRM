VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Editing and Display Tips"
   ClientHeight    =   4215
   ClientLeft      =   1095
   ClientTop       =   1770
   ClientWidth     =   7560
   Height          =   4905
   Icon            =   EDITING.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7560
   Top             =   1140
   Width           =   7680
   Begin CommandButton Command1 
      Caption         =   "Reprioritize"
      Height          =   375
      Left            =   2790
      TabIndex        =   1
      Top             =   3660
      Width           =   1875
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
      HeadingHeight   =   1.25
      Height          =   3255
      HorzLines       =   2  '3D
      Layout          =   EDITING.FRX:0302
      Left            =   210
      LinesPerRow     =   2
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   0   'False
      SplitTabMode    =   1  'Tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   210
      UseBookmarks    =   -1  'True
      Width           =   7035
      WrapCellPointer =   0   'False
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "TASKS.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "select * from Tasklist order by Priority"
      Top             =   3660
      Visible         =   0   'False
      Width           =   1755
   End
   Begin Menu FileOption 
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
         Caption         =   "&About Editing and Display Tips..."
         Index           =   4
      End
   End
End

Sub CenterForm (F As Form)

' Center the specified form within the screen

    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub Command1_Click ()
    ' Since RecordSource indicates that we order by priority,
    ' a simple refresh redisplays everything in the right
    ' priority order.

    Data1.Refresh
End Sub

Sub FileOption_Click ()

    Unload Form1

End Sub

Sub Form_Load ()
    Dim dbdir As String
    
    ' Set the databasename property of the data control
    dbdir = app.Path
    Data1.DatabaseName = dbdir + "\TASKS.MDB"

    ' Center Form on screen
    CenterForm Form1

    ' Assure that the description column in the right split is
    ' set so attributes will be queried as needed

    Table1.SplitIndex = 2
    Table1.ColumnCellAttrs(4) = True

    ' Set up value list combo box for the status column
    Table1.SplitIndex = 1
    Table1.VlistColumn = 2
    Table1.ColumnButton(2) = True
    Table1.VlistStyle = GVLS_COMBO Or GVLS_SORTCOMBO
    Table1.VlistData(1) = "Cancel"
    Table1.VlistData(2) = "Reschedule"
    Table1.VlistData(3) = "Satisfactory"
    Table1.VlistData(4) = "Unacceptable"
    Table1.VlistData(5) = "In Progress"
    Table1.VlistData(6) = "Not Started"
    Table1.VlistData(7) = "Pending"

    ' Set up up value list of text and pictures for Past Due Column
    Table1.SplitIndex = 2
    Table1.VlistColumn = 1
    Table1.VlistStyle = GVLS_TEXT Or GVLS_TRANSLATE
    Table1.VlistData(1) = " 0"
    Table1.VlistData(2) = "-1"
    Table1.VlistText(1) = "OK"
    Table1.VlistPicture(2) = LoadPicture(app.Path + "\warn.bmp")


End Sub

Sub mHelpOption_Click (index As Integer)
   
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for the Editing sample
    Select Case index
        Case 1
            HelpContext Form1, HELP_EDITING
        Case 2
            HelpOnHelp Form1
        Case 4
            About.Show 1
    End Select

End Sub

Sub Table1_Append ()
    ' Add a new record if they arrow down past the bottom

    If MsgBox("Add new record?", 1) = 2 Then Exit Sub
    Data1.Recordset.AddNew
    Data1.Recordset.Update
    Data1.Recordset.MoveLast
End Sub

Sub Table1_FetchAttributes (Status As Integer, Split As Integer, Row As Long, Col As Integer, FgColor As Long, BgColor As Long, FontStyle As Integer)
    
    ' Show past due items as white on red, unless the data
    ' has been updated, in which case we leave it alone.
    
    If Table1.ColumnText(1) = "-1" Then
        If Status And GFS_CHANGED Then Exit Sub
        FgColor = &HFFFFFF  ' white
        BgColor = &H80      ' red
    End If

End Sub

