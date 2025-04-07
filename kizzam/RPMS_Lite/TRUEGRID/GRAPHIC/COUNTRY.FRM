VERSION 2.00
Begin Form CountryFrm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "In-Cell Graphic Sample"
   ClientHeight    =   3900
   ClientLeft      =   1695
   ClientTop       =   2235
   ClientWidth     =   8160
   Height          =   4590
   Icon            =   COUNTRY.FRX:0000
   Left            =   1635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   8160
   Top             =   1605
   Width           =   8280
   Begin ListBox List1 
      Height          =   1005
      Left            =   240
      TabIndex        =   1
      Top             =   3870
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ";"
      DatabaseName    =   "NATIONS.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Sales"
      Top             =   3900
      Visible         =   0   'False
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
      Height          =   3525
      HorzLines       =   2  '3D
      Layout          =   COUNTRY.FRX:0302
      Left            =   210
      LinesPerRow     =   2
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   180
      UseBookmarks    =   -1  'True
      Width           =   7725
      WrapCellPointer =   0   'False
   End
   Begin Menu ExitMenuOption 
      Caption         =   "E&xit!"
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
         Caption         =   "&About In-Cell Graphic Sample"
         Index           =   4
      End
   End
End
Option Explicit

Dim Status As Integer

Sub AddCountry (txt As String)
    
    ' Add country to list box
    List1.AddItem txt
End Sub

Sub CenterForm (F As Form)

    ' Center the specified form within the screen
    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub Exit_Click ()
    
    ' Unload forms
    Unload CountryFrm
    Unload About

    End

End Sub

Sub ExitMenuOption_Click ()
    'Unload forms and exit

    Unload About
    Unload CountryFrm
    End

End Sub

Sub Form_Load ()
    Dim dbdir As String

    ' Center form on screen
    CenterForm CountryFrm

    ' Open database in current directory
    dbdir = App.Path
    Data1.DatabaseName = dbdir + "\nations.mdb"
    
    ' Display column button for flags column
    Table1.ColumnButton(1) = True

    ' Populate combo box
    AddCountry "Australia"
    AddCountry "Austria"
    AddCountry "Brazil"
    AddCountry "Canada"
    AddCountry "Denmark"
    AddCountry "Finland"
    AddCountry "France"
    AddCountry "Germany"
    AddCountry "Ireland"
    AddCountry "Italy"
    AddCountry "Japan"
    AddCountry "Mexico"
    AddCountry "Netherlands"
    AddCountry "Norway"
    AddCountry "New Zealand"
    AddCountry "Portugal"
    AddCountry "Spain"
    AddCountry "Sweden"
    AddCountry "Switzerland"
    AddCountry "Turkey"
    AddCountry "United Kingdom"
    AddCountry "United States"

End Sub

Sub HelpOption_Click (index As Integer)
    
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for True Browser
    Select Case index
    Case 1
	HelpContext CountryFrm, HELP_GRAPHIC
    Case 2
	HelpOnHelp CountryFrm
    Case 4
	About.Show 1
    End Select

End Sub

Sub List1_DblClick ()
    
    ' If the user double clicks a selection enter it in the database
    ' and make the list box invisible
    Table1.ColumnText(TgFindColumn(Table1, "Country")) = List1.List(List1.ListIndex)
    List1.Visible = False

End Sub

Sub List1_LostFocus ()
    
    ' If the user clicks elsewhere on the form
    ' then hide the list box
    List1.Visible = False

    ' Toggle Status variable
    Status = Status Xor True
End Sub

Sub Table1_Append ()
    
    ' Add a new record
    Data1.Recordset.AddNew
End Sub

Sub Table1_ClickButton ()


    
    ' If the user clicks on the flag column button then bring up list box
    If Table1.ColumnName(Table1.ColumnIndex) = "Country" Then
	Status = Status Xor True    ' Toggle between true and false
	If Status Then
	    List1.Left = Table1.CellRectLeft
	    List1.Top = Table1.CellRectTop + Table1.CellRectHeight
	    List1.Visible = True
	    List1.SetFocus
	Else
	    List1.Visible = False
	    Status = Status Xor True
	End If
    End If

End Sub

Sub Table1_KeyDown (KeyCode As Integer, Shift As Integer)
    
    ' The user wants to delete a row when not in edit mode
    If KeyCode = KEY_DELETE And Not Table1.EditActive Then Data1.Recordset.Delete
End Sub

