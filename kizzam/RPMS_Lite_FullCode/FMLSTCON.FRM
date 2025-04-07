VERSION 2.00
Begin Form FMlstcontact 
   Caption         =   "Contact List"
   ClientHeight    =   6030
   ClientLeft      =   225
   ClientTop       =   675
   ClientWidth     =   9105
   Height          =   6435
   Left            =   165
   LinkTopic       =   "Form4"
   ScaleHeight     =   6030
   ScaleWidth      =   9105
   Top             =   330
   Width           =   9225
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   9135
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin CrystalReport Report2 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Screen Selections - Contacts"
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data1"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   7920
         MarginBottom    =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginTop       =   0
         PrinterCollation=   2  'default
         PrinterCopies   =   1
         PrinterDriver   =   ""
         PrinterName     =   ""
         PrinterPort     =   ""
         PrinterStartPage=   0
         PrinterStopPage =   0
         PrintFileCharSepQuote=   ""
         PrintFileCharSepSeparator=   ""
         PrintFileName   =   ""
         PrintFileType   =   2  'Text
         PrintFileUseRptDateFmt=   0   'False
         PrintFileUseRptNumberFmt=   0   'False
         ReportFileName  =   ""
         ReportSource    =   1  'Bound TrueGrid Control
         ReportTitle     =   "CONTACTS"
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   4560
         UserName        =   ""
         WindowBorderStyle=   2  'Sizable
         WindowControlBox=   -1  'True
         WindowControls  =   -1  'True
         WindowHeight    =   300
         WindowLeft      =   100
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowParentHandle=   0
         WindowState     =   0  'normal
         WindowTitle     =   ""
         WindowTop       =   100
         WindowWidth     =   480
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeadingHeight   =   1
         Height          =   2895
         HorzLines       =   0  'None
         Layout          =   FMLSTCON.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   18
         Top             =   1560
         UseBookmarks    =   -1  'True
         Width           =   8655
         WrapCellPointer =   0   'False
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   0   'False
         BoundReportHeading=   ""
         Connect         =   ""
         CopiesToPrinter =   1
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   8160
         MarginBottom    =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginTop       =   0
         PrinterCollation=   2  'default
         PrinterCopies   =   1
         PrinterDriver   =   ""
         PrinterName     =   ""
         PrinterPort     =   ""
         PrinterStartPage=   0
         PrinterStopPage =   0
         PrintFileCharSepQuote=   ""
         PrintFileCharSepSeparator=   ""
         PrintFileName   =   ""
         PrintFileType   =   2  'Text
         PrintFileUseRptDateFmt=   0   'False
         PrintFileUseRptNumberFmt=   0   'False
         ReportFileName  =   ""
         ReportSource    =   0  'Report File
         ReportTitle     =   ""
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   4560
         UserName        =   ""
         WindowBorderStyle=   2  'Sizable
         WindowControlBox=   -1  'True
         WindowControls  =   -1  'True
         WindowHeight    =   300
         WindowLeft      =   100
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowParentHandle=   0
         WindowState     =   0  'normal
         WindowTitle     =   ""
         WindowTop       =   100
         WindowWidth     =   480
      End
      Begin CommonDialog CMDialog1 
         Left            =   7320
         Top             =   4680
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Order by"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   11
         Top             =   120
         Width           =   8055
         Begin SSOption byCategory 
            Caption         =   "by Category"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6600
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin SSOption byFederation 
            Caption         =   "by Federation"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5040
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin SSOption byClub 
            Caption         =   "by Club"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3960
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin SSOption byFirstName 
            Caption         =   "by First Name"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2400
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin SSOption byLastName 
            Caption         =   "by Last Name"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Display Only"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   8
         Top             =   720
         Width           =   8055
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMLSTCON.FRX:0411
            Contents        =   FMLSTCON.FRX:0451
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   315
            Left            =   3960
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   10
            Text            =   "ALL"
            TextDisplay     =   0  'first visible
            Top             =   240
            ValueCol        =   -1
            Width           =   975
         End
         Begin SSPanel Panel3D4 
            BevelOuter      =   0  'None
            Caption         =   "Category ="
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2760
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin SSPanel Panel3D3 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   720
         TabIndex        =   3
         Top             =   4680
         Width           =   7215
         Begin SSCommand cmdPrintScrn 
            AutoSize        =   1  'Adjust Picture Size To Button
            Caption         =   "Print &Screen"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   4800
            RoundedCorners  =   0   'False
            TabIndex        =   19
            Top             =   0
            Width           =   1215
         End
         Begin SSCommand Butt3Prt 
            AutoSize        =   1  'Adjust Picture Size To Button
            Caption         =   "&Print"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   3600
            RoundedCorners  =   0   'False
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin SSCommand Butt3Find 
            Caption         =   "&Find"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   2400
            TabIndex        =   7
            Top             =   0
            Width           =   1215
         End
         Begin SSCommand Butt3Cancel 
            Caption         =   "E&xit"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   6000
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
         Begin SSCommand Butt3Process 
            Caption         =   "&Process"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   1200
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin SSCommand Butt3Add 
            AutoSize        =   1  'Adjust Picture Size To Button
            Caption         =   "&Add"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   0
            RoundedCorners  =   0   'False
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
      End
      Begin TextBox RefreshData 
         Height          =   285
         Left            =   7800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7440
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "contacts"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1140
      End
   End
End
Option Explicit

Sub Butt3Add_Click ()
    FMcontact.Task.Text = "A"
    FMcontact.Show 1
    Data1.Refresh '29/8
End Sub

Sub Butt3Cancel_Click ()
Unload fmlstcontact
End Sub

Sub Butt3Find_Click ()
  Dim Key As String
  Dim Mnamel As Integer
  Dim Mnamef As String
  Dim Mcnt As Integer
  Dim Ma As Integer
  Dim Mkey1 As String

    'bit of work to do here, composite key Yr/Mark/RingNo.
    Key$ = InputBox("Enter Last Name to find :", "Find Contact Details")
    Key$ = UCase(Key$)
    If Trim(Key$) <> "" Then
        Data1.Recordset.FindFirst "namel = '" & Key$ & "'"
        If Data1.Recordset.NoMatch Then  ' Check if record is found.
            MsgBox "Contact Record Not in currently selected 'show only' list", 0, "Not Found"
        Else
            MsgBox "Contact Record on File", 0, "FOUND"
        End If
    End If
End Sub

Sub Butt3Process_Click ()

    Dim ii As Integer

    'if no record selected display message
    If Table1.BookmarkCount < 1 Then
        MsgBox "No Contact(s) selected", 0, "ERROR - Unable to Process"
    Else
        ' Process each selected row.
        For ii% = 0 To Table1.BookmarkCount - 1
            DoEvents
            Data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
            FMcontact.Task.Text = "M"
            FMcontact.Keynamel.Text = Data1.Recordset("namel")
            FMcontact.Keynamef.Text = Data1.Recordset("namef")
            FMcontact.Show 1
        Next ii%
    End If
    Table1.BookmarkCount = 0
    
    DoEvents
    
    'refresh data1 if records deleted
    If RefreshData.Text = "Y" Then
        Data1.Refresh
        DoEvents
        RefreshData.Text = "N"
    End If

End Sub

Sub Butt3Prt_Click ()
    Dim Formula As String
    Dim cType As String
    cType$ = "R"

    Report1.ReportFileName = DBPATH$ & "CON50916.rpt"
    
    '???OrderBy     'Set Order of records in Report
    Formula$ = ShowWhat("R") 'set orderby and selection of records
    Report1.SelectionFormula = Formula$
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr1
        CMDialog1.Copies = 1
        CMDialog1.FromPage = 1
        CMDialog1.ToPage = 1
        CMDialog1.CancelError = True
        CMDialog1.Action = 5
        DoEvents
        On Error GoTo 0
    End If

    If CMDialog1.PrinterDefault Then
        Report1.CopiesToPrinter = CMDialog1.Copies
    Else
        MsgBox "The printer you use MUST be selected as the default printer.", 16, "Printer Error"
        Exit Sub
    End If
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr1:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub


End Sub

Sub byCategory_Click (Value As Integer)
setselections
End Sub

Sub byClub_Click (Value As Integer)
setselections

End Sub

Sub byFederation_Click (Value As Integer)
setselections

End Sub

Sub byFirstName_Click (Value As Integer)
setselections

End Sub

Sub byLastName_Click (Value As Integer)
setselections
End Sub

Sub cmdPrintScrn_Click ()
Report2.Action = 1
DoEvents
End Sub

Sub CSCombo1_change ()
setselections
End Sub

Sub Form_Load ()
screen.MousePointer = 11
Data1.DatabaseName = DBFullFileName$
CSCombo1.Text = "ALL"
byLastName.Value = True
RefreshData.Text = "N" 'Default to no refresh of Data
setselections
screen.MousePointer = 0
End Sub

Function OrderBy () As String
    If byLastName Then
       OrderBy$ = "namel"
    ElseIf ByFirstName Then
        OrderBy$ = "namef, namel"
    ElseIf ByClub Then
        OrderBy$ = "club, namel"
    ElseIf ByFederation Then
        OrderBy$ = "federation, club"
    ElseIf ByCategory Then
        OrderBy$ = "category, namel"
    Else
       OrderBy$ = "namel"
    End If
End Function

Sub setselections ()
    Dim s As String
    Dim o As String
    Dim cType As String
    cType$ = ""

    On Error GoTo ByErr9

    s$ = ShowWhat(cType$)
    o$ = OrderBy()
    
    If s$ = "" Then
        Data1.RecordSource = "SELECT * FROM contacts ORDER BY " & o$
    Else
        Data1.RecordSource = "SELECT * FROM contacts WHERE " & s$ & " ORDER BY " & o$
    End If
    Data1.Refresh
    Exit Sub
    
ByErr9:
    If Err = 380 Then
        Resume Next
    End If
    MsgBox "Error :" & Err & " " & Error$, 16, "Error"
    Exit Sub
End Sub

Function ShowWhat (cType As String) As String

    Dim s As String

    'If ShowCurrent Then
    '   s$ = "indsoh = True"
    'End If

    s$ = CSCombo1.Text
    
    If cType$ = "R" Then 'Crystal Report Selection
        If Trim(s$) = "ALL" Then
            ShowWhat$ = ""
        Else
            ShowWhat$ = "{contacts.category} = '" & Trim(s$) & "'"
        End If
    Else
        If Trim(s$) = "ALL" Then
            ShowWhat$ = ""
        Else
            ShowWhat$ = "category = '" & Trim(s$) & "'"
        End If
    End If

End Function

