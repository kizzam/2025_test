VERSION 2.00
Begin Form FMcontact 
   Caption         =   "Contact Manager"
   ClientHeight    =   6735
   ClientLeft      =   570
   ClientTop       =   345
   ClientWidth     =   8775
   Height          =   7140
   Left            =   510
   LinkTopic       =   "Form3"
   ScaleHeight     =   6735
   ScaleWidth      =   8775
   Top             =   0
   Width           =   8895
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6360
      Width           =   8775
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin SSPanel Panel3D17 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Category:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Top             =   5040
         Width           =   855
      End
      Begin TextBox category 
         Height          =   285
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   5040
         Width           =   495
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMCONTAC.FRX:0000
         Contents        =   FMCONTAC.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   4800
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         TabIndex        =   6
         Text            =   "CSCombo1"
         TextDisplay     =   0  'first visible
         Top             =   2280
         ValueCol        =   -1
         Width           =   975
      End
      Begin SSPanel Panel3D16 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         Caption         =   "State:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   36
         Top             =   2280
         Width           =   615
      End
      Begin TextBox Keynamef 
         Height          =   285
         Left            =   7560
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Text14"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin TextBox Keynamel 
         Height          =   285
         Left            =   7560
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Text14"
         Top             =   5400
         Visible         =   0   'False
         Width           =   975
      End
      Begin TextBox Task 
         Height          =   285
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Text14"
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin SSPanel Panel3D15 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2640
         TabIndex        =   32
         Top             =   5640
         Width           =   3255
         Begin SSCommand ButtOK 
            Caption         =   "&Save"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2160
            TabIndex        =   18
            Top             =   120
            Width           =   855
         End
         Begin SSCommand ButtCancel 
            Caption         =   "E&xit"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1200
            TabIndex        =   17
            Top             =   120
            Width           =   735
         End
         Begin SSCommand ButtDelete 
            Caption         =   "&Delete"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   120
            Width           =   735
         End
      End
      Begin SSPanel Panel3D14 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Notes:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   3960
         Width           =   615
      End
      Begin TextBox notes 
         Height          =   975
         Left            =   2040
         TabIndex        =   14
         Text            =   "Text13"
         Top             =   3960
         Width           =   5295
      End
      Begin SSPanel Panel3D13 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Federation:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin SSPanel Panel3D12 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Club:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   3480
         Width           =   1095
      End
      Begin TextBox federation 
         Height          =   285
         Left            =   5160
         MaxLength       =   12
         TabIndex        =   13
         Text            =   "Text12"
         Top             =   3480
         Width           =   1575
      End
      Begin TextBox club 
         Height          =   285
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "Text11"
         Top             =   3480
         Width           =   1695
      End
      Begin TextBox faxwk 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   11
         Text            =   "Text10"
         Top             =   3120
         Width           =   1575
      End
      Begin TextBox fax 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   9
         Text            =   "Text9"
         Top             =   2760
         Width           =   1575
      End
      Begin TextBox phonewk 
         Height          =   285
         Left            =   2040
         MaxLength       =   16
         TabIndex        =   10
         Text            =   "Text8"
         Top             =   3120
         Width           =   1815
      End
      Begin TextBox phone 
         Height          =   285
         Left            =   2040
         MaxLength       =   16
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   2760
         Width           =   1815
      End
      Begin TextBox postcode 
         Height          =   285
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   2280
         Width           =   975
      End
      Begin TextBox city 
         Height          =   285
         Left            =   2040
         MaxLength       =   18
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   2280
         Width           =   1935
      End
      Begin TextBox address2 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1920
         Width           =   2775
      End
      Begin TextBox address1 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1560
         Width           =   2775
      End
      Begin TextBox namel 
         Height          =   285
         Left            =   2040
         MaxLength       =   18
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   960
         Width           =   1815
      End
      Begin TextBox namef 
         Height          =   285
         Left            =   2040
         MaxLength       =   18
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   1815
      End
      Begin SSPanel Panel3D11 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Work Fax:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   3120
         Width           =   975
      End
      Begin SSPanel Panel3D10 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Home Fax:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   2760
         Width           =   975
      End
      Begin SSPanel Panel3D9 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Work Phone:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   3120
         Width           =   1335
      End
      Begin SSPanel Panel3D8 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Home Phone:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2760
         Width           =   1335
      End
      Begin SSPanel Panel3D7 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Postcode:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5880
         TabIndex        =   24
         Top             =   2280
         Width           =   975
      End
      Begin SSPanel Panel3D6 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "City:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   2280
         Width           =   495
      End
      Begin SSPanel Panel3D5 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin SSPanel Panel3D4 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Last Name:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin SSPanel Panel3D3 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "First Name:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
   End
End
Option Explicit
Dim TblMaster As Table
Dim DB As Database
Dim GLindchange As Integer

Sub address1_Change ()
    GLindchange% = 1
End Sub

Sub address2_Change ()
    GLindchange% = 1
End Sub

Sub ButtCancel_Click ()
    Dim reply As Integer, zz As Integer

    If FMcontact.Task.Text = "A" Then
        'If RecordChangedInAdd() Then
        If GLindchange% Then ' has any fields changed
            reply% = MsgBox("Do you want to discard new record?", 4, "Warning?")
            If reply% <> 6 Then
                namef.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Add Record Cancelled", 16, "Notification Only"
        End If
    End If
    
    If FMcontact.Task.Text = "M" Then
        'If RecordChanged() Then
        If GLindchange% Then ' has any fields changed
            reply = MsgBox("Do you want to cancel changes?", 4, "Warning?")
            If reply <> 6 Then
                address1.SetFocus
                Exit Sub
            End If
        End If
    End If

    FMcontact.Hide

End Sub

Sub ButtDelete_Click ()
    Dim reply As Integer
    reply% = MsgBox("Confirm Deletion of record ", 4, "WARNING!")
    If reply% = 6 Then
        TblMaster.Delete
        MsgBox "Record Deleted", 16, "Notification Only"
        FMlstcontact.RefreshData.Text = "Y"
        FMcontact.Hide
    Else
        address1.SetFocus
    End If
End Sub

Sub ButtOK_Click ()
    On Error GoTo BokErr

    'Validate Data entry
    'Validate key first
    '   A = Add new record
    '   C = ?
    '   M = Maintain record

    If Task.Text = "A" Or Task.Text = "C" Then
        If Trim(namef.Text = "") Then
            MsgBox "First Name cannot be empty", 16, "Error: Cannot update record"
            namef.SetFocus
            Exit Sub
        End If
        'etc add in all validation for Add, maybe set up in Add Val func
        If Trim(namel.Text = "") Then
            MsgBox "Last Name cannot be empty", 16, "Error: Cannot update record"
            namel.SetFocus
            Exit Sub
        End If
    End If
    If Task.Text = "M" Or Task.Text = "A" Or Task.Text = "C" Then
        If Trim(CSCombo1.Value = "") Then
            MsgBox "State MUST NOT be empty", 16, "Error: Cannot update record"
            CSCombo1.SetFocus
            Exit Sub
        End If
        'etc add in all validation for ADD & Mod maybe set up in func
    End If

    'Add record
    If Task.Text = "A" Or Task.Text = "C" Then
        TblMaster.Seek "=", namel, namef 'Ensure unique index
        If TblMaster.NoMatch Then
            TblMaster.AddNew
            ScrToTbl
            MsgBox "Record Added", 16, "Notification Only"
        Else
            MsgBox "Cannot add record, Name already on file", 16, "Duplicate Key"
            namef.SetFocus
            namef.SelStart = 0
            namef.SelLength = Len(namef.Text)
            Exit Sub
        End If
    End If

    ' Update changes
    If Task.Text = "M" Then
        'If RecordChanged() Then
        If GLindchange% Then
            TblMaster.Edit
            ScrToTbl
            MsgBox "Record Changed", 16, "Notification Only"
            GLindchange% = 0
        Else
            MsgBox "No Changes effected", 16, "Notification Only"
        End If
    End If
    FMcontact.Hide
    Exit Sub

BokErr:
    Load FMerror
    FMerror.Text1.Text = "FMcontact"
    FMerror.Text2.Text = Str(Err)
    FMerror.Show 1
    Resume Next

End Sub

Sub category_Change ()
    GLindchange% = 1
End Sub

Sub city_Change ()
    GLindchange% = 1
End Sub

Sub club_Change ()
    GLindchange% = 1
End Sub

Sub CSCombo1_change ()
    GLindchange% = 1
End Sub

Sub DefaultFields ()
    namel.Text = ""
    namef.Text = ""
    address1.Text = ""
    address2.Text = ""
    CSCombo1.Text = "QLD"
    city.Text = ""
    postcode.Text = ""
    phone.Text = ""
    fax.Text = ""
    phonewk.Text = ""
    faxwk.Text = ""
    club.Text = ""
    federation.Text = ""
    notes.Text = ""
    category = "P"
End Sub

Sub fax_Change ()
    GLindchange% = 1
End Sub

Sub faxwk_Change ()
    GLindchange% = 1
End Sub

Sub federation_Change ()
    GLindchange% = 1
End Sub

Sub Form_Activate ()

    Dim strCriterion As String

    On Error GoTo FMerror

    If Task.Text = "M" Then
        ButtDelete.Enabled = True
    Else
        ButtDelete.Enabled = False
    End If

    If Task.Text = "A" Then
        namel.Enabled = True
        namef.Enabled = True
        Keynamel.Text = ""
        Keynamef.Text = ""
        nullfields      'Null Screen fields
        DefaultFields   'Setup Default Field Values
        namef.SetFocus
    Else
        TblMaster.Seek "=", Keynamel.Text, Keynamef.Text 'Find the required record.
        TblToScr          'Move fields from table to screen
        namef.Enabled = False
        namel.Enabled = False
        address1.SetFocus
    End If
    If Task.Text = "C" Then
        namel.Enabled = True
        namef.Enabled = True
        namel = ""
        namef = ""
        namef.SetFocus
    End If
    Screen.MousePointer = 0
    GLindchange% = 0 'reset changed record flag
    Exit Sub

FMerror:
    Select Case Err
    Case 380
        Resume Next
    Case Else
        MsgBox "Error " & Err & " " & Error$ & " found.", 16, "Program Error"
        Exit Sub
    End Select
End Sub

Sub Form_Load ()
    
    On Error GoTo FMErr:
    
    'Open Database
    Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
    'Open Table
    Set TblMaster = DB.OpenTable("contacts")
    'Open Index
    TblMaster.Index = "contacts"

    ' Set Global Form Loaded Flag
    GLindchange% = 0

    Exit Sub

FMErr:
    Select Case Err
    Case 3022
        MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
    End Select
    
    Resume Next
 
End Sub

Sub namef_Change ()
    GLindchange% = 0
End Sub

Sub namel_Change ()
    GLindchange% = 1
End Sub

Sub notes_Change ()
    GLindchange% = 1
End Sub

Sub nullfields ()
    namel.Text = ""
    namef.Text = ""
    address1.Text = ""
    address2.Text = ""
    city.Text = ""
    postcode.Text = ""
    phone.Text = ""
    fax.Text = ""
    phonewk.Text = ""
    faxwk.Text = ""
    club.Text = ""
    federation.Text = ""
    notes.Text = ""
    category.Text = ""
    CSCombo1.Value = "QLD"
End Sub

Sub phone_Change ()
    GLindchange% = 1
End Sub

Sub phonewk_Change ()
    GLindchange% = 1
End Sub

Sub postcode_Change ()
    GLindchange% = 1
End Sub

Function RecordChanged () As Integer
    
    Dim X As Integer
    
    X% = 0
    
    'Cannot change key, don't worry about checking key

    If IsNull(TblMaster("address1")) And address1.Text = "" Then
    Else
        If IsNull(TblMaster("address1")) Then 'CrsDesc changed
            X% = X% + 1
        Else
            If Trim(address1.Text) <> Trim(TblMaster("address1")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("address2")) And address2.Text = "" Then
    Else
        If IsNull(TblMaster("address2")) Then 'CrsDesc changed
            X% = X% + 1
        Else
            If Trim(address2.Text) <> Trim(TblMaster("address1")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("city")) And city.Text = "" Then
    Else
        If IsNull(TblMaster("city")) Then 'CrsDesc changed
            X% = X% + 1
        Else
            If Trim(city.Text) <> Trim(TblMaster("city")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("state")) And CSCombo1.Text = "" Then
    Else
        If IsNull(TblMaster("state")) Then ' changed
            X% = X% + 1
        Else
            If Trim(CSCombo1.Text) <> Trim(TblMaster("state")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("postcode")) And postcode.Text = "" Then
    Else
        If IsNull(TblMaster("postcode")) Then ' changed
            X% = X% + 1
        Else
            If Trim(postcode.Text) <> Trim(TblMaster("postcode")) Then
                X% = X% + 1
            End If
        End If
    End If

    If IsNull(TblMaster("phone")) And phone.Text = "" Then
    Else
        If IsNull(TblMaster("phone")) Then ' changed
            X% = X% + 1
        Else
            If Trim(phone.Text) <> Trim(TblMaster("phone")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("fax")) And fax.Text = "" Then
    Else
        If IsNull(TblMaster("fax")) Then ' changed
            X% = X% + 1
        Else
            If Trim(fax.Text) <> Trim(TblMaster("fax")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("phone")) And phone.Text = "" Then
    Else
        If IsNull(TblMaster("phone")) Then ' changed
            X% = X% + 1
        Else
            If Trim(phone.Text) <> Trim(TblMaster("phone")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("faxwk")) And faxwk.Text = "" Then
    Else
        If IsNull(TblMaster("faxwk")) Then ' changed
            X% = X% + 1
        Else
            If Trim(faxwk.Text) <> Trim(TblMaster("faxwk")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("club")) And club.Text = "" Then
    Else
        If IsNull(TblMaster("club")) Then ' changed
            X% = X% + 1
        Else
            If Trim(club.Text) <> Trim(TblMaster("club")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("federation")) And federation.Text = "" Then
    Else
        If IsNull(TblMaster("federation")) Then ' changed
            X% = X% + 1
        Else
            If Trim(federation.Text) <> Trim(TblMaster("federation")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("notes")) And notes.Text = "" Then
    Else
        If IsNull(TblMaster("notes")) Then ' changed
            X% = X% + 1
        Else
            If Trim(notes.Text) <> Trim(TblMaster("notes")) Then
                X% = X% + 1
            End If
        End If
    End If
    If IsNull(TblMaster("category")) And category.Text = "" Then
    Else
        If IsNull(TblMaster("category")) Then ' changed
            X% = X% + 1
        Else
            If Trim(category.Text) <> Trim(TblMaster("category")) Then
                X% = X% + 1
            End If
        End If
    End If


'* below to be fixed

    If X% <> 0 Then
        RecordChanged = -1
    Else
        RecordChanged = 0
    End If

End Function

Function RecordChangedInAdd () As Integer
    Dim X As Integer
    'Bred for stock or race not checked yet
    ' also why not exit sub if change found
    'mdob <> 0 or "" ?

    X% = 0
    
    If namef.Text <> "" Then
        X% = X% + 1
    End If
    
    If namel.Text <> "" Then
        X% = X% + 1
    End If
    
    If address1.Text <> "" Then
        X% = X% + 1
    End If
    
    If address2.Text <> "" Then
        X% = X% + 1
    End If

    If CSCombo1.Value <> "" Then
        X% = X% + 1
    End If
    If city.Text <> "" Then
        X% = X% + 1
    End If
    If phone.Text <> "" Then
        X% = X% + 1
    End If
    If fax.Text <> "" Then
        X% = X% + 1
    End If
    If phonewk.Text <> "" Then
        X% = X% + 1
    End If
    If faxwk.Text <> "" Then
        X% = X% + 1
    End If
    
    If club.Text <> "" Then
        X% = X% + 1
    End If
    
    If federation.Text <> "" Then
        X% = X% + 1
    End If

    If notes.Text = "" Then
        X% = X% + 1
    End If

    If category.Text <> "" Then
        X% = X% + 1
    End If

    If X% <> 0 Then
        RecordChangedInAdd = -1
    Else
        RecordChangedInAdd = 0
    End If

End Function

Sub ScrToTbl ()
    If Task.Text = "A" Or Task.Text = "C" Then 'Only allow key change if new record
        TblMaster("namel") = Trim(namel.Text)
        TblMaster("namef") = Trim(namef.Text)
    End If
    TblMaster("address1") = Trim(address1.Text)
    TblMaster("address2") = Trim(address2.Text)
    TblMaster("city") = Trim(city.Text)
    TblMaster("state") = Trim(CSCombo1.Text)
    TblMaster("postcode") = Trim(postcode.Text)
    TblMaster("phone") = Trim(phone.Text)
    TblMaster("fax") = Trim(fax.Text)
    TblMaster("phonewk") = Trim(phonewk.Text)
    TblMaster("faxwk") = Trim(faxwk.Text)
    TblMaster("notes") = Trim(notes.Text)
    TblMaster("club") = Trim(club.Text)
    TblMaster("federation") = Trim(federation.Text)
    TblMaster("category") = Trim(category.Text)
    TblMaster.Update
End Sub

Sub SetAddMode ()
    ButtDelete.Enabled = False
    ButtCancel.Enabled = True
    ButtOK.Enabled = True

    namel.Enabled = True
    namef.Enabled = True
    address1.Enabled = True
    address2.Enabled = True
    city.Enabled = True
    CSCombo1.Enabled = True
    postcode.Enabled = True
    phone.Enabled = True
    fax.Enabled = True
    phonewk.Enabled = True
    faxwk.Enabled = True
    club.Enabled = True
    federation.Enabled = True
    notes.Enabled = True
    category.Enabled = True

    namef.SetFocus

End Sub

Sub SetEnqMode ()
    ButtDelete.Visible = False
    ButtCancel.Visible = True

    namel.Enabled = False
    namef.Enabled = False
    address1.Enabled = False
    address2.Enabled = False
    city.Enabled = False
    CSCombo1.Enabled = False
    postcode.Enabled = False
    phone.Enabled = False
    fax.Enabled = False
    phonewk.Enabled = False
    faxwk.Enabled = False
    club.Enabled = False
    federation.Enabled = False
    notes.Enabled = False
    category.Enabled = True

    ButtOK.SetFocus

End Sub

Sub SetUpdMode ()
    ButtDelete.Visible = True
    ButtCancel.Visible = True

    namel.Enabled = False
    namef.Enabled = False
    address1.Enabled = True
    address2.Enabled = True
    city.Enabled = True
    CSCombo1.Enabled = True
    postcode.Enabled = True
    phone.Enabled = True
    fax.Enabled = True
    phonewk.Enabled = True
    faxwk.Enabled = True
    club.Enabled = True
    federation.Enabled = True
    notes.Enabled = True
    category.Enabled = True
    
    address1.SetFocus

End Sub

Sub TblToScr ()
    If IsNull(TblMaster("namel")) Then
        namel.Text = ""
    Else
        namel.Text = TblMaster("namel")
    End If
    If IsNull(TblMaster("namef")) Then
        namef.Text = ""
    Else
        namef.Text = TblMaster("namef")
    End If

    If IsNull(TblMaster("address1")) Then
        address1.Text = ""
    Else
        address1.Text = TblMaster("address1")
    End If
    If IsNull(TblMaster("address2")) Then
        address2.Text = ""
    Else
        address2.Text = TblMaster("address2")
    End If

    If IsNull(TblMaster("city")) Then
        city.Text = ""
    Else
        city.Text = TblMaster("city")
    End If
    If IsNull(TblMaster("state")) Then
        CSCombo1.Text = ""
    Else
        CSCombo1.Text = TblMaster("state")
    End If
    If IsNull(TblMaster("postcode")) Then
        postcode.Text = ""
    Else
        postcode.Text = TblMaster("postcode")
    End If
    If IsNull(TblMaster("phone")) Then
        phone.Text = ""
    Else
        phone.Text = TblMaster("phone")
    End If
    If IsNull(TblMaster("fax")) Then
        fax.Text = ""
    Else
        fax.Text = TblMaster("fax")
    End If
    If IsNull(TblMaster("phonewk")) Then
        phonewk.Text = ""
    Else
        phonewk.Text = TblMaster("phonewk")
    End If
    If IsNull(TblMaster("faxwk")) Then
        faxwk.Text = ""
    Else
        faxwk.Text = TblMaster("faxwk")
    End If
    If IsNull(TblMaster("notes")) Then
        notes.Text = ""
    Else
        notes.Text = TblMaster("notes")
    End If
    If IsNull(TblMaster("club")) Then
        club.Text = ""
    Else
        club.Text = TblMaster("club")
    End If
    If IsNull(TblMaster("federation")) Then
        federation.Text = ""
    Else
        federation.Text = TblMaster("federation")
    End If
    If IsNull(TblMaster("category")) Then
        category.Text = ""
    Else
        category.Text = TblMaster("category")
    End If
End Sub

