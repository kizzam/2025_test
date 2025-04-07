VERSION 2.00
Begin Form FMaddracepnt 
   Caption         =   "Add Race Point"
   ClientHeight    =   3960
   ClientLeft      =   885
   ClientTop       =   1650
   ClientWidth     =   7335
   Height          =   4365
   Left            =   825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7335
   Top             =   1305
   Width           =   7455
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   7335
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      Begin CSComboBox Series 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMADDRAC.FRX:0000
         Contents        =   FMADDRAC.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   1800
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TextDisplay     =   0  'first visible
         Top             =   1920
         ValueCol        =   0
         Width           =   3015
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5520
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin SSCommand ButtOK 
         Caption         =   "&Save"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin TextBox lat_sec 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   5160
         TabIndex        =   5
         Top             =   2520
         Width           =   735
      End
      Begin TextBox lng_sec 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3000
         Width           =   735
      End
      Begin TextBox lng_min 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   7
         Top             =   3000
         Width           =   615
      End
      Begin TextBox lng_deg 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3000
         Width           =   615
      End
      Begin TextBox lat_min 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2520
         Width           =   615
      End
      Begin TextBox lat_deg 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2520
         Width           =   615
      End
      Begin TextBox pntname 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1200
         Width           =   3015
      End
      Begin TextBox code 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   3120
         Width           =   975
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   3120
         Width           =   855
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   3120
         Width           =   975
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude :"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   2640
         Width           =   855
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   2640
         Width           =   975
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Series : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    code.Text = ""
    pntname.Text = ""
    Series.Text = ""
    lat_deg.Text = ""
    lat_min.Text = ""
    lat_sec.Text = ""
    lng_deg.Text = ""
    lng_min.Text = ""
    lng_sec.Text = ""
    FMaddracepnt.Hide

End Sub

Sub ButtCancel_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click Cancel to tell system to cancel/ignor data entered/typed"
End Sub

Sub ButtOK_Click ()
    
    On Error GoTo msgspot

    'Check all data entered
    If Trim(code.Text) = "" Then
        MsgBox "Race Point Code MUST be entered", 16, "Add Race Point Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    End If
    If Trim(pntname.Text) = "" Then
        MsgBox "Race Point description MUST be entered", 16, "Add Race Point Error"
        pntname.SetFocus
        pntname.SelStart = 0
        pntname.SelLength = Len(pntname.Text)
        Exit Sub
    End If
    If Trim(Series) = "" Then
        MsgBox "Race Point Series MUST be entered", 16, "Add Race Point Error"
        Series.SetFocus
        'series.SelStart = 0
        'series.SelLength = Len(series.Text)
        Exit Sub
    End If
    If Trim(lat_deg) <> "" Then
        If IsNumeric(lat_deg) Then
            If Val(lat_deg) >= 0 And Val(lat_deg) < 360 Then
            ' Value OK
            Else
                MsgBox "Lat Degree - Invalid Value", 16, "Fix Latitude Degree"
                lat_deg.SetFocus
                lat_deg.SelStart = 0
                lat_deg.SelLength = Len(lat_deg.Text)
                Exit Sub
            End If
        Else
            MsgBox "Lat Degree - Invalid Value", 16, "Fix Latitude Degree"
            lat_deg.SetFocus
            lat_deg.SelStart = 0
            lat_deg.SelLength = Len(lat_deg.Text)
            Exit Sub
        End If
    End If
    If Trim(lat_min) <> "" Then
        If IsNumeric(lat_min) Then
            If Val(lat_min) >= 0 And Val(lat_min) <= 59 Then
            ' Value OK
            Else
                MsgBox "Lat Minutes - Invalid Value", 16, "Fix Value"
                lat_min.SetFocus
                lat_min.SelStart = 0
                lat_min.SelLength = Len(lat_min.Text)
                Exit Sub
            End If
        Else
            MsgBox "Lat Minutes - Invalid Value", 16, "Fix Value"
            lat_min.SetFocus
            lat_min.SelStart = 0
            lat_min.SelLength = Len(lat_min.Text)
            Exit Sub
        End If
    End If
    If Trim(lat_sec) <> "" Then
        If IsNumeric(lat_sec) Then
            If Val(lat_sec) >= 0 And Val(lat_sec) < 59.99999 Then
            ' Value OK
            Else
                MsgBox "Lat Seconds - Invalid Value", 16, "Fix Value"
                lat_sec.SetFocus
                lat_sec.SelStart = 0
                lat_sec.SelLength = Len(lat_sec.Text)
                Exit Sub
            End If
        Else
            MsgBox "Lat Minutes - Invalid Value", 16, "Fix Value"
            lat_sec.SetFocus
            lat_sec.SelStart = 0
            lat_sec.SelLength = Len(lat_sec.Text)
            Exit Sub
        End If
    End If

'Longitude
    If Trim(lng_deg) <> "" Then
        If IsNumeric(lng_deg) Then
            If Val(lng_deg) >= 0 And Val(lng_deg) < 360 Then
            ' Value OK
            Else
                MsgBox "Longitude Degree - Invalid Value", 16, "Fix lngitude Degree"
                lng_deg.SetFocus
                lng_deg.SelStart = 0
                lng_deg.SelLength = Len(lng_deg.Text)
                Exit Sub
            End If
        Else
            MsgBox "Longitude Degree - Invalid Value", 16, "Fix lngitude Degree"
            lng_deg.SetFocus
            lng_deg.SelStart = 0
            lng_deg.SelLength = Len(lng_deg.Text)
            Exit Sub
        End If
    End If
    If Trim(lng_min) <> "" Then
        If IsNumeric(lng_min) Then
            If Val(lng_min) >= 0 And Val(lng_min) <= 59 Then
            ' Value OK
            Else
                MsgBox "Longitude Minutes - Invalid Value", 16, "Fix Value"
                lng_min.SetFocus
                lng_min.SelStart = 0
                lng_min.SelLength = Len(lng_min.Text)
                Exit Sub
            End If
        Else
            MsgBox "Longitude Minutes - Invalid Value", 16, "Fix Value"
            lng_min.SetFocus
            lng_min.SelStart = 0
            lng_min.SelLength = Len(lng_min.Text)
            Exit Sub
        End If
    End If
    If Trim(lng_sec) <> "" Then
        If IsNumeric(lng_sec) Then
            If Val(lng_sec) >= 0 And Val(lng_sec) < 59.99999 Then
            ' Value OK
            Else
                MsgBox "Longitude Seconds - Invalid Value", 16, "Fix Value"
                lng_sec.SetFocus
                lng_sec.SelStart = 0
                lng_sec.SelLength = Len(lng_sec.Text)
                Exit Sub
            End If
        Else
            MsgBox "Longitude Minutes - Invalid Value", 16, "Fix Value"
            lng_sec.SetFocus
            lng_sec.SelStart = 0
            lng_sec.SelLength = Len(lng_sec.Text)
            Exit Sub
        End If
    End If

    'Add new record
    FMracepnt.Data1.Recordset.AddNew
    FMracepnt.Data1.Recordset("code_pnt") = code.Text
    FMracepnt.Data1.Recordset("pntname") = pntname.Text
    FMracepnt.Data1.Recordset("series") = Series
    FMracepnt.Data1.Recordset("lat_deg") = Val(lat_deg.Text)
    FMracepnt.Data1.Recordset("lat_min") = Val(lat_min.Text)
    FMracepnt.Data1.Recordset("lat_sec") = Val(lat_sec.Text)
    FMracepnt.Data1.Recordset("lng_deg") = Val(lng_deg.Text)
    FMracepnt.Data1.Recordset("lng_min") = Val(lng_min.Text)
    FMracepnt.Data1.Recordset("lng_sec") = Val(lng_sec.Text)
    FMracepnt.Data1.Recordset.Update
    FMracepnt.Data1.Refresh
    code.Text = ""
    pntname.Text = ""
    Series = ""
    lat_deg.Text = ""
    lat_min.Text = ""
    lat_sec.Text = ""
    lng_deg.Text = ""
    lng_min.Text = ""
    lng_sec.Text = ""

    FMaddracepnt.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
        MsgBox "That Race Point already exists.", 16, "Add Race Point Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Race Point Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Race Point Error"
    Case Else
        Load FMerror
        FMerror.Text1.Text = "FMbird"
        FMerror.Text2.Text = Str(Err)
        FMerror.Show 1
        Resume Next
    End Select

    Exit Sub


End Sub

Sub ButtOK_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click OK to Accept details entered"
End Sub

Sub code_LostFocus ()
code.Text = UCase(code.Text)
End Sub

Sub code_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Enter Unique Race Point Code here"
End Sub

Sub Form_Activate ()
    code.SetFocus
End Sub

Sub Form_Unload (Cancel As Integer)
    FMracepnt.Table1.Refresh
End Sub

Sub lat_deg_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Type race point Latitude Degrees here"
End Sub

Sub lat_min_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
        MsgPanel.Caption = "Type race point Latitude Minutes here"
End Sub

Sub lat_sec_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Type race point Latitude Seconds here"
End Sub

Sub lng_deg_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
        MsgPanel.Caption = "Type race point Longitude Degrees here"
End Sub

Sub lng_min_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
        MsgPanel.Caption = "Type race point Longitude Minutes here"
End Sub

Sub lng_sec_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
        MsgPanel.Caption = "Type race point Longitude Seconds here"
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = ""
End Sub

Sub pntname_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Enter/Type Race Point name here"
End Sub

Sub series_LostFocus ()
'Series.Text = UCase(Series.Text)
End Sub

Sub series_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Enter/Type the direction this point is"
End Sub

