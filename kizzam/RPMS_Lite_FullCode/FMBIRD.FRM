VERSION 2.00
Begin Form FMbird 
   Caption         =   "Bird Maintenance/Enquiry"
   ClientHeight    =   8370
   ClientLeft      =   690
   ClientTop       =   1815
   ClientWidth     =   11220
   ClipControls    =   0   'False
   Height          =   8880
   Left            =   630
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   11220
   Top             =   1365
   Width           =   11340
   Begin SSPanel Panel3D2 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   7920
      Width           =   11175
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11175
      Begin SSCommand ButtDams 
         Caption         =   "Dams-Stock &Hens"
         Font3D          =   3  'Inset w/light shading
         Height          =   255
         Left            =   7680
         TabIndex        =   31
         Top             =   7440
         Width           =   1695
      End
      Begin SSCommand ButtSires 
         Caption         =   "Sires-Stock &Cocks"
         Font3D          =   3  'Inset w/light shading
         Height          =   255
         Left            =   7680
         TabIndex        =   27
         Top             =   7080
         Width           =   1695
      End
      Begin SSCommand ButtOk 
         Caption         =   "&Save"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   615
         Left            =   9600
         TabIndex        =   35
         Top             =   1200
         Width           =   1215
      End
      Begin SSCommand ButtDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   615
         Left            =   9600
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin SSCommand cmdDate 
         Caption         =   "D"
         Font3D          =   3  'Inset w/light shading
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   1920
         Width           =   255
      End
      Begin CSVList CSVListSires 
         AutoRecordCount =   -1  'True
         BorderStyle     =   1  'Fixed Single
         ColDelim        =   "-"
         ColWidth        =   FMBIRD.FRX:0000
         Height          =   1455
         HScroll         =   0   'False
         Left            =   3240
         ListCount       =   0
         ListDataFields  =   "yr,mark,ringno"
         ListDataSource  =   "Data5"
         MultiSelect     =   0  'None
         ShadowColor     =   &H00808080&
         TabIndex        =   66
         Top             =   5640
         ValueCol        =   -1
         Width           =   4335
      End
      Begin CSVList CSVListDams 
         AutoRecordCount =   -1  'True
         BorderStyle     =   1  'Fixed Single
         ColDelim        =   "-"
         ColWidth        =   FMBIRD.FRX:0040
         Height          =   1335
         HScroll         =   0   'False
         Left            =   3240
         ListCount       =   0
         ListDataFields  =   "yr,mark,ringno"
         ListDataSource  =   "Data6"
         MultiSelect     =   0  'None
         ShadowColor     =   &H00808080&
         TabIndex        =   67
         Top             =   6120
         ValueCol        =   -1
         Width           =   4335
      End
      Begin SSPanel LabDam 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Dam (Mother) from list"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   1335
         Left            =   2040
         TabIndex        =   64
         Top             =   6120
         Width           =   1215
      End
      Begin SSPanel LabSire 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Sire (Father) from list"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   1455
         Left            =   2040
         TabIndex        =   65
         Top             =   5640
         Width           =   1215
      End
      Begin VideoSoftIndexTab VSIndexTab1 
         BackTabColor    =   &H00808080&
         BorderWidth     =   1
         Caption         =   "Status|Bloodline|Notes"
         CurrTab         =   2
         ForeColor       =   &H00C0C0C0&
         FrontTabColor   =   &H00C0C0C0&
         Height          =   3855
         Left            =   1200
         Position        =   0  'Top
         Style           =   5  'Chamfered 3D
         TabIndex        =   21
         TabsPerPage     =   3
         Top             =   3000
         Width           =   9615
         Begin TextBox Mnotes 
            ForeColor       =   &H00C00000&
            Height          =   3450
            Left            =   60
            MaxLength       =   640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   345
            Width           =   9495
         End
         Begin VideoSoftElastic VSElastic2 
            BackColor       =   &H00C0C0C0&
            BevelInner      =   0  'None
            BevelOuter      =   0  'None
            Height          =   3450
            IntBkg          =   &H00C0C0C0&
            Left            =   -15645
            TabIndex        =   63
            Top             =   345
            Width           =   9495
            Begin SSCheck Mindref 
               Caption         =   "Reference"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   360
               TabIndex        =   12
               Top             =   600
               Width           =   1215
            End
            Begin SSCheck Mindsp2 
               Caption         =   "Special 2"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5880
               TabIndex        =   18
               Top             =   600
               Width           =   1095
            End
            Begin SSOption Mindbreds 
               Caption         =   "Bred for Stock"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2040
               TabIndex        =   19
               Top             =   960
               Width           =   1575
            End
            Begin SSOption Mindbredr 
               Caption         =   "Bred for Racing"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3720
               TabIndex        =   20
               Top             =   960
               Width           =   1575
            End
            Begin SSCheck Mindbpa 
               Caption         =   "Breeders Plate (A)"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3960
               TabIndex        =   15
               Top             =   240
               Width           =   1815
            End
            Begin SSCheck Mindbpb 
               Caption         =   "Breeders Plate (B)"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3960
               TabIndex        =   16
               Top             =   600
               Width           =   1815
            End
            Begin SSCheck Mindsp1 
               Caption         =   "Special 1"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5880
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin SSCheck Mindsoh 
               Caption         =   "Current"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   360
               TabIndex        =   11
               Top             =   240
               Width           =   1215
            End
            Begin SSCheck Mindrace 
               Caption         =   "Race Bird"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1920
               TabIndex        =   13
               Top             =   240
               Width           =   1095
            End
            Begin SSCheck Mindstock 
               Caption         =   "Stock Bird"
               Font3D          =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1920
               TabIndex        =   14
               Top             =   600
               Width           =   1215
            End
         End
         Begin VideoSoftElastic VSElastic1 
            BackColor       =   &H00C0C0C0&
            BevelInnerWidth =   0
            BevelOuter      =   0  'None
            BorderWidth     =   1
            Height          =   3450
            IntBkg          =   &H00C0C0C0&
            Left            =   -15570
            TabIndex        =   24
            Top             =   345
            Width           =   9495
            Begin SSCommand cmdBLine 
               Caption         =   "&Add New Bloodline"
               Font3D          =   3  'Inset w/light shading
               Height          =   375
               Left            =   5280
               TabIndex        =   22
               Top             =   240
               Width           =   1935
            End
            Begin CSComboBox CSCombo5 
               AutoSearch      =   0  'No Search
               BorderEffect    =   0  'None
               ColDelim        =   ";"
               ColWidth        =   FMBIRD.FRX:0080
               Contents        =   FMBIRD.FRX:00C0
               Dropped         =   0   'False
               ForeColor       =   &H00C00000&
               FoundIndex      =   0
               Height          =   405
               Left            =   1560
               ListBoxHeight   =   0
               ListBoxWidth    =   0
               ListDataFields  =   "Desc,Code"
               ListDataSource  =   "Data1"
               SearchString    =   ""
               Sorted          =   -1  'True
               StartSearch     =   0
               Style           =   2  'Dropdown List
               TabIndex        =   23
               TextDisplay     =   0  'first visible
               Top             =   240
               ValueCol        =   1
               Width           =   3135
            End
            Begin TextBox Mbloodline 
               Enabled         =   0   'False
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   1560
               MaxLength       =   26
               TabIndex        =   25
               Top             =   720
               Width           =   3135
            End
            Begin Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bloodline: "
               Height          =   255
               Left            =   240
               TabIndex        =   62
               Top             =   240
               Width           =   1215
            End
            Begin Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Description: "
               Height          =   255
               Left            =   240
               TabIndex        =   61
               Top             =   720
               Width           =   1215
            End
         End
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "BloodLine"
         Top             =   6240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Mdob 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "  /  /  "
         Top             =   1920
         Width           =   1335
      End
      Begin Data Data6 
         Caption         =   "Data6"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Data Data5 
         Caption         =   "Data5"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   7200
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Mdamyr 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   32
         Top             =   7440
         Width           =   615
      End
      Begin TextBox Msireyr 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   28
         Top             =   7080
         Width           =   615
      End
      Begin TextBox Mdamring 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6240
         MaxLength       =   8
         TabIndex        =   34
         Top             =   7440
         Width           =   1335
      End
      Begin TextBox Mdammark 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4320
         MaxLength       =   12
         TabIndex        =   33
         Top             =   7440
         Width           =   1095
      End
      Begin TextBox Msirering 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6240
         MaxLength       =   8
         TabIndex        =   30
         Top             =   7080
         Width           =   1335
      End
      Begin TextBox Msiremark 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4320
         MaxLength       =   12
         TabIndex        =   29
         Top             =   7080
         Width           =   1095
      End
      Begin CSComboBox CSCombo4 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMBIRD.FRX:00C2
         Contents        =   FMBIRD.FRX:0102
         Dropped         =   0   'False
         ForeColor       =   &H00C00000&
         FoundIndex      =   0
         Height          =   405
         Left            =   4440
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "desc,code"
         ListDataSource  =   "Data4"
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TextDisplay     =   0  'first visible
         Top             =   1440
         ValueCol        =   1
         Width           =   2415
      End
      Begin CSComboBox CSCombo3 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMBIRD.FRX:0104
         Contents        =   FMBIRD.FRX:0144
         Dropped         =   0   'False
         ForeColor       =   &H00C00000&
         FoundIndex      =   0
         Height          =   405
         Left            =   4440
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "desc,code"
         ListDataSource  =   "Data3"
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TextDisplay     =   0  'first visible
         Top             =   1080
         ValueCol        =   1
         Width           =   2415
      End
      Begin CSComboBox CSCombo2 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMBIRD.FRX:0146
         Contents        =   FMBIRD.FRX:0186
         Dropped         =   0   'False
         ForeColor       =   &H00C00000&
         FoundIndex      =   0
         Height          =   405
         Left            =   1200
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "desc,code"
         ListDataSource  =   "Data2"
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TextDisplay     =   0  'first visible
         Top             =   2400
         ValueCol        =   1
         Width           =   3495
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMBIRD.FRX:0188
         Contents        =   FMBIRD.FRX:01C8
         Dropped         =   0   'False
         ForeColor       =   &H00C00000&
         FoundIndex      =   0
         Height          =   405
         Left            =   1200
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TextDisplay     =   0  'first visible
         Top             =   1440
         ValueCol        =   1
         Width           =   1815
      End
      Begin TextBox Key_ring_no 
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Text            =   "Key_ring_no"
         Top             =   7080
         Visible         =   0   'False
         Width           =   975
      End
      Begin TextBox Key_mark 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Text            =   "Key_mark"
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
      End
      Begin TextBox Key_yr 
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Text            =   "Key_yr"
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
      End
      Begin TextBox Task 
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Text            =   "Task"
         Top             =   7440
         Visible         =   0   'False
         Width           =   855
      End
      Begin TextBox Mring 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "Text7"
         Top             =   360
         Width           =   1575
      End
      Begin TextBox Mmark 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "Text6"
         Top             =   360
         Width           =   2055
      End
      Begin TextBox Myr 
         BackColor       =   &H00E0FFFF&
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Te"
         Top             =   360
         Width           =   855
      End
      Begin Data Data4 
         Caption         =   "Data4"
         Connect         =   ""
         DatabaseName    =   "\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "eyesign"
         Top             =   6960
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Data Data3 
         Caption         =   "Data3"
         Connect         =   ""
         DatabaseName    =   "\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "colour"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Data Data2 
         Caption         =   "Data2"
         Connect         =   ""
         DatabaseName    =   "\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   9840
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "cage"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Mbredby 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   24
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin TextBox Mname 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No. :"
         Height          =   255
         Left            =   5640
         TabIndex        =   60
         Top             =   7440
         Width           =   495
      End
      Begin Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No. :"
         Height          =   255
         Left            =   5640
         TabIndex        =   59
         Top             =   7080
         Width           =   495
      End
      Begin Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mark :"
         Height          =   255
         Left            =   3600
         TabIndex        =   58
         Top             =   7440
         Width           =   615
      End
      Begin Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mark :"
         Height          =   255
         Left            =   3600
         TabIndex        =   57
         Top             =   7080
         Width           =   615
      End
      Begin Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ring No. : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mark : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   51
         Top             =   360
         Width           =   735
      End
      Begin Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   360
         Width           =   735
      End
      Begin Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Eye Sign : "
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   1440
         Width           =   975
      End
      Begin Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cage : "
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   2400
         Width           =   735
      End
      Begin Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Parentage - "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   6960
         Width           =   1335
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dam - Yr :"
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   7440
         Width           =   975
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sire  - Yr :"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   7080
         Width           =   975
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   1080
         Width           =   735
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth DOB. : "
         Height          =   255
         Left            =   3840
         TabIndex        =   43
         Top             =   1920
         Width           =   1815
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bred by : "
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1920
         Width           =   855
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sex : "
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Colour : "
         Height          =   255
         Left            =   3600
         TabIndex        =   40
         Top             =   1080
         Width           =   735
      End
   End
End
Option Explicit
Dim TblMaster As table
Dim DB As Database
Dim GLindchange As Integer

Sub ButtCancel_Click ()
    Dim reply As Integer, zz As Integer

    If FMBird.Task.Text = "A" Then
	'If RecordChangedInAdd() Then
	If GLindchange% Then ' has any fields changed
	    reply = MsgBox("Do you want to discard new record?", 4, "Warning?")
	    If reply <> 6 Then
		myr.SetFocus
		Exit Sub
	    End If
	Else
	    MsgBox "Add Record Cancelled", 16, "Notification Only"
	End If
    End If
    
    If FMBird.Task.Text = "M" Then
	'If RecordChanged() Then
	If GLindchange% Then ' has any fields changed
	    reply = MsgBox("Do you want to cancel changes?", 4, "Warning?")
	    If reply <> 6 Then
		Mname.SetFocus
		Exit Sub
	    End If
	End If
    End If

    FMBird.Hide
    'Unload FMbird
End Sub

Sub ButtDams_Click ()
    If CsvListDams.Visible = True Then
	VSIndexTab1.Visible = True
	CsvListDams.Enabled = False
	CsvListDams.Visible = False
	LabDam.Visible = False
    Else
	CsvListSires.Visible = False
	LabSire.Visible = False
	VSIndexTab1.Visible = False
	CsvListDams.Enabled = True
	CsvListDams.Visible = True
	LabDam.Visible = True
	CsvListDams.Refresh
    End If
End Sub

Sub ButtDelete_Click ()
    Dim reply
    Dim tbl1 As snapshot
    Dim tbl2 As snapshot
    Dim Criteria As String

    Criteria$ = "yr = " & TblMaster("yr") & " AND mark = '" & TblMaster("mark") & "' AND ringno = '" & TblMaster("ringno") & "'"
    reply = MsgBox("Confirm Deletion of record ", 4, "WARNING!")
    If reply = 6 Then
	'check to see if bird is used in TEAM
	Set tbl1 = DB.CreateSnapshot("racetrans")     'Open Table
	tbl1.FindFirst Criteria$
	If tbl1.NoMatch Then
	    'check if bird has been raced
	    Set tbl2 = DB.CreateSnapshot("teambird")     'Open Table
	    tbl2.FindFirst Criteria$
	    If tbl2.NoMatch Then
		'check if bird has been raced
		TblMaster.Delete
		MsgBox "Record Deleted", 16, "Notification Only"
		FMbrreg.RefreshData.Text = "Y"
	    Else
		MsgBox "Cannot delete while bird is member of team code - '" & tbl2("code") & "'", 16, "Warning - Cannot Delete record"
	    End If
	    tbl2.Close
	    Set tbl2 = Nothing
	Else
	    MsgBox "Cannot delete while bird has RACE records on file - Race Code - '" & tbl1("racecode") & "'", 16, "Warning - Cannot Delete record"
	End If
	tbl1.Close
	Set tbl1 = Nothing
	FMBird.Hide
    Else
	Mname.SetFocus
    End If

End Sub

Sub ButtOK_Click ()
    On Error GoTo BokErr

    'Validate Data entry
    If FMBird.Task.Text = "A" Or Task.Text = "C" Then
	If Trim(myr.Text = "") Then
	    MsgBox "Ring Year cannot be empty", 16, "Error: Cannot update record"
	    myr.SetFocus
	    Exit Sub
	End If
	If Val(Len(myr)) < 4 Then 'if numeric value < 4 chars
	    MsgBox "Ring Year MUST include century i.e. put in 2000 NOT 00", 0, "WARNING Year incorrect"
	    myr.SetFocus
	End If
	If Trim(Mmark.Text = "") Then
	    MsgBox "Ring Mark cannot be empty", 16, "Error: Cannot update record"
	    Mmark.SetFocus
	    Exit Sub
	End If
	If Trim(Mring.Text = "") Then
	    MsgBox "Ring Number cannot be empty", 16, "Error: Cannot update record"
	    Mring.SetFocus
	    Exit Sub
	End If
	'etc add in all validation for Add, maybe set up in Add Val func
    End If
    If FMBird.Task.Text = "M" Or Task.Text = "A" Or Task.Text = "C" Then
	If Trim(CSCombo3.Value = "") Then
	    MsgBox "Colour MUST NOT be empty", 16, "Error: Cannot update record"
	    CSCombo3.SetFocus
	    Exit Sub
	End If
    
	'Validate DOB
	'if year is entered as 00 then assume 2000 otherwise accept
	'still have slight problem with 29/02/00 assumes 1900 for isdate function & reports as invalid
	'didn't worry about this

	If Trim(mDOB.Text) = "" Then
	    mDOB.Text = "  /  /  "
	    MsgBox "WARNING - Date can't be empty", 0, "Date Of Birth required"
	    mDOB.SetFocus
	    mDOB.SelStart = 0
	    mDOB.SelLength = Len(mDOB.Text)
	    Exit Sub
	End If
    
	'Fix potential leap year problem otherwise won't pass isdate() check
	If mDOB.Text = "29/2/00" Or mDOB.Text = "29/02/00" Then
	    mDOB.Text = DateAdd("yyyy", 100, mDOB.Text)
	End If
    
	If IsDate(mDOB.Text) Then 'Must be Valid
	    mDOB.Text = DateFix(mDOB.Text)
	Else
	    MsgBox "WARNING - Date of Birth Invalid", 0, "Date Of Birth format incorrect"
	    mDOB.SetFocus
	    mDOB.SelStart = 0
	    mDOB.SelLength = Len(mDOB.Text)
	    Exit Sub
	End If

    'etc add in all validation for ADD & Mod maybe set up in func
    End If

    'Add record
    If FMBird.Task.Text = "A" Or Task.Text = "C" Then
	TblMaster.Seek "=", myr, Mmark, Mring 'Ensure unique index
	If TblMaster.NoMatch Then
	    TblMaster.AddNew
	    ScrToTbl
	    MsgBox "Record Added", 16, "Notification Only"
	Else
	    MsgBox "Cannot add record, Code already used", 16, "Duplicate Key"
	    myr.SetFocus
	    myr.SelStart = 0
	    myr.SelLength = Len(myr.Text)
	Exit Sub
	End If
    End If

    ' Update changes
    If FMBird.Task.Text = "M" Then
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
    FMBird.Hide
    'Unload FMbird
    Exit Sub

BokErr:
    Load FMerror
    FMerror.Text1.Text = "FMbird"
    FMerror.Text2.Text = Str(Err)
    FMerror.Show 1
    Resume Next

End Sub

Sub ButtSires_Click ()
    If CsvListSires.Visible = True Then
	VSIndexTab1.Visible = True
	CsvListSires.Enabled = False
	CsvListSires.Visible = False
	LabSire.Visible = False
    Else
	CsvListDams.Visible = False
	LabDam.Visible = False
	VSIndexTab1.Visible = False
	CsvListSires.Enabled = True
	CsvListSires.Visible = True
	LabSire.Visible = True
	CsvListSires.Refresh
    End If
End Sub

Sub cmdBLine_Click ()
Dim Mring As String
Mring$ = FMBird.Mring.Text
FMBline.Show 1
Data1.Refresh
FMBird.Mring.Text = Mring$
DoEvents
End Sub

Sub cmdDate_Click ()
Dim reply As Integer
reply% = MsgBox("Do you wish to set date of birth to Generic Breeding season date", 36, "Click 'Yes' to change date")
If reply% = 6 Then
    mDOB.Text = Format("Jun 01, " & myr.Text, "ddddd")
End If
End Sub

Sub CSCombo1_Change ()
    GLindchange% = -1
End Sub

Sub CSCombo2_Change ()
    GLindchange% = -1
End Sub

Sub CSCombo3_Change ()
    GLindchange% = -1
End Sub

Sub CSCombo4_Change ()
    GLindchange% = -1
End Sub

Sub CSCombo5_Change ()
    GLindchange% = -1
End Sub

Sub CSVListDams_Click ()
    Dim Pos1 As Integer
    Dim Pos2 As Integer
    Dim MarkLen As Integer
    Dim KeyStr As String

    KeyStr$ = CsvListDams.Value
    
    'Extract Dam Yr-Mark-Ring
    Pos1% = InStr(1, KeyStr$, "-")              ' Find first hyphen
    Pos2% = InStr(Pos1% + 1, KeyStr$, "-")      ' Find next hypen
    MarkLen% = (Pos2% - Pos1%) - 1              ' Calculate 2nd word length.
    Mdamyr.Text = Val(Left(KeyStr$, Pos1% - 1))           ' Calculate 1st word length.
    Mdamring.Text = Right$(KeyStr$, Len(KeyStr$) - Pos2%)  'Calculate 3rd word length.
    Mdammark.Text = Mid$(KeyStr$, Pos1% + 1, MarkLen%) ' Return 2nd word.
    CsvListDams.Enabled = False
    CsvListDams.Visible = False
    LabDam.Visible = False
    VSIndexTab1.Visible = True
End Sub

Sub CSVListSires_Click ()
    Dim Pos1 As Integer
    Dim Pos2 As Integer
    Dim MarkLen As Integer
    Dim KeyStr As String

    KeyStr$ = CsvListSires.Value
    
    'Extract Dam Yr-Mark-Ring
    Pos1% = InStr(1, KeyStr$, "-")              ' Find first hyphen
    Pos2% = InStr(Pos1% + 1, KeyStr$, "-")      ' Find next hypen
    MarkLen% = (Pos2% - Pos1%) - 1              ' Calculate 2nd word length.
    Msireyr.Text = Val(Left(KeyStr$, Pos1% - 1))           ' Calculate 1st word length.
    Msirering.Text = Right$(KeyStr$, Len(KeyStr$) - Pos2%)  'Calculate 3rd word length.
    Msiremark.Text = Mid$(KeyStr$, Pos1% + 1, MarkLen%) ' Return 2nd word.
    CsvListSires.Enabled = False
    CsvListSires.Visible = False
    LabSire.Visible = False
    VSIndexTab1.Visible = True
End Sub

Sub DefaultFields ()
    Let myr.Text = Trim(Str$(Year(Now)))
    Let Mmark.Text = GLMark$
    Let Mring.Text = ""
    Let Mname.Text = ""
    Let CSCombo1.Text = "Unsexed"   'Sex
    CSCombo1.Value = "U"
    CSCombo2.Text = "Not Available" ' Cage
    CSCombo2.Value = "NA"
    CSCombo3.Text = "Not Available" ' Colour
    CSCombo3.Value = "NA"
    CSCombo4.Text = "Not Available"  'eyesign
    CSCombo4.Value = "NA"
    CSCombo5.Text = "Not Available"  'Codes
    CSCombo5.Value = "NA"

    Mbredby.Text = GLOwner$
    'Mbredfor.Text = GLOwner$
    Mbloodline.Text = ""
    mIndSOH.Value = 1
    mIndRace.Value = 1
    mIndStock.Value = 0
    mIndRef.Value = 0
    mIndBpa.Value = 0
    mIndBpb.Value = 0
    mIndSp1.Value = 0
    mIndSp2.Value = 0
    mIndBreds.Value = 0
    mIndBredr.Value = 1
    'Mindstatus bredfor stock or racing?
    Mdamyr.Text = ""
    Mdammark.Text = ""
    Mdamring.Text = ""
    Msireyr.Text = ""
    Msiremark.Text = ""
    Msirering.Text = ""
    Mnotes.Text = ""
    'Mracelast.Text = "  /  /  "
    mDOB.Text = Format$(Now, "ddddd")

End Sub

Private Function FindRecord (ByVal Key_yr As Integer, ByVal Key_mark As String, ByVal Key_ring_no As String) As Integer
On Error GoTo errorsys2:
Dim key As String
    
    key$ = "yr = " & Key_yr
    key$ = key$ & " And " & "mark = '" & Key_mark & "' And"
    key$ = key$ & " " & "ringno = '" & Key_ring_no & "'"
    Exit Function

errorsys2:
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
 
End Function

Sub Form_Activate ()

    Dim strCriterion As String
    Dim Action As String

    On Error GoTo FMerror
    
    'Pickup any new Stock cocks & hens
    Action$ = "Stock Cocks"
    Data5.Refresh
    Action$ = "Stock Hens"
    Data6.Refresh
    Action = ""

    CsvListSires.Visible = False
    LabSire.Visible = False
    CsvListDams.Visible = False
    LabDam.Visible = False
    VSIndexTab1.Visible = True
    VSIndexTab1.CurrTab = 0

    If FMBird.Task.Text = "M" Then
	ButtDelete.Enabled = True
	ButtDelete.Visible = True
    Else
	ButtDelete.Visible = False
	ButtDelete.Enabled = False
    End If

    If FMBird.Task.Text = "A" Then
	myr.Enabled = True
	Mring.Enabled = True
	Mmark.Enabled = True
	FMBird.Key_yr.Text = ""
	FMBird.Key_mark.Text = ""
	FMBird.Key_ring_no.Text = ""
	nullfields      'Null Screen fields
	DefaultFields   'Setup Default Field Values
	cmdBLine.Visible = False
	myr.SetFocus
    Else
	TblMaster.Seek "=", Key_yr.Text, Key_mark.Text, Key_ring_no.Text 'Find the required record.
	TblToScr          'Move fields from table to screen
	myr.Enabled = False
	Mring.Enabled = False
	Mmark.Enabled = False
	cmdBLine.Visible = True
	Mname.SetFocus
    End If
    If FMBird.Task.Text = "C" Then
	myr.Enabled = True
	Mring.Enabled = True
	Mring = ""
	Mmark.Enabled = True
	Mring.SetFocus
    End If
    Screen.MousePointer = 0
    GLindchange% = 0 'reset changed record flag
    mIndSOH.Refresh
    mIndRace.Refresh
    mIndBpa.Refresh
    mIndBpb.Refresh
    mIndRef.Refresh
    mIndSp1.Refresh
    mIndSp2.Refresh
    mIndStock.Refresh
    mIndBredr.Refresh
    mIndBreds.Refresh
    
    Exit Sub

FMerror:
    Select Case Err
    Case 380
	Resume Next
    Case 3021
	MsgBox "No " & Action$ & " on file - Database File Empty", 16, "INFORMATION ONLY"
	Resume Next
    Case 3022
	MsgBox "Record already exists", 16, "Add record Error"
	Resume Next
	Case 3026
	MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
	Resume Next
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
	Resume Next
    Case Else
	MsgBox "Error" & Err & " " & Error$ & " found.", 16, "Program Error code " & Val(Error) & " "
    End Select
    Resume Next
End Sub

Sub Form_Load ()
    
    Dim Sires As String
    Dim Dams As String
    Dim Action As String

    On Error GoTo FMErr:
    
    ' Set Global Form Loaded Flag
    GLLfmbird% = -1     'ie indicate form is loaded
    GLindchange% = 0
    'Open Database
    Action$ = "Master file"
    Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
    Set TblMaster = DB.OpenTable("master")     'Open Table
    TblMaster.Index = "bird_ymr"              'Open Index

    Action$ = "Bloodline"
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "BloodLine"
    Data1.Refresh
    Action$ = "Cage/Loft"
    Data2.DatabaseName = DBFullFileName$
    Data2.RecordSource = "cage"
    Data2.Refresh
    Data3.DatabaseName = DBFullFileName$
    Action$ = "Colour"
    Data3.RecordSource = "colour"
    Data3.Refresh
    Action$ = "Eyesign"
    Data4.DatabaseName = DBFullFileName$
    Data4.RecordSource = "eyesign"
    Data4.Refresh
    
    Data5.DatabaseName = DBFullFileName$
    Data6.DatabaseName = DBFullFileName$

    Action$ = "Stock Hens"
    'Create Snapshot of Hens
    Dams$ = "SELECT * FROM master WHERE codesex = 'H' and indstock = -1 and indsoh = -1 ORDER BY yr ASC, mark, ringno;" 'or Primary key
    'Set TblDam = DB.CreateSnapshot(Dams$)
    'Data6.RecordSource = TblDam
    Data6.RecordSource = Dams$
    'Data6.Refresh
    'CsvListDams.ListDataFields = "Yr,Mark,Ringno"
    CsvListDams.Visible = False
    CsvListDams.Refresh

    Action$ = "Stock Cocks"
    'Create Snapshot of Cocks
    Sires$ = "SELECT * FROM master WHERE codesex = 'C' and indstock = -1 and indsoh = -1 ORDER BY yr ASC, mark, ringno;" ' or Primary key"
    'Set TblSire = DB.CreateSnapshot(Sires$)
    'Data5.RecordSource = TblSire
    Data5.RecordSource = Sires$
    'Data5.Refresh
    'CsvListSires.ListDataFields = "Yr,Mark,Ringno"
    CsvListSires.Visible = False
    CsvListSires.Refresh
    Action$ = ""
    
    Exit Sub

FMErr:
    Select Case Err
    Case 3021
	MsgBox "No " & Action$ & " on file - Database File Empty", 16, "INFORMATION ONLY"
	Resume Next
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

Sub Form_Unload (Cancel As Integer)
    'Close Database
'    Set TblSire = Nothing
'    Set TblDam = Nothing
    TblMaster.Close
    DB.Close

    ' Set Global Form Loaded Flag
    GLLfmbird% = 0     'ie indicate form is unloaded
'    Unload FMbird
    Set FMBird = Nothing
End Sub

Sub Mbloodline_Change ()
    GLindchange% = -1
End Sub

Sub Mbredby_Change ()
    GLindchange% = -1
End Sub

Sub Mdammark_Change ()
    GLindchange% = -1
End Sub

Sub Mdammark_GotFocus ()
    CsvListSires.Visible = False
End Sub

Sub Mdammark_LostFocus ()
    If GLindchange% = -1 Then
	Mdammark = UCase(Mdammark)
    End If
End Sub

Sub Mdamring_Change ()
    GLindchange% = -1
End Sub

Sub Mdamring_GotFocus ()
    CsvListSires.Visible = False
End Sub

Sub Mdamring_LostFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mdamyr_Change ()
    GLindchange% = -1
End Sub

Sub Mdamyr_GotFocus ()
    CsvListSires.Visible = False
End Sub

Sub Mdamyr_LostFocus ()
    Mdamyr = Trim(Mdamyr)
    If Trim(Mdamyr) <> "" Then
	If Val(Mdamyr) < 1000 Then
	    MsgBox "Year MUST include century i.e. put in 1994 NOT 94", 0, "WARNING Year incorrect"
	    Mdamyr.SetFocus
	    Mdamyr.SelStart = 0
	    Mdamyr.SelLength = Len(Mdamyr.Text)
	End If
    End If
End Sub

Sub Mdob_Change ()
    GLindchange% = -1
End Sub

Sub Mdob_LostFocus ()
    'if year is entered as 00 then assume 2000 otherwise accept
    'still have slight problem with 29/02/00 assumes 1900 for isdate function & reports as invalid
    'didn't worry about this

    If Trim(mDOB.Text) = "" Then
	mDOB.Text = "  /  /  "
	MsgBox "WARNING - Date can't be empty", 0, "Date Of Birth required"
	mDOB.SetFocus
	mDOB.SelStart = 0
	mDOB.SelLength = Len(mDOB.Text)
	Exit Sub
    End If
    
    'Fix potential leap year problem otherwise won't pass isdate() check
    If mDOB.Text = "29/2/00" Or mDOB.Text = "29/2/00" Then
	mDOB.Text = DateAdd("yyyy", 100, mDOB.Text)
    End If
    
    If IsDate(mDOB.Text) Then 'Must be Valid
	mDOB.Text = DateFix(mDOB.Text)
    Else
	MsgBox "WARNING - Date of Birth Invalid", 0, "Date Of Birth format incorrect"
	mDOB.SetFocus
	mDOB.SelStart = 0
	mDOB.SelLength = Len(mDOB.Text)
    End If
End Sub

Sub Mindbpa_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindbpa_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindbpb_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindbpb_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindbredr_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindbreds_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindrace_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindrace_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindref_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindsoh_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindsoh_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindsp1_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindsp1_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindsp2_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mindsp2_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mindstock_Click (Value As Integer)
    GLindchange% = -1
End Sub

Sub Mmark_Change ()
    GLindchange% = -1
End Sub

Sub Mmark_LostFocus ()
    If GLindchange% = -1 Then
	Mmark = UCase(Mmark)
    End If
End Sub

Sub Mname_Change ()
    GLindchange% = -1
End Sub

Sub Mnotes_Change ()
    GLindchange% = -1
End Sub

Sub Mnotes_GotFocus ()
    CsvListSires.Visible = False
    CsvListDams.Visible = False
End Sub

Sub Mnotes_KeyPress (KeyAscii As Integer)
'remove enter key
    If KeyAscii = 13 Then
	KeyAscii = Asc(" ")
    End If
End Sub

Sub Mring_Change ()
    GLindchange% = -1
End Sub

Sub Msiremark_Change ()
    GLindchange% = -1
End Sub

Sub Msiremark_GotFocus ()
    CsvListDams.Visible = False
End Sub

Sub Msiremark_LostFocus ()
    If GLindchange% = -1 Then
	Msiremark = UCase(Msiremark)
    End If
End Sub

Sub Msirering_Change ()
    GLindchange% = -1
End Sub

Sub Msirering_GotFocus ()
    CsvListDams.Visible = False
End Sub

Sub Msireyr_Change ()
    GLindchange% = -1
End Sub

Sub Msireyr_GotFocus ()
    CsvListDams.Visible = False
End Sub

Sub Msireyr_LostFocus ()
    Msireyr = Trim(Msireyr)
    If Trim(Msireyr) <> "" Then
	If Val(Msireyr) < 1000 Then
	    MsgBox "Year MUST include century i.e. put in 1994 NOT 94", 0, "WARNING Year incorrect"
	    Msireyr.SetFocus
	    Msireyr.SelStart = 0
	    Msireyr.SelLength = Len(Msireyr.Text)
	End If
    End If
End Sub

Sub Myr_Change ()
    GLindchange% = -1
End Sub

Sub Myr_LostFocus ()
    If Trim(myr <> "") Then
	If Val(myr) < 1000 Then 'if numeric value < 4 chars
	    MsgBox "Year MUST include century i.e. put in 1994 NOT 94", 0, "WARNING Year incorrect"
	    myr.SetFocus
	    myr.SelStart = 0
	    myr.SelLength = Len(myr.Text)
	Else
	End If
    End If
End Sub

Sub nullfields ()
    myr.Text = ""
    Mmark.Text = ""
    Mring.Text = ""
    Mname.Text = ""
    CSCombo1.Value = ""
    CSCombo2.Value = ""
    CSCombo3.Value = ""
    CSCombo4.Value = ""
    CSCombo1.Text = ""
    CSCombo2.Text = ""
    CSCombo3.Text = ""
    CSCombo4.Text = ""
    CSCombo5.Text = ""
    Mbredby.Text = ""
    'Mbredfor.Text = ""
    Mbloodline.Text = ""
    mIndSOH.Value = 0
    mIndRace.Value = 0
    mIndStock.Value = 0
    mIndRef.Value = 0
    mIndBpa.Value = 0
    mIndBpb.Value = 0
    mIndSp1.Value = 0
    mIndSp2.Value = 0
    mIndBreds.Value = 0
    mIndBredr.Value = 0
    mIndBredr.Value = 0
    Mdamyr.Text = ""
    Mdammark.Text = ""
    Mdamring.Text = ""
    Msireyr.Text = ""
    Msiremark.Text = ""
    Msirering.Text = ""
    mDOB.Text = "  /  /  "
    Mnotes.Text = ""
    'Let Status = ""
End Sub

Function RecordChanged () As Integer
    
    Dim X As Integer
    
    X = 0
    
    'Cannot change key ie YR MARK or RING NO, don't worry about checking key

    If IsNull(TblMaster("name")) And Mname = "" Then
    Else
	If IsNull(TblMaster("name")) Then 'CrsDesc changed
	    X = X + 1
	Else
	    If Trim(Mname) <> Trim(TblMaster("name")) Then
		X = X + 1
	    End If
	End If
    End If
    If IsNull(TblMaster("bredby")) And Mbredby = "" Then
    Else
	If IsNull(TblMaster("bredby")) Then 'CrsDesc changed
	    X = X + 1
	Else
	    If Trim(Mbredby) <> Trim(TblMaster("bredby")) Then
		X = X + 1
	    End If
	End If
    End If

    If IsNull(TblMaster("bloodline")) And Mbloodline = "" Then
    Else
	If IsNull(TblMaster("bloodline")) Then 'CrsDesc changed
	    X = X + 1
	Else
	    If Trim(Mbloodline) <> Trim(TblMaster("bloodline")) Then
		X = X + 1
	    End If
	End If
    End If

'* below to be fixed

    If X <> 0 Then
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

    X = 0
    
    If myr <> "" Then
	X = X + 1
    End If
    
    If Mmark <> "" Then
	X = X + 1
    End If
    
    If Mring <> "" Then
	X = X + 1
    End If
    
    If Mname <> "" Then
	X = X + 1
    End If

    If CSCombo1.Value <> "" Then
	X = X + 1
    End If
    If CSCombo2.Value <> "" Then
	X = X + 1
    End If
    If CSCombo3.Value <> "" Then
	X = X + 1
    End If
    If CSCombo4.Value <> "" Then
	X = X + 1
    End If
    If CSCombo5.Value <> "" Then
	X = X + 1
    End If
    If Mbredby <> "" Then
	X = X + 1
    End If
    If Mbloodline <> "" Then
	X = X + 1
    End If
    If mDOB = "" Then
	X = X + 1
    End If
    If mIndRace <> 0 Then
	X = X + 1
    End If
    If mIndStock <> 0 Then
	X = X + 1
    End If
    If mIndSOH <> 0 Then
	X = X + 1
    End If
    If mIndRef <> 0 Then
	X = X + 1
    End If
    If mIndBpa <> 0 Then
	X = X + 1
    End If
    If mIndBpb <> 0 Then
	X = X + 1
    End If
    If mIndSp1 <> 0 Then
	X = X + 1
    End If
    If mIndSp2 <> 0 Then
	X = X + 1
    End If
    If Mdamyr <> "" Then
	X = X + 1
    End If
    If Mdammark <> "" Then
	X = X + 1
    End If
    If Mdamring <> "" Then
	X = X + 1
    End If
    If Msireyr <> "" Then
	X = X + 1
    End If
    If Msiremark <> "" Then
	X = X + 1
    End If
    If Msirering <> "" Then
	X = X + 1
    End If
    If Mnotes <> "" Then
	X = X + 1
    End If
    If mDOB <> "" Then
	X = X + 1
    End If

    If X <> 0 Then
	RecordChangedInAdd = -1
    Else
	RecordChangedInAdd = 0
    End If

End Function

Sub ScrToTbl ()
    ' doesnt allow for bred for stock or race
    If Task.Text = "A" Or Task.Text = "C" Then 'Only allow key change if new record
	TblMaster("yr") = Val(myr.Text)
	TblMaster("mark") = Trim(Mmark.Text)
	TblMaster("ringno") = Trim(Mring.Text)
    End If
    TblMaster("birdname") = Trim(Mname.Text)
    TblMaster("bredby") = Trim(Mbredby.Text)
    'TblMaster("bredfor") = Trim(Mbredfor.Text)
    TblMaster("bloodline") = Trim(Mbloodline.Text)
    TblMaster("indsoh") = mIndSOH.Value
    TblMaster("indrace") = mIndRace.Value
    TblMaster("indstock") = mIndStock.Value
    TblMaster("indref") = mIndRef.Value
    TblMaster("indbpa") = mIndBpa.Value
    TblMaster("indbpb") = mIndBpb.Value
    TblMaster("indsp1") = mIndSp1.Value
    TblMaster("indsp2") = mIndSp2.Value
    TblMaster("notes") = Mnotes.Text

    If Val(Msireyr.Text) = 0 Then
	TblMaster("sireyr") = ""
    Else
	TblMaster("sireyr") = Val(Msireyr.Text)
    End If
    TblMaster("siremark") = Trim(Msiremark.Text)
    TblMaster("sireringno") = Trim(Msirering.Text)
    
    If Val(Mdamyr.Text) = 0 Then
	TblMaster("damyr") = ""
    Else
	TblMaster("damyr") = Val(Mdamyr.Text)
    End If
    TblMaster("dammark") = Trim(Mdammark.Text)
    TblMaster("damringno") = Trim(Mdamring.Text)
    TblMaster("ddob") = DateValue(mDOB.Text)
    TblMaster("dob_li") = DateValue(Trim(mDOB.Text))
    TblMaster("codesex") = CSCombo1.Value
    TblMaster("codecolour") = CSCombo3.Value
    TblMaster("codecage") = CSCombo2.Value
    TblMaster("codeeye") = CSCombo4.Value
    TblMaster("codebline") = CSCombo5.Value
    If mIndBreds.Value Then
	TblMaster("indbredforstock") = True
    Else
	TblMaster("indbredforstock") = False
    End If
    TblMaster.Update
End Sub

Sub SetAddMode ()
    ButtDelete.Visible = False
    ButtCancel.Visible = True
    ButtSires.Visible = True
    ButtDams.Visible = True

    'Frame3D3.Enabled = True
    myr.Enabled = True
    Mmark.Enabled = True
    Mring.Enabled = True
    Mname.Enabled = True
    Mnotes.Enabled = True
    mDOB.Enabled = True
    Mbredby.Enabled = True
    Mbloodline.Enabled = True
    'Mbredfor.Enabled = True

    CSCombo1.Enabled = True
    CSCombo2.Enabled = True
    CSCombo3.Enabled = True
    CSCombo4.Enabled = True
    CSCombo5.Enabled = True
    
    CsvListSires.Enabled = True
    CsvListSires.Visible = False
    CsvListDams.Enabled = True
    CsvListDams.Visible = False

    mIndBreds.Enabled = True
    mIndBredr.Enabled = True
    mIndSOH.Enabled = True
    mIndRace.Enabled = True
    mIndStock.Enabled = True
    mIndRef.Enabled = True
    mIndBpa.Enabled = True

    mIndBpb.Enabled = True

    mIndSp1.Enabled = True
    mIndSp2.Enabled = True

    myr.Text = Year(Now) - 1900
    Mmark.Text = GLMark
    Mring.Text = ""
    Mname.Text = ""
    Mnotes.Text = ""
    mDOB.Text = Date
    Mbredby.Text = GLOwner
    Mbloodline.Text = GLOwner
    'Mbredfor.Text = GLOwner

    CSCombo1.Text = "Unsexed"
    CSCombo1.Value = "U"
    CSCombo2.Text = ""
    CSCombo3.Text = ""
    CSCombo4.Text = "" 'Eyesign
    CSCombo5.Text = "" 'Bloodline

    mIndBredr.Value = True  'Default Race bird
    mIndBreds.Value = False

    mIndSOH.Value = True   'Default Current bird
    mIndRace.Value = True   'Default Race bird
    mIndRef.Value = False
    mIndStock.Value = False
    mIndSp1.Value = False

    mIndSp2.Value = False

    mIndBpa.Value = False
    mIndBpb.Value = False

    Mdamyr = 0
    Mdammark = ""
    Mdamring = ""
    Msireyr = 0
    Msiremark = ""
    Msirering = ""

    myr.SetFocus

End Sub

Sub SetEnqMode ()
'    ButtDelete.Visible = False
'    ButtCancel.Visible = False
'    ButtSires.Visible = False
'    ButtDams.Visible = False

    'Frame3D3.Enabled = False
'    myr.Enabled = False
'    Mmark.Enabled = False
'    Mring.Enabled = False
'    Mname.Enabled = False
'    Mnotes.Enabled = False
'    Mbloodline.Enabled = False
'    Mbredby.Enabled = False
    'Mbredfor.Enabled = False

'    Mdamyr.Enabled = False
'   Mdammark.Enabled = False
'    Mdamring.Enabled = False
'    Msireyr.Enabled = False
'    Msiremark.Enabled = False
'    Msirering.Enabled = False
    

'    mDOB.Enabled = False

'    CSCombo1.Enabled = False
'    CSCombo2.Enabled = False
'    CSCombo3.Enabled = False
'    CSCombo4.Enabled = False
'    CSCombo5.Enabled = False
    
'    mIndBredr.Enabled = False
'    mIndBreds.Enabled = False
'    mIndSOH.Enabled = False
'    mIndRace.Enabled = False
'    mIndRef.Enabled = False
'    mIndStock.Enabled = False
'    mIndSp1.Enabled = False

'    mIndSp2.Enabled = False

'    mIndBpa.Enabled = False
'    mIndBpb.Enabled = False

'    ButtOK.SetFocus

End Sub

Sub SetUpdMode ()
    ButtDelete.Visible = True
    ButtCancel.Visible = True
    ButtSires.Visible = True
    ButtDams.Visible = True

    'Frame3D3.Enabled = True
    myr.Enabled = False
    Mmark.Enabled = False
    Mring.Enabled = False
    Mname.Enabled = True
    Mnotes.Enabled = True
    mDOB.Enabled = True
    Mbredby.Enabled = True
    Mbloodline.Enabled = True
    'Mbredfor.Enabled = True

    Mdamyr.Enabled = True
    Mdammark.Enabled = True
    Mdamring.Enabled = True
    Msireyr.Enabled = True
    Msiremark.Enabled = True
    Msirering.Enabled = True

    CSCombo1.Enabled = True
    CSCombo2.Enabled = True
    CSCombo3.Enabled = True
    CSCombo4.Enabled = True
    CSCombo5.Enabled = True
    mIndBredr.Enabled = True
    mIndBreds.Enabled = True
    mIndSOH.Enabled = True
    mIndRef.Enabled = True
    mIndRace.Enabled = True
    mIndStock.Enabled = True
    mIndSp1.Enabled = True

    mIndSp2.Enabled = True

    mIndBpa.Enabled = True
    mIndBpb.Enabled = True

    Mname.SetFocus

End Sub

Sub TblToScr ()
    If IsNull(TblMaster("yr")) Then
	myr.Text = ""
    Else
	myr.Text = Format$(TblMaster("yr"), "")
    End If

    If IsNull(TblMaster("mark")) Then
	Mmark.Text = ""
    Else
	Mmark.Text = TblMaster("mark")
    End If

    If IsNull(TblMaster("ringno")) Then
	Mring.Text = ""
    Else
	Mring.Text = TblMaster("ringno")
    End If
    If IsNull(TblMaster("birdname")) Then
	Mname.Text = ""
    Else
	Mname.Text = TblMaster("birdname")
    End If
    If IsNull(TblMaster("bredby")) Then
	Mbredby.Text = ""
    Else
	Mbredby.Text = TblMaster("bredby")
    End If
    'If IsNull(TblMaster("bredfor")) Then
    '    Mbredfor.Text = ""
    'Else
    '    Mbredfor.Text = TblMaster("bredfor")
    'End If
    If IsNull(TblMaster("bloodline")) Then
	Mbloodline.Text = ""
    Else
	Mbloodline.Text = TblMaster("bloodline")
    End If
    If IsNull(TblMaster("notes")) Then
	Mnotes.Text = ""
    Else
	Mnotes.Text = TblMaster("notes")
    End If
    If IsNull(TblMaster("ddob")) Then
	mDOB.Text = "  /  /  " '0
    Else
	mDOB.Text = Format(TblMaster("ddob"), "ddddd")
    End If
    If IsNull(TblMaster("sireyr")) Then
	Msireyr.Text = ""
    Else
	Msireyr.Text = Format$(TblMaster("sireyr"), "")
    End If
    If IsNull(TblMaster("siremark")) Then
	Msiremark.Text = ""
    Else
	Msiremark.Text = TblMaster("siremark")
    End If
    If IsNull(TblMaster("sireringno")) Then
	Msirering.Text = ""
    Else
	Msirering.Text = TblMaster("sireringno")
    End If
    If IsNull(TblMaster("damyr")) Then
	Mdamyr.Text = ""
    Else
	Mdamyr.Text = Format$(TblMaster("damyr"), "")
    End If
    If IsNull(TblMaster("dammark")) Then
	Mdammark.Text = ""
    Else
	Mdammark.Text = TblMaster("dammark")
    End If
    If IsNull(TblMaster("damringno")) Then
	Mdamring.Text = ""
    Else
	Mdamring.Text = TblMaster("damringno")
    End If

    If IsNull(TblMaster("indsoh")) Then
	mIndSOH.Value = 0
    Else
	mIndSOH.Value = TblMaster("indsoh")
    End If
    If IsNull(TblMaster("indstock")) Then
	mIndStock.Value = 0
    Else
	mIndStock.Value = TblMaster("indstock")
    End If
    If IsNull(TblMaster("indrace")) Then
	mIndRace.Value = 0
    Else
	mIndRace.Value = TblMaster("indrace")
    End If
    If IsNull(TblMaster("indref")) Then
	mIndRef.Value = 0
    Else
	mIndRef.Value = TblMaster("indref")
    End If

    If IsNull(TblMaster("indbpa")) Then
	mIndBpa.Value = 0
    Else
	mIndBpa.Value = TblMaster("indbpa")
    End If
    If IsNull(TblMaster("indbpb")) Then
	mIndBpb.Value = 0
    Else
	mIndBpb.Value = TblMaster("indbpb")
    End If
    If IsNull(TblMaster("indsp1")) Then
	mIndSp1.Value = 0
    Else
	mIndSp1.Value = TblMaster("indsp1")
    End If
    If IsNull(TblMaster("indsp2")) Then
	mIndSp2.Value = 0
    Else
	mIndSp2.Value = TblMaster("indsp2")
    End If
    If IsNull(TblMaster("indbredforstock")) Then
	mIndBredr.Value = 1
    ElseIf TblMaster("indbredforstock") Then
	mIndBreds.Value = 1
    Else
	mIndBredr.Value = 1
    End If
    If IsNull(TblMaster("codecolour")) Then
	'CSCombo3.Text = "NA"
	CSCombo3.Value = "NA"
    Else
	CSCombo3.Value = TblMaster("codecolour")
    End If
    If IsNull(TblMaster("codesex")) Then
	'CSCombo1.Text = "Unsexed"
	CSCombo1.Value = "U"
    Else
	CSCombo1.Value = TblMaster("codesex")
    End If
    If IsNull(TblMaster("codeeye")) Then
	CSCombo4.Value = "NA"
    Else
	CSCombo4.Value = TblMaster("codeeye")
    End If
    If IsNull(TblMaster("codebline")) Then
	CSCombo5.Value = "NA"
    Else
	CSCombo5.Value = TblMaster("codebline")
    End If
    If IsNull(TblMaster("codecage")) Then
	CSCombo2.Value = ""
    Else
	CSCombo2.Value = TblMaster("codecage")
    End If

End Sub

