VERSION 2.00
Begin Form FMmain 
   Caption         =   "The Racing Pigeon Management System"
   ClientHeight    =   3975
   ClientLeft      =   1380
   ClientTop       =   1770
   ClientWidth     =   7800
   Height          =   4665
   Left            =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7800
   Top             =   1140
   Width           =   7920
   Begin PictureBox Picture1 
      Height          =   2655
      Left            =   360
      Picture         =   MAIN.FRX:0000
      ScaleHeight     =   2625
      ScaleWidth      =   1905
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Caption         =   "Panel3D2"
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Caption         =   "Panel3D1"
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   7815
   End
   Begin Menu mnuLoft 
      Caption         =   "&Loft Information"
      Begin Menu mnuOwner 
         Caption         =   "&Owner"
      End
      Begin Menu mnuRacePoints 
         Caption         =   "&Race Points"
      End
      Begin Menu mnuColour 
         Caption         =   "&Colour"
      End
      Begin Menu mnuEye 
         Caption         =   "&Eye Sign"
      End
      Begin Menu mnuBage 
         Caption         =   "C&ages"
      End
      Begin Menu mnuNotes 
         Caption         =   "&Notes"
         Enabled         =   0   'False
      End
      Begin Menu test 
         Caption         =   "&Test"
         Visible         =   0   'False
      End
   End
   Begin Menu mnuBreed 
      Caption         =   "&Breeding"
      Begin Menu mnuRegister 
         Caption         =   "&Register"
      End
      Begin Menu mnuMatings 
         Caption         =   "&Matings"
      End
      Begin Menu mnuBReports 
         Caption         =   "&Reports"
         Begin Menu mnuBRRegister 
            Caption         =   "&Register"
         End
         Begin Menu mnuBRPedg 
            Caption         =   "&Pedigree"
         End
      End
   End
   Begin Menu mnuRace 
      Caption         =   "&Racing"
      Begin Menu mnuRRace 
         Caption         =   "&Race"
      End
      Begin Menu mnuRReport 
         Caption         =   "R&eports"
         Begin Menu mnuRERaceF 
            Caption         =   "&Races Flown"
         End
         Begin Menu mnuREPFlown 
            Caption         =   "&Points Flown by Bird"
         End
         Begin Menu mnuRErsheet 
            Caption         =   "R&eprint Race Sheet"
         End
      End
   End
   Begin Menu mnuReports 
      Caption         =   "R&eports"
      Begin Menu mnuERacePoints 
         Caption         =   "&Race Points"
      End
      Begin Menu mnuEAll 
         Caption         =   "&Colours/EyeSign/Cages"
      End
      Begin Menu mnuDistances 
         Caption         =   "&Distances"
      End
   End
   Begin Menu mnuHelp 
      Caption         =   "&Help"
      Begin Menu mnuHHelp 
         Caption         =   "&Help"
      End
      Begin Menu mnuHOrder 
         Caption         =   "&Order Software"
      End
      Begin Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Option Explicit

Sub mnuBage_Click ()
    Load FMcage
    FMcage.Show
End Sub

Sub mnuColour_Click ()
    Load FMcolour
    FMcolour.Show
End Sub

Sub mnuExit_Click ()
    Unload FMmain
End Sub

Sub mnuEye_Click ()
    Load FMeyesign
    FMeyesign.Show
End Sub

Sub mnuOwner_Click ()
    Load FMowner
    FMowner.Show 1
End Sub

Sub mnuRacePoints_Click ()
 Load FMracepoints
 FMracepoints.Show 1
End Sub

