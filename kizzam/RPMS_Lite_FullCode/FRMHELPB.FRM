VERSION 2.00
Begin Form frmHelpBU 
   Caption         =   "Help - BACKUPS"
   ClientHeight    =   6480
   ClientLeft      =   435
   ClientTop       =   420
   ClientWidth     =   8520
   Height          =   6885
   Left            =   375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8520
   Top             =   75
   Width           =   8640
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   5880
         Width           =   1575
      End
      Begin Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "If your database grows larger than 1.44Mb (1 disk) you will need to use you operating system utilities to backup the database i.e., file 'rp.mdb'. Otherwise from the MAIN Menu select FILE then Backup."
         Height          =   855
         Left            =   3240
         TabIndex        =   9
         Top             =   4920
         Width           =   5055
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "How Do I Backup?:"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   4920
         Width           =   2055
      End
      Begin Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Backups are essential to the security of your data i.e., the information that you enter into this system. As your database grows so its value to you increases. From this information you can draw knowledge that can give you the edge in your breeding and racing of your pigeons. As well as providing excellent record keeping which we all know is essential in racing pigeons."
         Height          =   975
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   7815
      End
      Begin Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "At least three (3) different Backup sets. Diskettes can also fail so don't rely on 1 backup set. I recommend keeping about 5 disks label them Set 1 to Set 5. Each time you backup use the next set when you get to set 5 then start again at set 1. It is also a good idea to keep a backup from say the end of year."
         Height          =   1215
         Left            =   3240
         TabIndex        =   6
         Top             =   3600
         Width           =   4935
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "How many BACKUPS should I keep?:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   2895
      End
      Begin Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "You see the information that you key in is stored in a file known as a database. The database is stored on your system. If this file is deleted, becomes corrupted or your PC fails for some reason then you may have to go back and restore your last good BACKUP to retrieve your data, then rekey anything since that BACKUP."
         Height          =   1335
         Left            =   3240
         TabIndex        =   4
         Top             =   2160
         Width           =   4935
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Why are BACKUP important?:"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "After every session where you update your pigeon or race records, depending on the amount of changes you make. Or a least after every 2nd session."
         Height          =   735
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   4935
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "When should backups be done?:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

