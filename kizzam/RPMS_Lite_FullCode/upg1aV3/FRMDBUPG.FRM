VERSION 2.00
Begin Form frmDBupg 
   Caption         =   "PHASE 2 - Upgrade Database"
   ClientHeight    =   2820
   ClientLeft      =   1560
   ClientTop       =   2430
   ClientWidth     =   4590
   Height          =   3225
   Left            =   1500
   LinkTopic       =   "Form3"
   ScaleHeight     =   2820
   ScaleWidth      =   4590
   Top             =   2085
   Width           =   4710
   Begin Timer Timer1 
      Left            =   2040
      Top             =   2160
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Phase 2 - Upgrading Database - Please WAIT ......"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Option Explicit

Sub Form_Load ()
    Dim x As Integer
    Dim FldName As String
    Dim FldValue As String
    Dim RetVal As Integer
    Dim CmdLine As String
    Dim result As Integer
    '
    FldName$ = "RsRaceFormType"
    If Not FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_INTEGER, Null)
    End If
    FldName$ = "DBVersion"
    If Not FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_TEXT, 5)
    End If
    FldName$ = "LastBackup"
    If Not FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_DATE, Null)
    End If
    FldName$ = "Copyright"
    If Not FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_TEXT, 32)
    End If
    FldName$ = "DistanceType"
    If Not FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_TEXT, 1)
    End If
    FldName$ = "OwnerFed"
    If FieldExists("loft", FldName$) Then
        result% = AppendDeleteField("loft", "DELETE", FldName$, Null, Null)
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_TEXT, 45)
    Else
        result% = AppendDeleteField("loft", "APPEND", FldName$, DB_TEXT, 45)
    End If
    '
'MedCodes
    FldName$ = "Dosage"
    If Not FieldExists("MedCodes", FldName$) Then
        result% = AppendDeleteField("MedCodes", "APPEND", FldName$, DB_SINGLE, Null)
    End If
    FldName$ = "Status"
    If Not FieldExists("MedCodes", FldName$) Then
        result% = AppendDeleteField("MedCodes", "APPEND", FldName$, DB_TEXT, 1)
    End If
'End
        '
'Race
    FldName$ = "Distance"
    If Not FieldExists("Race", FldName$) Then
        result% = AppendDeleteField("Race", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
'End
'
'RaceTrans
    FldName$ = "Distance"
    If Not FieldExists("RaceTrans", FldName$) Then
        result% = AppendDeleteField("RaceTrans", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "VelocityCalc"
    If Not FieldExists("RaceTrans", FldName$) Then
        result% = AppendDeleteField("RaceTrans", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
'End
        '
'Points
    Set td = db.TableDefs("points")
    FldName$ = "DistCalcGCM_TDPF"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistCalcGCM_UK"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistCalcGCM_Intl"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistCalcGEOD"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "MyPoint"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_TEXT, 1)
    End If
    FldName$ = "Club"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_TEXT, 16)
    End If
    FldName$ = "DistCalcGCM_TDPFyds"
    If Not FieldExists("points", FldName$) Then
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    '
    FldName$ = "DistCalcGCM_UKyds"
    If FieldExists("points", FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, DB_DOUBLE)
    Else
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistCalcGCM_Intlyds"
    If FieldExists("points", FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, DB_DOUBLE)
    Else
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistCalcGEODyds"
    If FieldExists("points", FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, DB_DOUBLE)
    Else
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_DOUBLE, Null)
    End If
    FldName$ = "DistType"
    If FieldExists("points", FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, DB_TEXT, 1)
    Else
        result% = AppendDeleteField("points", "APPEND", FldName$, DB_TEXT, 1)
    End If
'end Points
    '
    db.Close
    
'Tables (new)
    If Not TableExists("Matings") Then
        result% = CreateTable(DBFullFileName$, "Matings")
        Set db = OpenDatabase(DBFullFileName$)
        'Set td = db.TableDefs("Matings")
        result% = AppendDeleteField("Matings", "APPEND", "Code", DB_TEXT, 5)
        result% = AppendDeleteField("Matings", "APPEND", "Rating", DB_LONG, Null)
        result% = AppendDeleteField("Matings", "APPEND", "RatingPerf", DB_LONG, Null)
        result% = AppendDeleteField("Matings", "APPEND", "RatingBloodLine", DB_LONG, Null)
        result% = AppendDeleteField("Matings", "APPEND", "RatingCustom", DB_LONG, Null)
        result% = AppendDeleteField("Matings", "APPEND", "StsSystem", DB_TEXT, 1)
        result% = AppendDeleteField("Matings", "APPEND", "StsMating", DB_TEXT, 1)
        result% = AppendDeleteField("Matings", "APPEND", "StsCurrent", DB_TEXT, 1)
        result% = AppendDeleteField("Matings", "APPEND", "StsCustom", DB_TEXT, 1)
        result% = AppendDeleteField("Matings", "APPEND", "SireYr", DB_INTEGER, Null)
        result% = AppendDeleteField("Matings", "APPEND", "SireMark", DB_TEXT, 12)
        result% = AppendDeleteField("Matings", "APPEND", "SireRingNo", DB_TEXT, 8)
        result% = AppendDeleteField("Matings", "APPEND", "DamYr", DB_INTEGER, Null)
        result% = AppendDeleteField("Matings", "APPEND", "DamMark", DB_TEXT, 12)
        result% = AppendDeleteField("Matings", "APPEND", "DamRingNo", DB_TEXT, 8)
        result% = AppendDeleteField("Matings", "APPEND", "Notes", DB_MEMO, Null)
        'Result% = AppendDeleteField("Matings", "APPEND", "Picture", dbOLEObject)
        db.Close
    End If
    '
    'Upgrade Bloodline
    'Set td = db.TableDefs("Bloodline")
    If Not FieldExists("Bloodline", "RatingBreeding") Then
        Set db = OpenDatabase(DBFullFileName$)
        result% = AppendDeleteField("Bloodline", "APPEND", "Rating", DB_LONG, Null)
        result% = AppendDeleteField("Bloodline", "APPEND", "RatingBreeding", DB_LONG, Null)
        result% = AppendDeleteField("Bloodline", "APPEND", "RatingPerf", DB_LONG, Null)
        result% = AppendDeleteField("Bloodline", "APPEND", "RatingCustom", DB_LONG, Null)
        db.Close
    End If
    '
    'Upgrade Master
    'Set td = db.TableDefs("Master")
    If Not FieldExists("Master", "Rating") Then
        Set db = OpenDatabase(DBFullFileName$)
        result% = AppendDeleteField("Master", "APPEND", "Rating", DB_LONG, Null)
        result% = AppendDeleteField("Master", "APPEND", "RatingPerf", DB_LONG, Null)
        result% = AppendDeleteField("Master", "APPEND", "RatingBloodLine", DB_LONG, Null)
        result% = AppendDeleteField("Master", "APPEND", "RatingCustom", DB_LONG, Null)
        result% = AppendDeleteField("Master", "APPEND", "RatingBreeder", DB_LONG, Null)
        db.Close
    End If
'
'Create Schedule DB
    If Not TableExists("Schedule") Then
        result% = CreateTable(DBFullFileName$, "Schedule")
        Set db = OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Schedule")
        result% = AppendDeleteField("Schedule", "APPEND", "SchCode", DB_TEXT, 5)
        result% = AppendDeleteField("Schedule", "APPEND", "SchDesc", DB_TEXT, 36)
        result% = AppendDeleteField("Schedule", "APPEND", "SchType", DB_TEXT, 12)
        result% = AppendDeleteField("Schedule", "APPEND", "SchDate", DB_DATE, Null)
        result% = AppendDeleteField("Schedule", "APPEND", "SchSeries", DB_LONG, Null)
        result% = AppendDeleteField("Schedule", "APPEND", "SchEvent", DB_LONG, Null)
        result% = AppendDeleteField("Schedule", "APPEND", "SchSystem", DB_TEXT, 1)
        result% = AppendDeleteField("Schedule", "APPEND", "SchStatus", DB_TEXT, 1)
        result% = AppendDeleteField("Schedule", "APPEND", "SchClub", DB_TEXT, 16)
        db.Close
    End If
    '
'Create Drugs DB
    If Not TableExists("Drugs") Then
        result% = CreateTable(DBFullFileName$, "Drugs")
        Set db = OpenDatabase(DBFullFileName$)
        'Set td = db.TableDefs("Drugs")
        result% = AppendDeleteField("Drugs", "APPEND", "DrugCode", DB_TEXT, 3)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugDesc", DB_TEXT, 32)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugBrand", DB_TEXT, 24)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugGeneric", DB_TEXT, 24)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugDosageRate", DB_TEXT, 24)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugEffectivePeriod", DB_INTEGER, Null)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugUse", DB_TEXT, 32)
        result% = AppendDeleteField("Drugs", "APPEND", "DrugSupplier", DB_TEXT, 50)
        db.Close
    End If
    '
'Create Event DB
    If Not TableExists("Event") Then
        result% = CreateTable(DBFullFileName$, "Event")
        Set db = OpenDatabase(DBFullFileName$)
        'Set "Event" = db.TableDefs("Event")
        result% = AppendDeleteField("Event", "APPEND", "EventCode", DB_TEXT, 3)
        result% = AppendDeleteField("Event", "APPEND", "Eventdesc", DB_TEXT, 32)
        result% = AppendDeleteField("Event", "APPEND", "EventStatus", DB_TEXT, 1)
        result% = AppendDeleteField("Event", "APPEND", "EventDate", DB_DATE, Null)
        result% = AppendDeleteField("Event", "APPEND", "EventType", DB_TEXT, 3)
        db.Close
    End If
    '
'Create Notes DB
    If Not TableExists("Notes") Then
        result% = CreateTable(DBFullFileName$, "Notes")
        Set db = OpenDatabase(DBFullFileName$)
        'Set td = db.TableDefs("Notes")
        result% = AppendDeleteField("Notes", "APPEND", "NoteCode", DB_DOUBLE, Null)
        result% = AppendDeleteField("Notes", "APPEND", "NoteDate", DB_DATE, Null)
        result% = AppendDeleteField("Notes", "APPEND", "NoteMemo", DB_TEXT, 120)
        result% = AppendDeleteField("Notes", "APPEND", "NoteStatus", DB_TEXT, 1)
        db.Close
    End If
    '
'Create Treatment DB
    If Not TableExists("Treatment") Then
        result% = CreateTable(DBFullFileName$, "Treatment")
        Set db = OpenDatabase(DBFullFileName$)
        'Set td = db.TableDefs("Treatment")
        result% = AppendDeleteField("Treatment", "APPEND", "DrugCode", DB_TEXT, 3) 'Key
        result% = AppendDeleteField("Treatment", "APPEND", "NoteCode", DB_DOUBLE, Null)
        result% = AppendDeleteField("Treatment", "APPEND", "TreatmentDosage", DB_TEXT, 24)
        result% = AppendDeleteField("Treatment", "APPEND", "TreatmentDate", DB_DATE, Null)
        result% = AppendDeleteField("Treatment", "APPEND", "EventType", DB_TEXT, 3)
        db.Close
    End If
    
    'Fix Y2k Problem in Race Trans & Race tables
    Fix_relDate
    '
End Sub

'
Sub Timer1_Timer ()
    'If Timer.Interval > 20000 Then
        Unload Me
    'End If
End Sub

