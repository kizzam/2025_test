VERSION 5.00
Begin VB.Form frmDBupg 
   Caption         =   "Upgrade Database"
   ClientHeight    =   3645
   ClientLeft      =   1305
   ClientTop       =   2220
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   360
      Top             =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Phase 2 - Please wait while Database is being upgraded.........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
End
Attribute VB_Name = "frmDBupg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim db As Database
    Dim td As TableDef
    Dim Fld As Field
    Dim FldName As String
    Dim x As Integer
    Dim FldValue As String
    Dim RetVal As Integer
    Dim CmdLine As String
    Dim result As Integer
    
    Screen.MousePointer = 11
    Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
    Set td = db.TableDefs("loft")
    
    FldName$ = "RsRaceFormType"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbInteger)
    End If
    FldName$ = "DBVersion"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 5)
    End If
    FldName$ = "LastBackup"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDate)
    End If
    FldName$ = "Copyright"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 32)
    End If
    FldName$ = "DistanceType"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 1)
    End If
    FldName$ = "OwnerFed"
    If FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "DELETE", FldName$)
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 45)
    Else
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 45)
    End If
    
'MedCodes
    Set td = db.TableDefs("MedCodes")
    FldName$ = "Dosage"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbSingle)
    End If
    FldName$ = "Status"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 1)
    End If
'End
        
'Race
    Set td = db.TableDefs("Race")
    FldName$ = "Distance"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
'End

'RaceTrans
    Set td = db.TableDefs("RaceTrans")
    FldName$ = "Distance"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "VelocityCalc"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
'End
        
'Points
    Set td = db.TableDefs("points")
    FldName$ = "DistCalcGCM_TDPF"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistCalcGCM_UK"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistCalcGCM_Intl"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistCalcGEOD"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "MyPoint"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 1)
    End If
    FldName$ = "Club"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 16)
    End If
    FldName$ = "DistCalcGCM_TDPFyds"
    If Not FieldExists(td, FldName$) Then
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    
    FldName$ = "DistCalcGCM_UKyds"
    If FieldExists(td, FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    Else
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistCalcGCM_Intlyds"
    If FieldExists(td, FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    Else
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistCalcGEODyds"
    If FieldExists(td, FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    Else
        result% = AppendDeleteField(td, "APPEND", FldName$, dbDouble)
    End If
    FldName$ = "DistType"
    If FieldExists(td, FldName$) Then
        'Result% = AppendDeleteField(td, "DELETE", FldName$)
        'Result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 1)
    Else
        result% = AppendDeleteField(td, "APPEND", FldName$, dbText, 1)
    End If
'end Points
    
    If Not TableExists(db, "Matings") Then
        result% = CreateTable(DBFullFileName$, "Matings")
        'Set Db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Matings")
        result% = AppendDeleteField(td, "APPEND", "Code", dbText, 5)
        result% = AppendDeleteField(td, "APPEND", "Rating", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingPerf", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingBloodLine", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingCustom", dbLong)
        result% = AppendDeleteField(td, "APPEND", "StsSystem", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "StsMating", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "StsCurrent", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "StsCustom", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "SireYr", dbInteger)
        result% = AppendDeleteField(td, "APPEND", "SireMark", dbText, 12)
        result% = AppendDeleteField(td, "APPEND", "SireRingNo", dbText, 8)
        result% = AppendDeleteField(td, "APPEND", "DamYr", dbInteger)
        result% = AppendDeleteField(td, "APPEND", "DamMark", dbText, 12)
        result% = AppendDeleteField(td, "APPEND", "DamRingNo", dbText, 8)
        result% = AppendDeleteField(td, "APPEND", "Notes", dbMemo)
        'Result% = AppendDeleteField(td, "APPEND", "Picture", dbOLEObject)
        db.Close
    End If
    
    'Upgrade Bloodline
    Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
    Set td = db.TableDefs("Bloodline")
    If Not FieldExists(td, "RatingBreeding") Then
        result% = AppendDeleteField(td, "APPEND", "Rating", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingBreeding", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingPerf", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingCustom", dbLong)
    End If
    db.Close
    
    'Upgrade Master
    Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
    Set td = db.TableDefs("Master")
    If Not FieldExists(td, "Rating") Then
        result% = AppendDeleteField(td, "APPEND", "Rating", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingPerf", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingBloodLine", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingCustom", dbLong)
        result% = AppendDeleteField(td, "APPEND", "RatingBreeder", dbLong)
    End If

'Create Schedule DB
    If Not TableExists(db, "Schedule") Then
        result% = CreateTable(DBFullFileName$, "Schedule")
        'Set Db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Schedule")
        result% = AppendDeleteField(td, "APPEND", "SchCode", dbText, 5)
        result% = AppendDeleteField(td, "APPEND", "SchDesc", dbText, 36)
        result% = AppendDeleteField(td, "APPEND", "SchType", dbText, 12)
        result% = AppendDeleteField(td, "APPEND", "SchDate", dbDate)
        result% = AppendDeleteField(td, "APPEND", "SchSeries", dbLong)
        result% = AppendDeleteField(td, "APPEND", "SchEvent", dbLong)
        result% = AppendDeleteField(td, "APPEND", "SchSystem", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "SchStatus", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "SchClub", dbText, 16)
        'db.Close
    End If
    
'Create Drugs DB
    If Not TableExists(db, "Drugs") Then
        result% = CreateTable(DBFullFileName$, "Drugs")
        'Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Drugs")
        result% = AppendDeleteField(td, "APPEND", "DrugCode", dbText, 3)
        result% = AppendDeleteField(td, "APPEND", "DrugDesc", dbText, 32)
        result% = AppendDeleteField(td, "APPEND", "DrugBrand", dbText, 24)
        result% = AppendDeleteField(td, "APPEND", "DrugGeneric", dbText, 24)
        result% = AppendDeleteField(td, "APPEND", "DrugDosageRate", dbText, 24)
        result% = AppendDeleteField(td, "APPEND", "DrugEffectivePeriod", dbInteger)
        result% = AppendDeleteField(td, "APPEND", "DrugUse", dbText, 32)
        result% = AppendDeleteField(td, "APPEND", "DrugSupplier", dbText, 50)
    End If
    
'Create Event DB
    If Not TableExists(db, "Event") Then
        result% = CreateTable(DBFullFileName$, "Event")
        'Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Event")
        result% = AppendDeleteField(td, "APPEND", "EventCode", dbText, 3)
        result% = AppendDeleteField(td, "APPEND", "EventDesc", dbText, 32)
        result% = AppendDeleteField(td, "APPEND", "EventStatus", dbText, 1)
        result% = AppendDeleteField(td, "APPEND", "EventDate", dbDate)
        result% = AppendDeleteField(td, "APPEND", "EventType", dbText, 3)
    End If
    
'Create Notes DB
    If Not TableExists(db, "Notes") Then
        result% = CreateTable(DBFullFileName$, "Notes")
        'Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Notes")
        result% = AppendDeleteField(td, "APPEND", "NoteCode", dbDouble)
        result% = AppendDeleteField(td, "APPEND", "NoteDate", dbDate)
        result% = AppendDeleteField(td, "APPEND", "NoteMemo", dbText, 120)
        result% = AppendDeleteField(td, "APPEND", "NoteStatus", dbText, 1)
    End If
    
'Create Treatment DB
    If Not TableExists(db, "Treatment") Then
        result% = CreateTable(DBFullFileName$, "Treatment")
        'Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        db.Close
        Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
        Set td = db.TableDefs("Treatment")
        result% = AppendDeleteField(td, "APPEND", "DrugCode", dbText, 3) 'Key
        result% = AppendDeleteField(td, "APPEND", "NoteCode", dbDouble)
        result% = AppendDeleteField(td, "APPEND", "TreatmentDosage", dbText, 24)
        result% = AppendDeleteField(td, "APPEND", "TreatmentDate", dbDate)
        result% = AppendDeleteField(td, "APPEND", "EventType", dbText, 3)
    End If
    
    db.Close
            
    Set td = Nothing
    Set db = Nothing
    
    'Fix Y2k Problem in Race Trans & Race tables
    Fix_relDate
    
End Sub

Private Sub Timer1_Timer()
    'If Timer.Interval > 20000 Then
        Unload Me
    'End If
End Sub
