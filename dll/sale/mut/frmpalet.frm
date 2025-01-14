VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpalet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mutasi Lot/Palet"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton optmut 
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   60
      Width           =   2505
      _Version        =   851970
      _ExtentX        =   4419
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Laporan Mutasi Lot/Palet"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Value           =   -1  'True
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1275
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpalet.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   750
      TabIndex        =   1
      Top             =   825
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   134807553
      CurrentDate     =   42039
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   825
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   134807553
      CurrentDate     =   42039
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   75
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   2205
      TabIndex        =   5
      Top             =   1275
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpalet.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton optwip 
      Height          =   255
      Left            =   210
      TabIndex        =   7
      Top             =   375
      Width           =   2505
      _Version        =   851970
      _ExtentX        =   4419
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Laporan WIP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "s.d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2235
      TabIndex        =   4
      Top             =   885
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dari"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   885
      Width           =   465
   End
End
Attribute VB_Name = "frmpalet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub viewpalet()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport
    If optmut.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_daftarlot_gudang)"
        Crystal.ReportFileName = AppPath & "\reports\sale\mut\cetak_lotgudang.rpt"
    ElseIf optwip.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_daftarlot_wip)"
        Crystal.ReportFileName = AppPath & "\reports\sale\mut\cetak_wip.rpt"
    End If
    Crystal.ParameterFields(0) = "@tgl1;" & Format(Date1, "yyyy/MM/dd") & ";true"
    Crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
    Crystal.ParameterFields(2) = "@user;" & nmuser & ";True"

    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdview_Click()
    If Date1 > date2 Then
        MsgBox "Batas tanggal tidak benar..!", vbCritical, AppName
        Exit Sub
    End If
    viewpalet
End Sub

Private Sub Form_Load()
    Date1 = Date
    date2 = Date
End Sub
