VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpermint_report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Permintaan Barang"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnolot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1215
      Width           =   2430
   End
   Begin XtremeSuiteControls.RadioButton optlot 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
      _Version        =   851970
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   " Permintaan By Lot"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
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
      MICON           =   "frmpermint_report.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   3960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton opttgl 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2055
      _Version        =   851970
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   " Permintaan By Date"
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
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
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
      MICON           =   "frmpermint_report.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GB1 
      Height          =   810
      Left            =   240
      TabIndex        =   5
      Top             =   1700
      Width           =   4260
      _Version        =   851970
      _ExtentX        =   7514
      _ExtentY        =   1429
      _StockProps     =   79
      BackColor       =   -2147483634
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   705
         TabIndex        =   6
         Top             =   315
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
         Format          =   135069697
         CurrentDate     =   42039
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   315
         Left            =   2535
         TabIndex        =   7
         Top             =   315
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
         Format          =   135069697
         CurrentDate     =   42039
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
         Left            =   2190
         TabIndex        =   9
         Top             =   375
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
         Left            =   150
         TabIndex        =   8
         Top             =   375
         Width           =   465
      End
   End
   Begin Chameleon.chameleonButton cmdcari 
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cari"
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
      MICON           =   "frmpermint_report.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "frmpermint_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcari_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If optlot.Value = True Then
        If txtnolot = "" Then Exit Sub
        
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(am_cetake_report)"
        crystal.ReportFileName = AppPath & "\reports\produksi\take_pack_report.rpt"
        crystal.ParameterFields(0) = "@nolot;" & txtnolot.text & ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
    ElseIf opttgl.Value = True Then
        If date1 > date2 Then
            MsgBox "Batas tanggal tidak benar..!", vbCritical, AppName
            Exit Sub
        End If
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(am_cetake_repdate)"
        crystal.ReportFileName = AppPath & "\reports\produksi\take_pack_repdate.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(2) = "@username;" & nmuser & ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
    End If
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Private Sub optlot_Click()
    If optlot.Value = True Then
        GB1.Enabled = False
        cmdcari.Enabled = True
    End If
End Sub

Private Sub opttgl_Click()
    If opttgl.Value = True Then
        GB1.Enabled = True
        cmdcari.Enabled = False
        txtnolot = ""
    End If
End Sub
