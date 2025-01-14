VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlaplot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Pemakaian Kemasan per Lot"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton opttgl 
      Height          =   225
      Left            =   195
      TabIndex        =   1
      Top             =   165
      Width           =   2625
      _Version        =   851970
      _ExtentX        =   4630
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   " Pemakaian Kemasan By Date"
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
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   3255
      TabIndex        =   0
      Top             =   2775
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
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
   Begin XtremeSuiteControls.RadioButton optlot 
      Height          =   270
      Left            =   195
      TabIndex        =   2
      Top             =   825
      Width           =   2550
      _Version        =   851970
      _ExtentX        =   4498
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   " Pemakaian Kemasan perLot"
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
   Begin XtremeSuiteControls.GroupBox GB1 
      Height          =   810
      Left            =   165
      TabIndex        =   3
      Top             =   1920
      Width           =   4140
      _Version        =   851970
      _ExtentX        =   7302
      _ExtentY        =   1429
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   705
         TabIndex        =   4
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
         Format          =   134479873
         CurrentDate     =   42039
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   315
         Left            =   2535
         TabIndex        =   5
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
         Format          =   134479873
         CurrentDate     =   42039
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
         TabIndex        =   7
         Top             =   375
         Width           =   465
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
         TabIndex        =   6
         Top             =   375
         Width           =   315
      End
   End
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   465
      Left            =   2175
      TabIndex        =   11
      Top             =   2775
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "View"
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
   Begin Crystal.CrystalReport crystal 
      Left            =   90
      Top             =   2775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton optgudang 
      Height          =   225
      Left            =   195
      TabIndex        =   12
      Top             =   1170
      Width           =   3105
      _Version        =   851970
      _ExtentX        =   5477
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   " Mutasi Palet (Perolehan Barang Jadi)"
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
   Begin XtremeSuiteControls.GroupBox GB2 
      Height          =   810
      Left            =   165
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   4140
      _Version        =   851970
      _ExtentX        =   7302
      _ExtentY        =   1429
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtnolot 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Top             =   300
         Width           =   2895
      End
      Begin XtremeSuiteControls.PushButton cmdnolot 
         Height          =   300
         Left            =   105
         TabIndex        =   10
         Top             =   300
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "NO LOT :"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.RadioButton optByLot 
      Height          =   270
      Left            =   195
      TabIndex        =   13
      Top             =   480
      Width           =   2550
      _Version        =   851970
      _ExtentX        =   4498
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   " Pemakaian Kemasan By Lot"
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
   Begin XtremeSuiteControls.RadioButton optWIP 
      Height          =   225
      Left            =   195
      TabIndex        =   14
      Top             =   1515
      Width           =   1875
      _Version        =   851970
      _ExtentX        =   3307
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   " Laporan WIP"
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
   Begin VB.Image Image1 
      Height          =   555
      Left            =   2850
      Picture         =   "frmlaplot.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmlaplot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmdnolot_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk "
    carisql1 = carisql1 + "where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnolot_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If opttgl.Value = True Or optgudang.Value = True Then
        If date1 > date2 Then
            MsgBox "Batas tanggal tidak benar..!", vbCritical, AppName
            Exit Sub
        End If
    End If

    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    If optlot.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarlot)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_lot.rpt"
        crystal.ParameterFields(0) = "@nolot;" & txtnolot & ";true"
        
    ElseIf opttgl.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarlot_tgl)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_lottgl.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(2) = "@user;" & nmuser & ";True"
    
    ElseIf optByLot.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarlot_tgl)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_bylot_new.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(2) = "@user;" & nmuser & ";True"
    ElseIf optWIP.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarlot_wip)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_wip.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(2) = "@user;" & nmuser & ";True"
        
    Else
        crystal.DataFiles(0) = "Proc(am_daftarlot_gudang)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_lotgudang.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(2) = "@user;" & nmuser & ";True"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Private Sub optByLot_Click()
    GB1.Visible = True
    GB2.Visible = False
    txtnolot = ""
End Sub

Private Sub optgudang_Click()
    GB1.Visible = True
    GB2.Visible = False
    txtnolot = ""
End Sub

Private Sub optlot_Click()
    GB2.Visible = True
    GB1.Visible = False
    txtnolot.SetFocus
End Sub

Private Sub opttgl_Click()
    GB1.Visible = True
    GB2.Visible = False
    txtnolot = ""
End Sub

Private Sub optWIP_Click()
    GB1.Visible = True
    GB2.Visible = False
    txtnolot = ""
End Sub
