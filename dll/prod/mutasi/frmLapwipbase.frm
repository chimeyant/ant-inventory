VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLapwipbase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Mutasi WIP Base"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkode2 
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
      Height          =   285
      Left            =   2085
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox txtproduk2 
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
      Height          =   285
      Left            =   2085
      TabIndex        =   11
      Top             =   1680
      Width           =   2430
   End
   Begin VB.TextBox txtkode 
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
      Height          =   285
      Left            =   2085
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox txtproduk 
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
      Height          =   285
      Left            =   2085
      TabIndex        =   8
      Top             =   1320
      Width           =   2430
   End
   Begin XtremeSuiteControls.RadioButton RadioButton2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "By Produk"
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
   Begin XtremeSuiteControls.RadioButton RadioButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
      _Version        =   851970
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "By Lot"
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
   Begin VB.TextBox txtnolot2 
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
      Left            =   2085
      TabIndex        =   2
      Top             =   720
      Width           =   2430
   End
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
      Left            =   2085
      TabIndex        =   1
      Top             =   360
      Width           =   2430
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2520
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
      MICON           =   "frmLapwipbase.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
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
      MICON           =   "frmLapwipbase.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlot 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "From Lot"
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
      MICON           =   "frmLapwipbase.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlot2 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "To Lot"
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
      MICON           =   "frmLapwipbase.frx":094E
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
      Left            =   0
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdproduk 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Base WIP"
      ENAB            =   0   'False
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
      MICON           =   "frmLapwipbase.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdproduk2 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Base WIP"
      ENAB            =   0   'False
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
      MICON           =   "frmLapwipbase.frx":0F82
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
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmLapwipbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdlot_Click()
    txtnolot = ""
    namatabel = "nolot base"

    carisql1 = "Select distinct a.nolot,a.kodebahan,b.NamaBarang From am_stoklot a"
    carisql1 = carisql1 + "  inner join am_apitemmst b on a.kodebahan = b.KodeBarang "
    frmsearch.Show vbModal
End Sub

Private Sub cmdlot_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdlot2_Click()
    txtnolot2 = ""
    namatabel = "nolot base"

    carisql1 = "Select distinct a.nolot,a.kodebahan,b.NamaBarang From am_stoklot a"
    carisql1 = carisql1 + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang "
    frmsearch.Show vbModal
End Sub

Private Sub cmdlot2_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdproduk_Click()
    txtproduk = ""
    txtkode = ""
    namatabel = "item base"

    carisql1 = "Select distinct a.kodebahan,b.NamaBarang,b.KodeSatuan From am_stoklot a"
    carisql1 = carisql1 + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang "
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduk_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    txtproduk = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdproduk2_Click()
    txtproduk2 = ""
    txtkode2 = ""
    namatabel = "item base"

    carisql1 = "Select distinct a.kodebahan,b.NamaBarang,b.KodeSatuan From am_stoklot a"
    carisql1 = carisql1 + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang "
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduk2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode2 = hasil
    txtproduk2 = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdview_Click()
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    If RadioButton1.Value = True Then
        crystal.DataFiles(0) = "Proc(am_lapmut_wipbase)"
        crystal.ReportFileName = AppPath & "\reports\produksi\mut\lapmutbasewip.rpt"
        crystal.ParameterFields(0) = "@nolot;" & txtnolot & ";true"
        crystal.ParameterFields(1) = "@nolot2;" & txtnolot2 & ";true"
    ElseIf RadioButton2.Value = True Then
        crystal.DataFiles(0) = "Proc(am_lapmut_kodebase)"
        crystal.ReportFileName = AppPath & "\reports\produksi\mut\lapmutbasecode.rpt"
        crystal.ParameterFields(0) = "@kode;" & txtkode & ";true"
        crystal.ParameterFields(1) = "@kode2;" & txtkode2 & ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub RadioButton1_Click()
    txtkode = ""
    txtkode2 = ""
    txtproduk = ""
    txtproduk2 = ""
    cmdproduk.Enabled = False
    cmdproduk2.Enabled = False
    cmdlot.Enabled = True
    cmdlot2.Enabled = True
End Sub

Private Sub RadioButton2_Click()
    txtnolot = ""
    txtnolot2 = ""
    cmdproduk.Enabled = True
    cmdproduk2.Enabled = True
    cmdlot.Enabled = False
    cmdlot2.Enabled = False
End Sub
