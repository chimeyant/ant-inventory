VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmlaporanminstok 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Minimum Stock"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GBox 
      Height          =   1830
      Left            =   2070
      TabIndex        =   4
      Top             =   15
      Width           =   2580
      _Version        =   851970
      _ExtentX        =   4551
      _ExtentY        =   3228
      _StockProps     =   79
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtkode2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1515
         Width           =   1050
      End
      Begin VB.TextBox txtkode1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1230
         Width           =   1050
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   3075
         TabIndex        =   8
         Top             =   1230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "From Kode"
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
         MICON           =   "frmlaporanminstok.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "To Kode"
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
         MICON           =   "frmlaporanminstok.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton optsemua 
         Height          =   300
         Left            =   525
         TabIndex        =   11
         Top             =   960
         Width           =   1545
         _Version        =   851970
         _ExtentX        =   2725
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Tampilkan semua"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optmin 
         Height          =   300
         Left            =   525
         TabIndex        =   12
         Top             =   660
         Width           =   2025
         _Version        =   851970
         _ExtentX        =   3572
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Tampilkan qty minimum"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
      Begin VB.ComboBox cmbkode 
         Enabled         =   0   'False
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
         ItemData        =   "frmlaporanminstok.frx":0634
         Left            =   1065
         List            =   "frmlaporanminstok.frx":0636
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   2205
         Picture         =   "frmlaporanminstok.frx":0638
         Stretch         =   -1  'True
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Divisi"
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
         Left            =   30
         TabIndex        =   10
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.OptionButton optDivisi 
      Caption         =   "Per Sub Divisi"
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
      Left            =   195
      TabIndex        =   3
      Top             =   420
      Width           =   1590
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Semua Sub Divisi"
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
      Left            =   195
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1665
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      Top             =   1965
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
      MICON           =   "frmlaporanminstok.frx":1C3F
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
      Left            =   2850
      TabIndex        =   1
      Top             =   1965
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview"
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
      MICON           =   "frmlaporanminstok.frx":1F59
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   2235
      Top             =   1905
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2505
      Left            =   -15
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "frmlaporanminstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim str As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select a.KodeBarang,a.NamaBarang,a.KodeSatuan,b.NamaSatuan from am_apitemmst a"
    carisql1 = carisql1 + " left join am_apunit b on a.KodeSatuan = b.KodeSatuan  where a.KodeProduk = '" & cmbkode & "'"
    namatabel = "Barang per Divisi "
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode1 = hasil
    hasil = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select a.KodeBarang,a.NamaBarang,a.KodeSatuan,b.NamaSatuan from am_apitemmst a"
    carisql1 = carisql1 + " left join am_apunit b on a.KodeSatuan = b.KodeSatuan  where a.KodeProduk = '" & cmbkode & "'"
    namatabel = "Barang per Divisi "
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode2 = hasil
    hasil = ""
End Sub

Private Sub cmdview_Click()
    If optmin.Value = True Then str = "min"
    If optsemua.Value = True Then str = "all"
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    
    If optDivisi = True Then
        Crystal.DataFiles(0) = "Proc(am_posisi_mindivisi)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\tables\minstockdivisi.rpt"
        Crystal.ParameterFields(0) = "@kode1;" + Format(Date, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@kode2;" + Format(Date, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        Crystal.ParameterFields(3) = "@divisi;" + cmbkode + ";true"
        Crystal.ParameterFields(4) = "@pilih;" + str + ";True"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    ElseIf optAll = True Then
        Crystal.DataFiles(0) = "Proc(am_posisi_min)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\tables\minstockall.rpt"
        Crystal.ParameterFields(0) = "@kode1;" + Format(Date, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@kode2;" + Format(Date, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        Crystal.ParameterFields(3) = "@pilih;" + str + ";True"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    End If
End Sub

Private Sub Form_Load()
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub optAll_Click()
    cmbkode = ""
    cmbkode.Enabled = False
    cmdsearch1.Enabled = False
    cmdsearch2.Enabled = False
    txtkode1 = ""
    txtkode2 = ""
    Image1.Visible = True
End Sub

Private Sub optDivisi_Click()
    Image1.Visible = False
    cmbkode.Enabled = True
    cmdsearch1.Enabled = True
    cmdsearch2.Enabled = True
End Sub
