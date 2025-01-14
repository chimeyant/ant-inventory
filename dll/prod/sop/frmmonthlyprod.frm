VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmmonthlyprod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Pemakaian bahan baku (Monthly)"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkdbahan 
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
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtbahan 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   2625
   End
   Begin VB.TextBox txtnmbarang 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   2625
   End
   Begin VB.TextBox txtkdbarang 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      _Version        =   851970
      _ExtentX        =   3625
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "View By"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000004&
         Caption         =   "Produk"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000004&
         Caption         =   "Bahan Baku"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   135790595
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   135790595
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date_tahun 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
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
      CustomFormat    =   "yyyy"
      Format          =   135790595
      CurrentDate     =   37464
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tahun"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "frmmonthlyprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        namatabel = "produk"
        carisql1 = "select kode_produk,nama_produk from list_produk_master"
        frmsearch.Show vbModal
    ElseIf Check1.Value = Unchecked Then
        txtkdbarang = ""
        txtnmbarang = ""
        txtkdbahan = ""
        txtbahan = ""
        Check2.Value = Unchecked
    End If
End Sub

Private Sub Check1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkdbarang = hasil
    txtnmbarang = hasil1
    hasil = ""
    hasil1 = ""
    namatabel = ""
    carisql1 = ""
End Sub

Private Sub Check2_Click()
    If Check2.Value = Checked Then
        If Check1.Value = Checked Then
            namatabel = "Bahan Tambahan."
            carisql1 = "select distinct kode_bahan,nama_bahan,inisial,kode_satuan from list_produk_child"
            carisql1 = carisql1 + " Where kode_produk = '" & txtkdbarang & "'"
        Else
            namatabel = "Bahan Tambahan"
            carisql1 = "select distinct kode_bahan,nama_bahan,inisial,kode_satuan from list_produk_child"
        End If
        frmsearch.Show vbModal
    ElseIf Check2.Value = Unchecked Then
        txtkdbahan = ""
        txtbahan = ""
    End If
End Sub

Private Sub Check2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkdbahan = hasil
    txtbahan = hasil1
    hasil = ""
    hasil1 = ""
    namatabel = ""
    carisql1 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    If Check1.Value = Checked And Check2.Value = Checked Then
        crystal.DataFiles(0) = "Proc(am_monthly_produk)"
        crystal.ReportFileName = AppPath & "\reports\produksi\monthly_prod.rpt"
    ElseIf Check1.Value = Unchecked And Check2.Value = Checked Then
        crystal.DataFiles(0) = "Proc(am_monthly_prodbahan)"
        crystal.ReportFileName = AppPath & "\reports\produksi\monthly_bahan.rpt"
    Else
        MsgBox "Laporan berdasarkan produk saja tidak tersedia", vbCritical, AppName
        Exit Sub
    End If
    
    crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(2) = "@bahan;" & txtkdbahan & ";true"
    If Check1.Value = Checked Then
        crystal.ParameterFields(3) = "@pilih;" & txtkdbarang & ";true"
    Else
        crystal.ParameterFields(3) = "@pilih;" & "semua" & ";true"
    End If
    crystal.ParameterFields(4) = "@namauser;" & UserOnline & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    Dim thn, bln, bln2, tgl, tgl2 As String
    date_tahun = Date
    
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
End Sub
Private Sub date_tahun_Change()
    Dim thn, bln, bln2, tgl, tgl2 As String
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
End Sub
