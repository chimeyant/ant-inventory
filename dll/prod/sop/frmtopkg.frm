VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmtopkg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Top Produk by. kilogram"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
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
      Height          =   315
      Left            =   915
      TabIndex        =   10
      Top             =   840
      Width           =   1050
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
      Height          =   315
      Left            =   1995
      TabIndex        =   9
      Top             =   840
      Width           =   3060
   End
   Begin VB.TextBox txtkodeproduk 
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
      Height          =   315
      Left            =   915
      TabIndex        =   7
      Top             =   480
      Width           =   1050
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
      Height          =   315
      Left            =   1995
      TabIndex        =   6
      Top             =   480
      Width           =   3060
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   3960
      TabIndex        =   0
      Top             =   2160
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
   Begin XtremeSuiteControls.GroupBox GB1 
      Height          =   810
      Left            =   915
      TabIndex        =   1
      Top             =   1320
      Width           =   4140
      _Version        =   851970
      _ExtentX        =   7302
      _ExtentY        =   1429
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   705
         TabIndex        =   2
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
         Format          =   142934017
         CurrentDate     =   42039
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   315
         Left            =   2535
         TabIndex        =   3
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
         Format          =   142934017
         CurrentDate     =   42039
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         TabIndex        =   5
         Top             =   375
         Width           =   315
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         TabIndex        =   4
         Top             =   375
         Width           =   465
      End
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.PushButton cmdproduksi 
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   495
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "From :"
      BackColor       =   16777215
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
      TextAlignment   =   1
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton cmdproduk 
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   855
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "To :"
      BackColor       =   16777215
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
      TextAlignment   =   1
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   465
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   3810
      _Version        =   851970
      _ExtentX        =   6720
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
   Begin XtremeSuiteControls.RadioButton optlem 
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Lem"
      BackColor       =   16777215
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
   Begin XtremeSuiteControls.RadioButton optkaret 
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      Width           =   855
      _Version        =   851970
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Karet"
      BackColor       =   16777215
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
   Begin XtremeSuiteControls.RadioButton optKL 
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Semua produk"
      BackColor       =   16777215
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
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Height          =   1305
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmtopkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnview_Click()
    If txtkodeproduk = "" And txtkode = "" Then Exit Sub
    If date1 > date2 Then
        MsgBox "Batas tanggal tidak benar..!", vbCritical, AppName
        Exit Sub
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetaklistkgtop)"
    crystal.ReportFileName = AppPath & "\reports\produksi\daftar_kgsop.rpt"
    crystal.ParameterFields(0) = "@kode1;" + txtkodeproduk + ";true"
    crystal.ParameterFields(1) = "@kode2;" + txtkode + ";true"
    crystal.ParameterFields(2) = "@kode3;" & Format(date1, "yyyy/MM/dd") & ";true"
    crystal.ParameterFields(3) = "@kode4;" & Format(date2, "yyyy/MM/dd") & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdproduk_Click()
    If optlem.Value = True Then
        namatabel = "produk."
        carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'L%'"
    ElseIf optkaret.Value = True Then
        namatabel = "produk."
        carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'K%'"
    ElseIf optKL.Value = True Then
        namatabel = "produk"
        carisql1 = "select kode_produk,nama_produk from list_produk_master"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduk_GotFocus()
    If hasil = "" Then Exit Sub
    If optKL.Value = True Then
        txtkode = hasil
        txtproduk2 = hasil1
    Else
        txtkode = hasil1
        txtproduk2 = hasil2
    End If
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdproduksi_Click()
    If optlem.Value = True Then
        namatabel = "produk."
        carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'L%'"
    ElseIf optkaret.Value = True Then
        namatabel = "produk."
        carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'K%'"
    ElseIf optKL.Value = True Then
        namatabel = "produk"
        carisql1 = "select kode_produk,nama_produk from list_produk_master"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduksi_GotFocus()
    If hasil = "" Then Exit Sub
    If optKL.Value = True Then
        txtkodeproduk = hasil
        txtproduk = hasil1
    Else
        txtkodeproduk = hasil1
        txtproduk = hasil2
    End If
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Private Sub clearform()
    txtkode = ""
    txtkodeproduk = ""
    txtproduk = ""
    txtproduk2 = ""
End Sub

Private Sub optkaret_Click()
    Call clearform
End Sub

Private Sub optKL_Click()
    Call clearform
End Sub

Private Sub optlem_Click()
    Call clearform
End Sub
