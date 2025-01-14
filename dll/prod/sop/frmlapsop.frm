VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlapsop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cetak Laporan SOP"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin XtremeSuiteControls.RadioButton optlem 
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   1200
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Lem"
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
   Begin VB.TextBox txtkdbarang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   14
      Top             =   1575
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtbarang 
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
      Height          =   360
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   2640
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   12
      Top             =   1905
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail"
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
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   675
      Left            =   195
      TabIndex        =   6
      Top             =   510
      Width           =   3810
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua"
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
         Left            =   2835
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
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
         Left            =   2010
         TabIndex        =   9
         Top             =   300
         Width           =   900
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lengkap"
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
         Left            =   960
         TabIndex        =   8
         Top             =   285
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proses"
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
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   900
      End
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   750
      TabIndex        =   3
      Top             =   150
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
      Format          =   143392769
      CurrentDate     =   42039
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   5520
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   2895
      TabIndex        =   0
      Top             =   2520
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton cmdcetak 
      Height          =   465
      Left            =   1755
      TabIndex        =   1
      Top             =   2520
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "View"
      BackColor       =   -2147483633
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   315
      Left            =   2580
      TabIndex        =   5
      Top             =   150
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
      Format          =   143392769
      CurrentDate     =   42039
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   465
      Left            =   600
      TabIndex        =   15
      Top             =   2520
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Clear"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.RadioButton optkaret 
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   1200
      Width           =   855
      _Version        =   851970
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Karet"
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
      Left            =   2640
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " Karet dan Lem"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   675
      Left            =   0
      Top             =   2415
      Width           =   4080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Top             =   210
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   2
      Top             =   210
      Width           =   465
   End
End
Attribute VB_Name = "frmlapsop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Check2_Click()
    If Check2.Value = Checked Then
        If optlem.Value = True Then
            namatabel = "produk."
            carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'L%'"
        ElseIf optkaret.Value = True Then
            namatabel = "produk."
            carisql1 = "select kode_produk,nama_produk from list_produk_master Where kode_produk like 'K%'"
        Else
            namatabel = "produk"
            carisql1 = "select kode_produk,nama_produk from list_produk_master"
        End If
        frmsearch.Show vbModal
    ElseIf Check2.Value = Unchecked Then
        txtkdbarang = ""
        txtbarang = ""
    End If
End Sub

Private Sub Check2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkdbarang = hasil
    txtbarang = hasil1
    hasil = ""
    hasil1 = ""
    namatabel = ""
    carisql1 = ""
End Sub

Private Sub cmdcetak_Click()
'On Error Resume Next
Dim cetak_ke As Integer
Dim akses As Boolean
    If date1 > date2 Then
        MsgBox "Batas tanggal tidak benar..!", vbCritical, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From LIST_USERS Where username = '" & nmuser & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!gl = "1" Then
            akses = True
        Else
            akses = False
            If nmuser = "ENAH" Then akses = True
        End If
    Else
        If nmuser = "Creator" Then akses = True
    End If
    OBJ.Close
    
    'proses cetak sop
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport

    If Check2.Value = Unchecked Then
        If Check1.Value = Checked Then
            'DETAIL LOT ALL
            If optKL.Value = True Then
                crystal.DataFiles(0) = "Proc(am_cetaklistsopdetail)"
                If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sop_detail.rpt"
                If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sop_detailB.rpt"
                'crystal.ReportFileName = AppPath & "\reports\produksi\cetaksop_byprodetailtes.rpt"
            ElseIf optkaret.Value = True Or optlem.Value = True Then
                crystal.DataFiles(0) = "Proc(am_cetaklistsopdetailKL)"
                If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sop_detailkl.rpt"
                If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sop_detailBkl.rpt"
            End If
        Else
            'REKAP LOT ALL
            If optKL.Value = True Then
                crystal.DataFiles(0) = "Proc(am_cetaklistsop)"
                If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sopA.rpt"
                If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sopB.rpt"
            ElseIf optkaret.Value = True Or optlem.Value = True Then
                crystal.DataFiles(0) = "Proc(am_cetaklistsopKL)"
                If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sopAkl.rpt"
                If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\daftar_sopBkl.rpt"
            End If
        End If
        crystal.ParameterFields(0) = "@kode1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@kode2;" & Format(date2, "yyyy/MM/dd") & ";true"
        If Option1 = True Then 'Proses
            crystal.ParameterFields(2) = "@kode3;" & "1" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "3" & ";true"
            If optkaret.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "K" & ";true"
            If optlem.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "L" & ";true"
        End If
        If Option2 = True Then 'Lengkap
            crystal.ParameterFields(2) = "@kode3;" & "4" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "4" & ";true"
            If optkaret.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "K" & ";true"
            If optlem.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "L" & ";true"
        End If
        If Option3 = True Then 'Close
            crystal.ParameterFields(2) = "@kode3;" & "5" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "5" & ";true"
            If optkaret.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "K" & ";true"
            If optlem.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "L" & ";true"
        End If
        If Option4 = True Then 'Semua
            crystal.ParameterFields(2) = "@kode3;" & "1" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "5" & ";true"
            If optkaret.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "K" & ";true"
            If optlem.Value = True Then crystal.ParameterFields(4) = "@kode5;" & "L" & ";true"
        End If
        
    ElseIf Check2.Value = Checked Then
        If Check1.Value = Checked Then
            'DETAIL LOT BY PRODUK
            crystal.DataFiles(0) = "Proc(am_cetaksop_byprodetail)"
            If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\cetaksop_byprodetail.rpt"
            If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\cetaksop_byprodetailB.rpt"
            crystal.ParameterFields(4) = "@produk;" & txtkdbarang & ";true"
        Else
            'REKAP LOT BY PRODUK
            crystal.DataFiles(0) = "Proc(am_cetaksop_byproduk)"
            If akses = True Then crystal.ReportFileName = AppPath & "\reports\produksi\cetaksop_byprodukA.rpt"
            If akses = False Then crystal.ReportFileName = AppPath & "\reports\produksi\cetaksop_byprodukB.rpt"
            crystal.ParameterFields(4) = "@produk;" & txtkdbarang & ";true"
        End If
        crystal.ParameterFields(0) = "@kode1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@kode2;" & Format(date2, "yyyy/MM/dd") & ";true"
        If Option1 = True Then 'Proses
            crystal.ParameterFields(2) = "@kode3;" & "1" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "3" & ";true"
        End If
        If Option2 = True Then 'Lengkap
            crystal.ParameterFields(2) = "@kode3;" & "4" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "4" & ";true"
        End If
        If Option3 = True Then 'Close
            crystal.ParameterFields(2) = "@kode3;" & "5" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "5" & ";true"
        End If
        If Option4 = True Then 'Semua
            crystal.ParameterFields(2) = "@kode3;" & "1" & ";true"
            crystal.ParameterFields(3) = "@kode4;" & "5" & ";true"
        End If
    End If

    crystal.RetrieveDataFiles
    crystal.Action = 1
    txtkdbarang = ""
    txtbarang = ""
    Check2.Value = Unchecked
End Sub

Private Sub cmdclear_Click()
    txtkdbarang = ""
    txtbarang = ""
    Check2.Value = Unchecked
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub
