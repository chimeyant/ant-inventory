VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmgroup_report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Group Pengisian & Pengemasan"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton optdate 
      Height          =   360
      Left            =   300
      TabIndex        =   9
      Top             =   165
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "By Date"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   1215
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
      Height          =   360
      Left            =   300
      TabIndex        =   10
      Top             =   570
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   635
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
   End
   Begin XtremeSuiteControls.GroupBox GB2 
      Height          =   945
      Left            =   1650
      TabIndex        =   6
      Top             =   105
      Visible         =   0   'False
      Width           =   4095
      _Version        =   851970
      _ExtentX        =   7223
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   "By Lot"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "News706 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      BorderStyle     =   1
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
         Height          =   360
         Left            =   1095
         TabIndex        =   7
         Top             =   330
         Width           =   2265
      End
      Begin XtremeSuiteControls.PushButton cmdcari 
         Height          =   345
         Left            =   165
         TabIndex        =   8
         Top             =   330
         Width           =   840
         _Version        =   851970
         _ExtentX        =   1482
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "No. Lot"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
      End
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   465
      Left            =   3480
      TabIndex        =   11
      Top             =   1215
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
      Top             =   1245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.GroupBox GB1 
      Height          =   780
      Left            =   1650
      TabIndex        =   1
      Top             =   105
      Width           =   4095
      _Version        =   851970
      _ExtentX        =   7223
      _ExtentY        =   1376
      _StockProps     =   79
      Caption         =   "By Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "News706 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      BorderStyle     =   1
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   690
         TabIndex        =   2
         Top             =   330
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
         Format          =   136708097
         CurrentDate     =   42039
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   330
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
         Format          =   136708097
         CurrentDate     =   42039
      End
      Begin VB.Label Label3 
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
         Left            =   135
         TabIndex        =   5
         Top             =   390
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
         Left            =   2175
         TabIndex        =   4
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      Top             =   1110
      Width           =   5790
   End
End
Attribute VB_Name = "frmgroup_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnview_Click()
    If optdate.Value = True Then
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
    If optdate.Value = True Then
        crystal.DataFiles(0) = "Proc(am_produksigroup_bydate)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_grouplot_bydate.rpt"
        crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
        crystal.ParameterFields(2) = "@namauser;" & nmuser & ";True"
    ElseIf optlot.Value = True Then
        crystal.DataFiles(0) = "Proc(am_produksigroup_bylot)"
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_grouplot.rpt"
        crystal.ParameterFields(0) = "@nolot;" & txtnolot & ";true"
        crystal.ParameterFields(1) = "@namauser;" & nmuser & ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdcari_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    carisql1 = carisql1 + " where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Private Sub optdate_Click()
    optdate.Value = True
    GB1.Visible = True
    GB2.Visible = False
End Sub

Private Sub optlot_Click()
    optlot.Value = True
    GB2.Visible = True
    GB1.Visible = False
    txtnolot = ""
End Sub
