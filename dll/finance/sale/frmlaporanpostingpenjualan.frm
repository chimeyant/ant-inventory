VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmlaporanpostingpenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Posting Penjualan"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   360
      Left            =   3285
      TabIndex        =   0
      Top             =   1110
      Width           =   1140
      _Version        =   851970
      _ExtentX        =   2011
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker date1 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Top             =   345
      Width           =   1440
      _Version        =   851970
      _ExtentX        =   2540
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.DateTimePicker date2 
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   675
      Width           =   1440
      _Version        =   851970
      _ExtentX        =   2540
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton cmdprint 
      Height          =   360
      Left            =   2130
      TabIndex        =   4
      Top             =   1110
      Width           =   1140
      _Version        =   851970
      _ExtentX        =   2011
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Print"
      UseVisualStyle  =   -1  'True
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   1065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Caption         =   "KODE TRANSAKSI"
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Dari Tanggal"
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   390
      Width           =   1410
   End
   Begin VB.Label Label3 
      Caption         =   "S.D Tanggal"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   705
      Width           =   1410
   End
End
Attribute VB_Name = "frmlaporanpostingpenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdprint_Click()
    If date1 > date2 Then
        MsgBox "Invalid date range, preview abort.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If txtkode = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
            
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_lap_posting_penerimaan)"
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\rpt_lap_posting_penjualan.rpt"
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(1) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(2) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(3) = "@kode;" + txtkode + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

