VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmsorderlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Sales Order"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsorderlist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   4560
      Width           =   3255
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
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   24
         Top             =   0
         Width           =   2055
      End
      Begin Chameleon.chameleonButton cmdbrg 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Produk"
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
         MICON           =   "frmsorderlist.frx":2372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblkdbrg 
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Produk"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   133496835
      CurrentDate     =   37464
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   3255
      Begin VB.TextBox txtkodecust 
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Customer"
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
         MICON           =   "frmsorderlist.frx":268C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Customer"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1455
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Sales Order [CLOSING]"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Outstanding Sales Order (Tamansari)"
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
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Sales Order [CANCEL]"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by No. SO"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Import Sales Order (Semua)"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   2895
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Outstanding Sales Order"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Import Sales Order (Baru)"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Sales Order Belum Export"
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
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Sales Order Sudah Export"
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Semua Sales Order"
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   133496835
      CurrentDate     =   37464
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   5040
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
      MICON           =   "frmsorderlist.frx":29A6
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
      Left            =   1200
      TabIndex        =   2
      Top             =   5040
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmsorderlist.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   3225
      Left            =   -240
      TabIndex        =   16
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "To Date"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3750
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3390
      Width           =   1215
   End
End
Attribute VB_Name = "frmsorderlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String

Private Sub Check2_Click()
    If Check2.Value = 0 Then
        Frame1.Visible = False
    Else
        Frame1.Visible = True
        txtkodecust = ""
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
        txtproduk = ""
        lblkdbrg = ""
    End If
End Sub

Private Sub cmdbrg_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    frmsearch.Show vbModal
End Sub

Private Sub cmdbrg_GotFocus()
    If hasil = "" Then Exit Sub
    lblkdbrg = hasil
    txtproduk = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecust = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdview_Click()
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    If Option5.Value = True Then
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowExportBtn = True
        crystal.WindowShowProgressCtls = True
        crystal.WindowShowRefreshBtn = True
        crystal.Connect = dsnreport
        If Check2.Value = 1 And Check3.Value = 0 Then
            'Outstanding SO By Customer only 16/12/2021
            crystal.DataFiles(0) = "Proc(am_outstanding_bycust)"
            If Check1.Value = 0 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstanding_bycust.rpt"
            If Check1.Value = 1 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstandingso_bycust.rpt"
            crystal.ParameterFields(0) = "@kdcust;" + txtkodecust + ";true"
            crystal.ParameterFields(1) = "@kode2;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@kode3;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(3) = "@namauser ;" + nmuser + ";true"
        ElseIf Check3.Value = 1 And Check2.Value = 0 Then
            'Outstanding SO By Produk only
            crystal.DataFiles(0) = "Proc(am_outstanding_byprod)"
            If Check1.Value = 0 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstanding_byprod.rpt"
            If Check1.Value = 1 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstandingso_byprod.rpt"
            crystal.ParameterFields(0) = "@kdbrg;" + lblkdbrg + ";true"
            crystal.ParameterFields(1) = "@kode2;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@kode3;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(3) = "@namauser ;" + nmuser + ";true"
        ElseIf Check2.Value = 1 And Check3.Value = 1 Then
            'Outstanding SO By Customer & Produk
            crystal.DataFiles(0) = "Proc(am_outstanding_byprodcust)"
            If Check1.Value = 0 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstanding_byprodcust.rpt"
            If Check1.Value = 1 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstandingso_byprodcust.rpt"
            crystal.ParameterFields(0) = "@kdcust;" + txtkodecust + ";true"
            crystal.ParameterFields(1) = "@kode2;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@kode3;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(3) = "@kdbrg;" + lblkdbrg + ";true"
            crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
        ElseIf Check2.Value = 0 Then
            crystal.DataFiles(0) = "Proc(am_outstanding)"
            If Check1.Value = 0 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstanding.rpt"
            If Check1.Value = 1 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\outstandingso_.rpt"
            crystal.ParameterFields(0) = "@kode1;a;true"
            crystal.ParameterFields(1) = "@kode2;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@kode3;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(3) = "@namauser ;" + nmuser + ";true"
        End If
        crystal.RetrieveDataFiles
        crystal.Action = 1
    ElseIf Option8.Value = True Then
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowExportBtn = True
        crystal.WindowShowProgressCtls = True
        crystal.WindowShowRefreshBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(am_outstandingsonya)"
        crystal.ReportFileName = AppPath & "\reports\sale\inv\outstandingsonya.rpt"
        crystal.ParameterFields(0) = "@namauser ;" + nmuser + ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
    ElseIf Option9.Value = True Then
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowExportBtn = True
        crystal.WindowShowProgressCtls = True
        crystal.WindowShowRefreshBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(am_soclosing)"
        crystal.ReportFileName = AppPath & "\reports\sale\inv\daftarsoclose.rpt"
        crystal.ParameterFields(0) = "@namauser ;" + nmuser + ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
    Else
        If Option1.Value = True Then str1 = "a"
        If Option2.Value = True Then str1 = "b"
        If Option3.Value = True Then str1 = "c"
        If Option4.Value = True Then str1 = "d"
        If Option6.Value = True Then str1 = "e"
        If Option7.Value = True Then str1 = "f"
        
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowExportBtn = True
        crystal.WindowShowProgressCtls = True
        crystal.WindowShowRefreshBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(am_daftarso)"
        If Check1.Value = 0 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\daftarso.rpt"
        If Check1.Value = 1 Then crystal.ReportFileName = AppPath & "\reports\sale\inv\daftarso_.rpt"
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
        crystal.ParameterFields(3) = "@pilih ;" + str1 + ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
End Sub

Private Sub Option1_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option2_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option3_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option4_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option5_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = True
    Check2.Value = 0
    Check3.Visible = True
    Check3.Value = 0
End Sub

Private Sub Option6_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option7_Click()
    date1.Enabled = True
    date2.Enabled = True
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option8_Click()
    date1.Enabled = False
    date2.Enabled = False
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub

Private Sub Option9_Click()
    date1.Enabled = False
    date2.Enabled = False
    Check2.Visible = False
    Check3.Visible = False
    Frame1.Visible = False
End Sub
