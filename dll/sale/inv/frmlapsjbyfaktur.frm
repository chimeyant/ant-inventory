VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmsjbyfaktur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan surat jalan By Faktur"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton optterbit 
      Height          =   435
      Left            =   270
      TabIndex        =   6
      Top             =   195
      Width           =   2865
      _Version        =   851970
      _ExtentX        =   5054
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "SJ sudah terbit Faktur"
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1320
      Left            =   210
      TabIndex        =   1
      Top             =   1710
      Width           =   3300
      _Version        =   851970
      _ExtentX        =   5821
      _ExtentY        =   2328
      _StockProps     =   79
      Caption         =   "Tanggal Surat Jalan"
      ForeColor       =   -2147483631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   1215
         TabIndex        =   2
         Top             =   420
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
         Format          =   143458307
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   285
         Left            =   1230
         TabIndex        =   3
         Top             =   780
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
         Format          =   143458307
         CurrentDate     =   37426
      End
      Begin VB.Label Label7 
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
         Left            =   375
         TabIndex        =   5
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Left            =   375
         TabIndex        =   4
         Top             =   450
         Width           =   975
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3210
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
      MICON           =   "frmlapsjbyfaktur.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton optblm 
      Height          =   435
      Left            =   270
      TabIndex        =   7
      Top             =   525
      Width           =   2865
      _Version        =   851970
      _ExtentX        =   5054
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "SJ belum terbit Faktur"
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
   Begin XtremeSuiteControls.RadioButton optsemua 
      Height          =   435
      Left            =   270
      TabIndex        =   8
      Top             =   855
      Width           =   2865
      _Version        =   851970
      _ExtentX        =   5054
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "All"
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
      Left            =   1695
      TabIndex        =   9
      Top             =   3210
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
      MICON           =   "frmlapsjbyfaktur.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton optmanual 
      Height          =   435
      Left            =   270
      TabIndex        =   10
      Top             =   1185
      Width           =   2865
      _Version        =   851970
      _ExtentX        =   5054
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "SJ Manual"
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
      Left            =   180
      Top             =   3060
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
End
Attribute VB_Name = "frmsjbyfaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    If optterbit.Value = True Then str1 = "a"
    If optblm.Value = True Then str1 = "b"
    If optsemua.Value = True Then str1 = "c"
    If optmanual.Value = True Then str1 = "d"
        
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowExportBtn = True
    crystal.WindowShowProgressCtls = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_sjbyfaktur)"
    If optterbit.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\sjbyfaktur.rpt"
    If optblm.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\sjbyfaktur2.rpt"
    If optsemua.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\sjbyfakturall.rpt"
    If optmanual.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\sjbyfakturman.rpt"
    
    crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
    crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
    crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    crystal.ParameterFields(3) = "@pilih ;" + str1 + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub
