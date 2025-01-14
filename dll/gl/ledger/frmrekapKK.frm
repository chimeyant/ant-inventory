VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmrekapKK 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rekap Pengeluaran Kas"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   330
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Close"
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
   Begin MSComCtl2.DTPicker date1 
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
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
      Format          =   121438209
      CurrentDate     =   41743
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
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
      Format          =   121438209
      CurrentDate     =   41743
   End
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "View"
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
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TDBNumber6Ctl.TDBNumber txtsaldo 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   503
      Calculator      =   "frmrekapKK.frx":0000
      Caption         =   "frmrekapKK.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekapKK.frx":008C
      Keys            =   "frmrekapKK.frx":00AA
      Spin            =   "frmrekapKK.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtkm 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   503
      Calculator      =   "frmrekapKK.frx":0114
      Caption         =   "frmrekapKK.frx":0134
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekapKK.frx":01A0
      Keys            =   "frmrekapKK.frx":01BE
      Spin            =   "frmrekapKK.frx":0200
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Kas Masuk"
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
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Saldo Awal"
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
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmrekapKK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_rekapkk)"
    Crystal.ReportFileName = AppPath & "\reports\gl\ledger\rekapkk.rpt"
    Crystal.ParameterFields(0) = "@tgl1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(1) = "@tgl2;" + Format(date2, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(2) = "@user;" + nmuser + ";true"
    If txtsaldo.Value = "0.00" Or txtsaldo = "" Then
        Crystal.ParameterFields(3) = "@saldo;" + "0" + ";true"
    Else
        Crystal.ParameterFields(3) = "@saldo;" + Str(txtsaldo) + ";true"
    End If
    If txtkm.Value = "0.00" Or txtkm = "" Then
        Crystal.ParameterFields(4) = "@kasmasuk;" + "0" + ";true"
    Else
        Crystal.ParameterFields(4) = "@kasmasuk;" + Str(txtkm) + ";true"
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub
