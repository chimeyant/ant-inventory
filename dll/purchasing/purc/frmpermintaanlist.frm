VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpermintaanlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Permintaan"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton optoutstanding 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Outstanding"
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
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
      MICON           =   "frmpermintaanlist.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2880
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
      MICON           =   "frmpermintaanlist.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
      Format          =   134021123
      CurrentDate     =   38679
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
      Format          =   134021123
      CurrentDate     =   38679
   End
   Begin XtremeSuiteControls.RadioButton optclose 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Selesai"
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
   Begin XtremeSuiteControls.RadioButton optcancel 
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cancel"
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
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton optPO 
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   240
      Width           =   2175
      _Version        =   851970
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Outstanding (With PO)"
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label5 
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
      TabIndex        =   5
      Top             =   2310
      Width           =   975
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
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmpermintaanlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL1 As String

Private Sub cmdclear_Click()
    If (date1 > date2) Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    If optPO.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_perminpo)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\nota_permintaan.rpt"
    Else
        Crystal.DataFiles(0) = "Proc(am_perminlist)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\listnotapermintaan.rpt"
    End If
    
    Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
    If optoutstanding.Value = True Then
        Crystal.ParameterFields(2) = "@status;" + "0" + ";true"
    ElseIf optclose.Value = True Then
        Crystal.ParameterFields(2) = "@status;" + "1" + ";true"
    ElseIf optcancel.Value = True Then
        Crystal.ParameterFields(2) = "@status;" + "2" + ";true"
    ElseIf optPO.Value = True Then
        Crystal.ParameterFields(2) = "@status;" + "1" + ";true"
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
    
    
    'With rptpermintaanlist
         'SQL1 = "Exec am_printperminlist '" & tanggal1 & "','" & tanggal2 & "'"
        '.DataControl1.Source = SQL1
        '.DataControl1.ConnectionString = dsn
        '.Show
    'End With
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    date1.Value = Date
    date2.Value = Date
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
