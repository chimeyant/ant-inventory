VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlaporan3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listing Pembayaran Hutang"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   192348163
      CurrentDate     =   38679
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   240
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
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
      MICON           =   "frmlaporan3.frx":0000
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
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
      MICON           =   "frmlaporan3.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   192348163
      CurrentDate     =   38679
   End
   Begin VB.Label Label1 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "frmlaporan3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='55' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    If date1 > date2 Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    'crystal.DataFiles(0) = "Proc(am_payment)" 'tidak ada no voucher
    crystal.DataFiles(0) = "Proc(am_payment_new)" 'tambah no voucher
    crystal.ReportFileName = AppPath & "\reports\finance\purc\hutang_bayar.rpt"
    crystal.ParameterFields(0) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
    crystal.ParameterFields(1) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
    crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    date1 = Date
    date2 = Date
End Sub
