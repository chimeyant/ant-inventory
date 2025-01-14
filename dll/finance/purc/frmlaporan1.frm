VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlaporan1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Confirm/UnConfirm Penerimaan+Retur"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
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
   ScaleHeight     =   1920
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Retur Penerimaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   374079491
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
      Left            =   5040
      TabIndex        =   4
      Top             =   1400
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
      MICON           =   "frmlaporan1.frx":0000
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
      Left            =   4140
      TabIndex        =   3
      Top             =   1395
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
      MICON           =   "frmlaporan1.frx":031A
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
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   374079491
      CurrentDate     =   38679
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   630
      Width           =   975
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2566;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Caption         =   "From Date                                                     To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   990
      Width           =   3735
   End
End
Attribute VB_Name = "frmlaporan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Combo2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='105' and b.kodeuser = '2" & kuser & "'"
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
    If Combo2 = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If date1 > date2 Then
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
    If Check1.Value = 0 Then
        Crystal.DataFiles(0) = "Proc(am_laporbeli_confirm)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\pembelian_detail_confirm.rpt"
    Else
        Crystal.DataFiles(0) = "Proc(am_laporbeli_returconfirm)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\pembelian_retur_confirm.rpt"
    End If
    Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(3) = "@pilih;" + Mid(Combo2, 1, 1) + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    Combo2.AddItem "Confirm"
    Combo2.AddItem "UnConfirm"
End Sub
