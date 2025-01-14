VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarTT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan T.T. Pembayaran"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
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
   Icon            =   "frmdaftarTT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   720
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
      Format          =   89784323
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
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
      MICON           =   "frmdaftarTT.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
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
      MICON           =   "frmdaftarTT.frx":268C
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
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
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
      Format          =   89784323
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   240
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmdaftarTT.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1110
      Width           =   975
   End
End
Attribute VB_Name = "frmdaftarTT"
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
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='334' and b.kodeuser = '1" & kuser & "'"
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

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtkode_LostFocus()
    If txtKode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Customer " & txtKode & " Not Found.", vbExclamation, "Warning"
        txtKode = ""
        txtKode.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    hasil = ""
    hasil1 = ""
    date1.SetFocus
End Sub

Private Sub cmdadd_Click()
    If txtKode = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_daftartt)"
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\daftartt.rpt"
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(1) = "@tanggal1 ;" + Format(date1, "yyyymmdd") + ";true"
    Crystal.ParameterFields(2) = "@tanggal2 ;" + Format(date2, "yyyymmdd") + ";true"
    Crystal.ParameterFields(3) = "@kode;" & txtKode & ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    date1.Value = Date
    date2.Value = Date
End Sub
