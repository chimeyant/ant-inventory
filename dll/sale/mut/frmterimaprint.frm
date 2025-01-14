VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmterimaprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print/List ..."
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmterimaprint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nilai Retur Penjualan"
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
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Request for Stock (Pabrik)"
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
      Left            =   360
      TabIndex        =   12
      Top             =   3960
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pindah Gudang"
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
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Request for Stock"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nota Retur"
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
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
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
      Format          =   91291651
      CurrentDate     =   38773
   End
   Begin VB.TextBox txtnodo2 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtnodo1 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   3000
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Dr No Bukti"
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
      MICON           =   "frmterimaprint.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
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
      MICON           =   "frmterimaprint.frx":268C
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
      Left            =   1320
      TabIndex        =   9
      Top             =   2760
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
      MICON           =   "frmterimaprint.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "s/d No Bukti"
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
      MICON           =   "frmterimaprint.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
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
      Format          =   91291651
      CurrentDate     =   38773
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "s/d Tanggal"
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
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Dari Tanggal"
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
      TabIndex        =   13
      Top             =   1950
      Width           =   1215
   End
End
Attribute VB_Name = "frmterimaprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cariinv1()
    If txtnodo1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_bpbhdr where nobpb = '" & txtnodo1 & "' and type = '04'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Nota Retur " & txtnodo1 & " Not Found.", vbExclamation, "Warning"
        txtnodo1 = ""
        txtnodo1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariinv2()
    If txtnodo2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_bpbhdr where nobpb = '" & txtnodo2 & "' and type = '04'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Nota Retur " & txtnodo2 & " Not Found.", vbExclamation, "Warning"
        txtnodo2 = ""
        txtnodo2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where type = '04' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
    Else
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where type = '04'"
    End If
    namatabel = "Mutasi Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtnodo1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where type = '04' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
    Else
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where type = '04'"
    End If
    namatabel = "Mutasi Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnodo2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If Option1.Value = True Then
        If txtnodo1 = "" Or txtnodo2 = "" Then Exit Sub
        
        If txtnodo1 > txtnodo2 Then
            MsgBox "Error on No bukti.", vbInformation, "information"
            Exit Sub
        End If
    End If
    
    If date1 > date2 Then
        MsgBox "Error on Date.", vbInformation, "information"
        Exit Sub
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    If Option1.Value = True Then
        crystal.DataFiles(0) = "Proc(am_notaretur)"
        crystal.ReportFileName = AppPath & "\reports\sale\mut\notaretur.rpt"
        crystal.ParameterFields(0) = "@noinv1;" & txtnodo1 & ";true"
        crystal.ParameterFields(1) = "@noinv2;" & txtnodo2 & ";true"
        crystal.ParameterFields(2) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(3) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
    ElseIf Option2.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarrequest)"
        crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarequest.rpt"
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    ElseIf Option3.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarpindah)"
        crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarpindah.rpt"
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    ElseIf Option4.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarrequestpabrik)"
        crystal.ReportFileName = AppPath & "\reports\sale\mur\daftarequestpabrik.rpt"
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    ElseIf Option5.Value = True Then
        crystal.DataFiles(0) = "Proc(am_nilairetur)"
        crystal.ReportFileName = AppPath & "\reports\sale\mut\nilairetur.rpt"
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='114' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
            "Please contact your Administrator.", vbCritical, "User Rights"
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

Private Sub Form_Load()
   
    
    date1 = Date
    date2 = Date
End Sub

Private Sub Option1_Click()
    txtnodo1.Enabled = True
    txtnodo2.Enabled = True
    cmdsearch.Enabled = True
    cmdsearch1.Enabled = True
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub Option2_Click()
    txtnodo1 = ""
    txtnodo2 = ""
    txtnodo1.Enabled = False
    txtnodo2.Enabled = False
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub Option3_Click()
    txtnodo1 = ""
    txtnodo2 = ""
    txtnodo1.Enabled = False
    txtnodo2.Enabled = False
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub Option4_Click()
    txtnodo1 = ""
    txtnodo2 = ""
    txtnodo1.Enabled = False
    txtnodo2.Enabled = False
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub Option5_Click()
    txtnodo1 = ""
    txtnodo2 = ""
    txtnodo1.Enabled = False
    txtnodo2.Enabled = False
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub txtnodo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnodo2.SetFocus
End Sub

Private Sub txtnodo1_LostFocus()
    cariinv1
End Sub

Private Sub txtnodo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtnodo2_LostFocus()
    cariinv2
End Sub
