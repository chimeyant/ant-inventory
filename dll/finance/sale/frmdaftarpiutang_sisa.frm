VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarpiutang_sisa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Sisa Piutang"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Salesman"
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
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rekap Period by Area/Rayon"
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
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rekap Semester by Area/Rayon"
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
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Area/Rayon"
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
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
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
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seluruh Saldo"
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo < 0"
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
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo > 1000"
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
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtinv3 
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
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtinv4 
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
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   4320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   3120
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
      MICON           =   "frmdaftarpiutang_sisa.frx":0000
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
      Left            =   2520
      TabIndex        =   12
      Top             =   3120
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
      MICON           =   "frmdaftarpiutang_sisa.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Code"
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
      MICON           =   "frmdaftarpiutang_sisa.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch4 
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Code"
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
      MICON           =   "frmdaftarpiutang_sisa.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   2640
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
      Format          =   90112003
      CurrentDate     =   37845
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM yyyy"
      Format          =   90112003
      UpDown          =   -1  'True
      CurrentDate     =   37845
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   -105
      TabIndex        =   18
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "To  D a t e"
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
      TabIndex        =   16
      Top             =   2670
      Width           =   975
   End
End
Attribute VB_Name = "frmdaftarpiutang_sisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2 As String
Dim Los As Boolean
Dim Paj As Boolean

Private Sub cariinv3()
    If txtinv3 = "" Then Exit Sub
    
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Then
        OBJ.Open dsn
        SQL = "select kode from am_area where kode = '" & txtinv3 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv3 = ""
            txtinv3.SetFocus
        End If
        OBJ.Close
    ElseIf Option4.Value = True Then
        OBJ.Open dsn
        SQL = "select kodecust from am_customer where kodecust = '" & txtinv3 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv3 = ""
            txtinv3.SetFocus
        End If
        OBJ.Close
    ElseIf Option8.Value = True Then
        OBJ.Open dsn
        SQL = "select kodesales from am_salesman where kodesales = '" & txtinv3 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv3 = ""
            txtinv3.SetFocus
        End If
        OBJ.Close
    End If
End Sub

Private Sub cariinv4()
    If txtinv4 = "" Then Exit Sub
    
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Then
        OBJ.Open dsn
        SQL = "select kode from am_area where kode = '" & txtinv4 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
        End If
        OBJ.Close
    ElseIf Option4.Value = True Then
        OBJ.Open dsn
        SQL = "select kodecust from am_customer where kodecust = '" & txtinv4 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
        End If
        OBJ.Close
    ElseIf Option8.Value = True Then
        OBJ.Open dsn
        SQL = "select kodesales from am_salesman where kodesales = '" & txtinv4 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data not found.", vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
        End If
        OBJ.Close
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If txtinv3 = "" Or txtinv4 = "" Then Exit Sub
    If txtinv4 < txtinv3 Then
        MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
        txtinv4 = ""
        txtinv4.SetFocus
        Exit Sub
    End If
    
    If Option1.Value = True Then str1 = "1"
    If Option3.Value = True Then str1 = "2"
    If Option2.Value = True Then str1 = "3"
    
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Then str2 = "area"
    If Option4.Value = True Then str2 = "cust"
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowPrintBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    
    If Option5.Value = True And Los = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_area_L.rpt"
        'MsgBox "L area"
    ElseIf Option5.Value = True And Paj = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_area.rpt"
        'MsgBox "P area"
    ElseIf Option5.Value = True And Los = False And Paj = False Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_area_P.rpt"
        'MsgBox "All area"
    End If
    
    If Option6.Value = True Then Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_rekap.rpt"
    If Option7.Value = True Then Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_period.rpt"
    If Option8.Value = True Then Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_sales.rpt"
    
    If Option4.Value = True And Los = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_L.rpt"
        'MsgBox "1"
    ElseIf Option4.Value = True And Paj = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa.rpt"
        'MsgBox "2"
    ElseIf Option4.Value = True And Los = False And Paj = False Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_sisa_P.rpt"
        'MsgBox "3"
    End If
    
    If Option7.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_piutang_sisaperiod)"
    ElseIf Option8.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_piutang_sisasales)"
    Else
        If Los = True Then
            Crystal.DataFiles(0) = "Proc(am_piutang_sisa_L)"
            'MsgBox "Proc area L"
        ElseIf Paj = True Then
            Crystal.DataFiles(0) = "Proc(am_piutang_sisa_P)"
            'MsgBox "Proc area P"
        Else
            Crystal.DataFiles(0) = "Proc(am_piutang_sisa)"
            'MsgBox "Proc area All"
        End If
    End If
    Los = False: Paj = False
    
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(1) = "@tanggal2 ;" + Format(date2, "yyyymmdd") + ";true"
    Crystal.ParameterFields(2) = "@kode1;" & txtinv3 & ";true"
    Crystal.ParameterFields(3) = "@kode2;" & txtinv4 & ";true"
    Crystal.ParameterFields(4) = "@pilih;" & str1 & ";true"
    If Option8.Value = False Then Crystal.ParameterFields(5) = "@pilihx;" & str2 & ";true"
    If Option7.Value = True Then Crystal.ParameterFields(6) = "@tanggal1 ;" + Format(date1, "yyyymmdd") + ";true"
    
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub date1_Change()
    date2.Day = 28
    date2.Month = date1.Month
    date2.Year = date1.Year
    date2 = date2 + 5
    date2.Day = 1
    date2 = date2 - 1
    date2.Month = date1.Month
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='304' and b.kodeuser = '1" & kuser & "'"
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then   'F2
        Los = True
        cmdview_Click
    ElseIf KeyCode = 112 Then   'F1
        Paj = True
        cmdview_Click
    End If
End Sub

Private Sub Form_Load()
   
    
    date2 = Date
    date1 = Date
End Sub

Private Sub Option4_Click()
    txtinv3 = ""
    txtinv4 = ""
    date2 = Date
    date1 = Date
    date1.Enabled = False
    date2.Enabled = True
End Sub

Private Sub Option5_Click()
    txtinv3 = ""
    txtinv4 = ""
    date2 = Date
    date1 = Date
    date1.Enabled = False
    date2.Enabled = True
End Sub

Private Sub Option6_Click()
    txtinv3 = ""
    txtinv4 = ""
    date2 = Date
    date1 = Date
    date1.Enabled = False
    date2.Enabled = True
End Sub

Private Sub Option7_Click()
    txtinv3 = ""
    txtinv4 = ""
    date2 = Date
    date1 = Date
    date1.Enabled = True
    date2.Enabled = False
    
    date2.Day = 28
    date2.Month = date1.Month
    date2.Year = date1.Year
    date2 = date2 + 5
    date2.Day = 1
    date2 = date2 - 1
    date2.Month = date1.Month
End Sub

Private Sub Option8_Click()
    txtinv3 = ""
    txtinv4 = ""
    date2 = Date
    date1 = Date
    date1.Enabled = False
    date2.Enabled = True
End Sub

Private Sub txtinv3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtinv4.SetFocus
End Sub

Private Sub txtinv4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option8.Value = True Then
        carisql1 = "select kodesales, namasales, idupdate from AM_salesman"
        namatabel = "Sales"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_Click()
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option8.Value = True Then
        carisql1 = "select kodesales, namasales, idupdate from AM_salesman"
        namatabel = "Sales"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv3 = hasil
    If Option8.Value = True Then carisales
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv4 = hasil
    If Option8.Value = True Then carisales2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub txtinv3_LostFocus()
    cariinv3
End Sub

Private Sub txtinv4_LostFocus()
    cariinv4
End Sub
Private Sub carisales()
    If txtinv3 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtinv3 & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & RST!namasales & " is not active !", vbExclamation, "Warning"
        txtinv3 = ""
    End If
    OBJ.Close
End Sub
Private Sub carisales2()
    If txtinv4 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtinv4 & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & RST!namasales & " is not active !", vbExclamation, "Warning"
        txtinv4 = ""
    End If
    OBJ.Close
End Sub
