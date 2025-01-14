VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarpiutang_kartu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Kartu Piutang"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   11
      Top             =   360
      Width           =   1335
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
      TabIndex        =   10
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
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
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   960
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
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
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
      Format          =   135659523
      CurrentDate     =   37845
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
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
      MICON           =   "frmdaftarpiutang_kartu.frx":0000
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
      TabIndex        =   4
      Top             =   2640
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
      MICON           =   "frmdaftarpiutang_kartu.frx":031A
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
      TabIndex        =   7
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From"
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
      MICON           =   "frmdaftarpiutang_kartu.frx":0634
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
      TabIndex        =   8
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To"
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
      MICON           =   "frmdaftarpiutang_kartu.frx":094E
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
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
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
      Format          =   138608643
      CurrentDate     =   37845
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label3 
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
      TabIndex        =   9
      Top             =   2190
      Width           =   975
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
      TabIndex        =   6
      Top             =   1830
      Width           =   975
   End
End
Attribute VB_Name = "frmdaftarpiutang_kartu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cariinv3()
    If txtinv3 = "" Then Exit Sub
    
    If Option5.Value = True Then
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
    End If
End Sub

Private Sub cariinv4()
    If txtinv4 = "" Then Exit Sub
    
    If Option5.Value = True Then
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
    If Option5.Value = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_kartuarea.rpt"
        Crystal.DataFiles(0) = "Proc(am_piutang_kartuarea)"
    ElseIf Option4.Value = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_kartu.rpt"
        'Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_kartu_new.rpt"
        Crystal.DataFiles(0) = "Proc(am_piutang_kartu)"
    End If
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(1) = "@tanggal1 ;" + Format(date1, "yyyymmdd") + ";true"
    Crystal.ParameterFields(2) = "@tanggal2 ;" + Format(date2, "yyyymmdd") + ";true"
    Crystal.ParameterFields(3) = "@kode1;" & txtinv3 & ";true"
    Crystal.ParameterFields(4) = "@kode2;" & txtinv4 & ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='294' and b.kodeuser = '1" & kuser & "'"
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

Private Sub Form_Load()
   
    date1 = Date
    date2 = Date
End Sub

Private Sub Option4_Click()
    txtinv3 = ""
    txtinv4 = ""
End Sub

Private Sub Option5_Click()
    txtinv3 = ""
    txtinv4 = ""
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
    If Option5.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_Click()
    If Option5.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv3 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv4 = hasil
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
