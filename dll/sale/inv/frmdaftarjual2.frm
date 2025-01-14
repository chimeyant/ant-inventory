VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarjual2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan Monthly..."
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2865
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   2865
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(monthly) Penjualan"
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
      Top             =   240
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(monthly) per Customer"
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
      Top             =   480
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(monthly) per Area Customer"
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
      Top             =   720
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(monthly) per Sales"
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
      Top             =   960
      Width           =   2415
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
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
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
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   2880
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
      Left            =   1800
      TabIndex        =   2
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
      MICON           =   "frmdaftarjual2.frx":0000
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
      Left            =   840
      TabIndex        =   3
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmdaftarjual2.frx":031A
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
      Left            =   120
      TabIndex        =   4
      Top             =   1680
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
      MICON           =   "frmdaftarjual2.frx":0634
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
      Left            =   120
      TabIndex        =   5
      Top             =   2040
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
      MICON           =   "frmdaftarjual2.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date_tahun 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
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
      CustomFormat    =   "yyyy"
      Format          =   135659523
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      CurrentDate     =   37464
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tahun"
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
      Left            =   840
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmdaftarjual2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, status As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch3_Click()
    If Option2.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option3.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
        namatabel = "Sales"
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

Private Sub cmdsearch4_Click()
    If Option2.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option3.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option4.Value = True Then
        carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
        namatabel = "Sales"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv4 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If txtinv4 < txtinv3 Then
        MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
        txtinv4 = ""
        txtinv4.SetFocus
        Exit Sub
    End If
    
    If Option1.Value = True Then str1 = "mall"
    If Option2.Value = True Then str1 = "mcust"
    If Option3.Value = True Then str1 = "marea"
    If Option4.Value = True Then str1 = "msales"
    status = "P"
        
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    
    crystal.Connect = dsnreport
    If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Then
        crystal.DataFiles(0) = "Proc(am_monthly)"
    ElseIf Option4.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualbulansale)"
    Else
    
    End If
    
    If Option1.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_all.rpt"
    ElseIf Option2.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl.rpt"
    ElseIf Option3.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_area.rpt"
    ElseIf Option4.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_sales.rpt"
    End If
    crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
    crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
    
    If Option4.Value = True Then
        crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
        crystal.ParameterFields(5) = "@PL ;" + status + ";true"
    Else
        crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
        crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
        crystal.ParameterFields(6) = "@PL ;" + status + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub date_tahun_Change()
    Dim thn, bln, bln2, tgl, tgl2 As String
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        'F2 Hanya penjualan yang L
        Me.MousePointer = vbHourglass
        If txtinv4 < txtinv3 Then
            MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
            Exit Sub
        End If
        
        If Option1.Value = True Then str1 = "mall"
        If Option2.Value = True Then str1 = "mcust"
        If Option3.Value = True Then str1 = "marea"
        If Option4.Value = True Then str1 = "msales"
        status = "L"
            
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowRefreshBtn = True
        
        crystal.Connect = dsnreport
        If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Then
            crystal.DataFiles(0) = "Proc(am_monthly)"
        ElseIf Option4.Value = True Then
            crystal.DataFiles(0) = "Proc(am_daftarjualbulansale)"
        Else
        
        End If
        
        If Option1.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_all.rpt"
        ElseIf Option2.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl.rpt"
        ElseIf Option3.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_area.rpt"
        ElseIf Option4.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_sales.rpt"
        End If
        
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
        crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
        
        If Option4.Value = True Then
            crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
            crystal.ParameterFields(5) = "@PL ;" + status + ";true"
        Else
            crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
            crystal.ParameterFields(6) = "@PL ;" + status + ";true"
        End If
        crystal.RetrieveDataFiles
        crystal.Action = 1
        
    ElseIf KeyCode = 112 Then
        'F1 Semua penjualan (P & L)
        Me.MousePointer = vbHourglass
        If txtinv4 < txtinv3 Then
            MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
            Exit Sub
        End If
        
        If Option1.Value = True Then str1 = "mall"
        If Option2.Value = True Then str1 = "mcust"
        If Option3.Value = True Then str1 = "marea"
        If Option4.Value = True Then str1 = "msales"
        status = ""
            
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowRefreshBtn = True
        
        crystal.Connect = dsnreport
        If Option1.Value = True Or Option2.Value = True Then
            crystal.DataFiles(0) = "Proc(am_monthly)"
        ElseIf Option3.Value = True Then
            crystal.DataFiles(0) = "Proc(am_monthlyall)"
        ElseIf Option4.Value = True Then
            crystal.DataFiles(0) = "Proc(am_daftarjualbulansale)"
        Else
        
        End If
        
        If Option1.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_all.rpt"
        ElseIf Option2.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl.rpt"
        ElseIf Option3.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlyall_area.rpt"
        ElseIf Option4.Value = True Then
            crystal.ReportFileName = AppPath & "\reports\sale\inv\monthlypl_sales.rpt"
        End If
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyy-MM-dd") & ";true"
        crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
        crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
        If Option4.Value = True Then
            crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
            crystal.ParameterFields(5) = "@PL ;" + status + ";true"
        Else
            crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
            crystal.ParameterFields(6) = "@PL ;" + status + ";true"
        End If
        crystal.RetrieveDataFiles
        crystal.Action = 1
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim thn, bln, bln2, tgl, tgl2 As String
    date_tahun = Date
    
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
End Sub

Sub disablein()
    txtinv3.Enabled = False
    txtinv4.Enabled = False
    cmdsearch3.Enabled = False
    cmdsearch4.Enabled = False
End Sub
Sub undisablein()
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    txtinv3 = ""
    txtinv4 = ""
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
End Sub

Private Sub Option1_Click()
    Call disablein
    caricust
End Sub

Private Sub Option2_Click()
    Call undisablein
End Sub

Private Sub Option3_Click()
    Call undisablein
End Sub

Private Sub caricust()
    OBJ.Open dsn
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust asc"
    Set RST = OBJ.Execute(SQL)
    txtinv3 = RST!kodecust
    
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust desc"
    Set RST = OBJ.Execute(SQL)
    txtinv4 = RST!kodecust
    OBJ.Close
End Sub

Private Sub Option4_Click()
    Call undisablein
End Sub
