VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlaporanpembelian 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Penerimaan + Retur"
   ClientHeight    =   2760
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
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
      TabIndex        =   13
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cmbkode 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   130809859
      CurrentDate     =   38679
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmlaporanpembelian.frx":0000
      Left            =   1320
      List            =   "frmlaporanpembelian.frx":000D
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   5640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtkode2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   2250
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
      MICON           =   "frmlaporanpembelian.frx":0030
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
      Left            =   4080
      TabIndex        =   6
      Top             =   2250
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
      MICON           =   "frmlaporanpembelian.frx":034A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Kode"
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
      MICON           =   "frmlaporanpembelian.frx":0664
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Kode"
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
      MICON           =   "frmlaporanpembelian.frx":097E
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
      Height          =   300
      Left            =   4080
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   170721283
      CurrentDate     =   38679
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "Sub Divisi"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Group By"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From Date                                                     To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1830
      Width           =   3735
   End
End
Attribute VB_Name = "frmlaporanpembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str2 As String

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        cmbkode.Enabled = True
        Combo1.Enabled = True
        txtkode1.Enabled = True
        txtkode2.Enabled = True
        cmdsearch1.Enabled = True
        cmdsearch2.Enabled = True
        cmbkode.SetFocus
    Else
        cmbkode = ""
        Combo1 = ""
        txtkode1 = ""
        txtkode2 = ""
        
        cmbkode.Enabled = False
        Combo1.Enabled = False
        txtkode1.Enabled = False
        txtkode2.Enabled = False
        cmdsearch1.Enabled = False
        cmdsearch2.Enabled = False
    End If
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo1_Click()
    If Combo1 = "Tanggal" Then
        cmdsearch1.Enabled = False
        cmdsearch2.Enabled = False
        txtkode1 = ""
        txtkode2 = ""
        txtkode1.Enabled = False
        txtkode2.Enabled = False
        date1.Enabled = True
        date2.Enabled = True
        date1 = Date
        date2 = Date
    ElseIf Combo1 = "Supplier" Then
        cmdsearch1.Enabled = True
        cmdsearch2.Enabled = False
        txtkode1 = ""
        txtkode2 = ""
        txtkode1.Enabled = True
        txtkode2.Enabled = False
        date1 = Date
        date2 = Date
        date1.Enabled = False
        date2.Enabled = False
    Else
        cmdsearch1.Enabled = True
        cmdsearch2.Enabled = True
        txtkode1 = ""
        txtkode2 = ""
        txtkode1.Enabled = True
        txtkode2.Enabled = True
        date1 = Date
        date2 = Date
        date1.Enabled = False
        date2.Enabled = False
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    If Check1.Value = 0 Then
        If Combo1 = "" Or cmbkode = "" Then
            MsgBox "Data entry not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If txtkode1 > txtkode2 And Combo1 = "Bahan Baku" Then
            MsgBox "From Kode Greather To Kode.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If txtkode1 = "" And Combo1 = "Supplier" Then
            MsgBox "Data entry not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If date1 > date2 Then
            MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If txtkode1 = "" And Combo1 = "Bahan Baku" Then txtkode1 = "0"
        If txtkode2 = "" And Combo1 = "Bahan Baku" Then txtkode2 = "z"
        
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_laporbeli_detail)"
        
        If Combo1 = "Supplier" Then Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\pembelian_detail_supplier.rpt"
        If Combo1 = "Tanggal" Then Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\pembelian_detail_tanggal.rpt"
        If Combo1 = "Bahan Baku" Then Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\pembelian_detail_barang.rpt"
        
        If Combo1 = "Tanggal" Then
            Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
            Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
        Else
            Crystal.ParameterFields(0) = "@kode1;" + txtkode1 + ";true"
            Crystal.ParameterFields(1) = "@kode2;" + txtkode2 + ";true"
        End If
        Crystal.ParameterFields(2) = "@pilih;" + Combo1 + ";true"
        Crystal.ParameterFields(3) = "@pilih1;" + cmbkode + ";true"
        Crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
        
        txtkode1 = ""
        txtkode2 = ""
    Else
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
        Crystal.DataFiles(0) = "Proc(am_laporbeli_retur)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\pembelian_retur.rpt"
        Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    If Combo1 = "Supplier" Then
        carisql1 = "select namasupp, AlamatSupp1,kodesupp from am_supplier"
        namatabel = "Supplier"
    Else
        carisql1 = "select kodebarang, namabarang from am_apitemmst"
        namatabel = "Bahan Baku"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    
    If Combo1 = "Supplier" Then txtkode1 = hasil2 Else txtkode1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
        
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub Form_Load()
    
    date1 = Date
    date2 = Date
    
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            RST.MoveNext
        Loop
    End If
    OBJ.Close
      
End Sub

Private Sub txtkode1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtkode2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkode2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
