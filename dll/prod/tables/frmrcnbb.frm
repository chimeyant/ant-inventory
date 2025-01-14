VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmrcnbb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rencana Pemakaian Bahan Baku"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbstatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11760
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtinisial 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      Picture         =   "frmrcnbb.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      Picture         =   "frmrcnbb.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   12000
      TabIndex        =   0
      Top             =   6720
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
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
   Begin XtremeSuiteControls.PushButton btnshow 
      Height          =   345
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Show"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmrcnbb.frx":0630
      Caption         =   "frmrcnbb.frx":0650
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrcnbb.frx":06BC
      Keys            =   "frmrcnbb.frx":06DA
      Spin            =   "frmrcnbb.frx":071C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   5475
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   9657
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   12855
      _Version        =   851970
      _ExtentX        =   22675
      _ExtentY        =   450
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin XtremeSuiteControls.PushButton btnAdd 
      Height          =   465
      Left            =   120
      TabIndex        =   21
      Top             =   6720
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Add Kategori"
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
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   465
      Left            =   11040
      TabIndex        =   25
      Top             =   6720
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Save"
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
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perhatian : Pada Mode Edit ini, data Raw material usage  plan sebelumnya telah dihapus"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   26
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblnobb 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   11400
      TabIndex        =   24
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Kg"
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
      Left            =   10320
      TabIndex        =   23
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Total Bahan Baku :"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lbltotalkg 
      Alignment       =   2  'Center
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
      Left            =   9120
      TabIndex        =   18
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblsatuan 
      BackStyle       =   0  'Transparent
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
      Left            =   7320
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblqty 
      BackStyle       =   0  'Transparent
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
      Left            =   6480
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Rencana Produksi :"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblnamabrg 
      BackStyle       =   0  'Transparent
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
      Left            =   2400
      TabIndex        =   12
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblkode 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Barang Jadi"
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
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblnamaproduk 
      BackStyle       =   0  'Transparent
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
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "PRODUK"
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
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblnorcn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   11400
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblkdproduk 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmrcnbb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RS      As ADODB.Recordset
Private OBJ     As New ADODB.Connection
Private SQL     As String
Private RST     As ADODB.Recordset
Private OBJ1    As New ADODB.Connection
Private SQL1    As String

Private poscol As Integer
Private posrow As Integer
Private datebahan As Date
Dim intktgori As Integer

Private Sub btnAdd_Click()
    frmkategori.Show vbModal
End Sub

Private Sub btnAdd_GotFocus()
    cmbstatus.Clear
    OBJ.Open dsn
    SQL = "SELECT * FROM am_kategori"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        cmbstatus.AddItem RST!kategori
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub btnClose_Click()
    If hasil4 = "" Then
        Unload Me
    Else
        If MsgBox("Apakah No.BB : " & hasil4 & " akan dihapus ?" & vbLf & "Klik Yes untuk hapus kode Dan klik No untuk lanjut input bahan baku", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            hasil = "hapus"
            Unload Me
        End If
    End If
End Sub

Private Sub btnSave_Click()
    If lblnobb = "" Or lblnorcn = "" Then Exit Sub
    If grid.TextMatrix(1, 1) = "" Then Exit Sub
    'PERIKSA KATEGORI
    Call cektgori
    If intktgori <> 1 Then
        MsgBox "Kategori belum lengkap", vbCritical, AppName
        Exit Sub
    End If

    If MsgBox("Apakah data sudah diisi dengan benar" & vbLf & _
    "Klik Yes untuk menyimpan data", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    'SIMPAN KE TABEL am_rcnpack
    OBJ.Open dsn
    SQL = "Select * From am_rcnbb Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
   
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!No_RCNBB = lblnobb
        RST!Kd_RCN = lblnorcn
        RST!Kd_Produk = lblkdproduk
        RST!Kd_Barang = lblkode
        RST!Kd_bahan = grid.TextMatrix(grid.Row, 1)
        RST!Inisial = grid.TextMatrix(grid.Row, 3)
        RST!Qty = grid.TextMatrix(grid.Row, 4)
        RST!Stok = grid.TextMatrix(grid.Row, 7)
        RST!Kdktgori = grid.TextMatrix(grid.Row, 10)
        RST!kategori = grid.TextMatrix(grid.Row, 9)
        RST!tgl = Date
        RST!Line = grid.Row
        RST!hpp = "0"
        RST.Update
        If grid.Rows = grid.Row + 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    'If grid.TextMatrix(grid.Row, 1) = "" Then
        'hasil = ""
    'Else
        hasil = lblnobb
        'lblkdpack = ""
    'End If
    Unload Me
End Sub

Private Sub btnshow_Click()
    opendata
    btnshow.Enabled = False
End Sub

Private Sub cmbstatus_Click()
    grid.TextMatrix(grid.Row, 9) = cmbstatus
    OBJ.Open dsn
    SQL = "Select * From am_kategori Where kategori='" & grid.TextMatrix(grid.Row, 9) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        grid.TextMatrix(grid.Row, 10) = RST!kdkategori
    End If
    
    'SIMPAN PRODUK KATEGORI
    SQL = "Select * From am_kategoribahan Where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    If RST.EOF Then
        'SIMPAN
        RST.AddNew
        RST!kodebarang = grid.TextMatrix(grid.Row, 1)
        RST!kdkategori = grid.TextMatrix(grid.Row, 10)
        RST!kategori = grid.TextMatrix(grid.Row, 9)
    Else
        'UPDATE
        RST!kdkategori = grid.TextMatrix(grid.Row, 10)
        RST!kategori = grid.TextMatrix(grid.Row, 9)
    End If
    RST.Update
    OBJ.Close
    grid.SetFocus
End Sub

Private Sub cmbstatus_LostFocus()
    cmbstatus.Visible = False
End Sub

Private Sub Form_Load()
    With grid
        .Cols = 11
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode Bahan"
        .TextMatrix(0, 2) = "Nama Bahan"
        .TextMatrix(0, 3) = "Inisial"
        .TextMatrix(0, 4) = "Qty"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "Stok"
        .TextMatrix(0, 8) = "Baris"
        .TextMatrix(0, 9) = "Kategori"
        .TextMatrix(0, 10) = "Kode"
    End With
    setGrid
    opendata
    
    cmbstatus.Clear
    OBJ.Open dsn
    SQL = "SELECT * FROM am_kategori"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        cmbstatus.AddItem RST!kategori
        RST.MoveNext
    Loop
    OBJ.Close
    
    If hasil4 <> "" Then
        'EDIT MODE (Data dihapus dulu)
        'hasil4 berisikan data kode bb sebagai parameter jika edit mode dicancel
        Call hapusdata
        lblinfo.Visible = True
    Else
        'Add Mode
        lblnobb = getkdbb
        lblinfo.Visible = False
    End If
End Sub

Sub hapusdata()
    OBJ.Open dsn
    SQL = "DELETE FROM am_rcnbb WHERE No_RCNBB='" & hasil4 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub opendata()
    On Error GoTo Err_handler:
    Dim stokbahan As Double
    OBJ.Open dsn
    
    SQL = "Select COUNT(kode_bahan)'jml' from list_produk_child Where kode_produk='" & lblkdproduk & "'"
    Set RS = OBJ.Execute(SQL)
    Pg.Max = RS!jml
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "select a.*, b.namasatuan from list_produk_child a "
    SQL = SQL + "inner join am_apunit b on a.kode_satuan = b.kodesatuan "
    SQL = SQL + " where kode_produk='" & lblkdproduk & "' order by a.line "
    
    Set RS = OBJ.Execute(SQL)
    If RS.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    Do While Not RS.EOF
        grid.TextMatrix(grid.Row, 1) = RS!kode_bahan
        grid.TextMatrix(grid.Row, 2) = RS!nama_bahan
        grid.TextMatrix(grid.Row, 3) = RS!Inisial
        grid.TextMatrix(grid.Row, 4) = Format(RS!Qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = RS!KODE_SATUAN
        grid.TextMatrix(grid.Row, 6) = RS!namasatuan
        GetStokBarang Format(Date, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), , , stokbahan
        grid.TextMatrix(grid.Row, 7) = Format(stokbahan, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 8) = RS!Line
        'CEK KATEGORI
        OBJ1.Open dsn
        SQL1 = "Select distinct kdkategori,kategori From am_kategoribahan Where kodebarang='" & RS!kode_bahan & "'"
        Set RST = OBJ1.Execute(SQL1)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 9) = RST!kategori
            grid.TextMatrix(grid.Row, 10) = RST!kdkategori
        End If
        OBJ1.Close
        grid.Col = 0
        Set grid.CellPicture = uncheck
        SetAlternatingGrid grid.Row
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        RS.MoveNext
    Loop
    Call Totalkg
    Pg.Visible = False
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Proses open data tidak berhasil...! " + Err.Description, vbCritical, AppName
End Sub


Private Function SetAlternatingGrid(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If (i Mod 2) = 0 Then
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
        Next
    End If
End Function

Private Sub setbaris()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 8) = grid.Row
        grid.Row = grid.Row + 1
    Loop
End Sub
Private Sub cektgori()
    Dim i As Integer
    i = 0
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        If grid.TextMatrix(grid.Row, 9) <> "" Then
            i = i + 1
        End If
        grid.Row = grid.Row + 1
    Loop
    intktgori = grid.Row - i
End Sub
Private Sub Totalkg()
On Error Resume Next
    grid.Row = 1
    tkg = 0
    Do While True
        DoEvents
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        tkg = CDbl(Format(grid.TextMatrix(grid.Row, 4), "general number") + CDbl(tkg))
        grid.Row = grid.Row + 1
    Loop
    tkg = Format(tkg, "##,###,##0.00")
    lbltotalkg = tkg
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    grid.TextMatrix(grid.Row, 10) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            grid.TextMatrix(grid.Row, 10) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.TextMatrix(grid.Row, 10) = grid.TextMatrix(grid.Row + 1, 10)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.TextMatrix(grid.Row, 10) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 250
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(7) = 1200
        .ColWidth(8) = 800
        .ColWidth(9) = 1200
        .ColWidth(10) = 0
    End With
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If lblkdproduk = "" Then Exit Sub
    
    poscol = grid.Col
    posrow = grid.Row
    
    With grid
        Select Case .Col
            Case 0:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    setbaris
                    Totalkg
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
                End If
                
            Case 1:
                    If .TextMatrix(.Row, 1) <> "" Then Exit Sub
                    carisql1 = "select kodebarang, namabarang from am_apitemmst"
                    namatabel = "Bahan Baku"
                    frmsearch.Show vbModal
            Case 3:
                    If .TextMatrix(.Row, 1) = "" Then Exit Sub
                    txtinisial.Width = grid.ColWidth(grid.Col) - 40
                    txtinisial = grid.TextMatrix(grid.Row, grid.Col)
                    txtinisial.Left = grid.Left + grid.CellLeft
                    txtinisial.Top = grid.Top + grid.CellTop
                    txtinisial.Visible = True
                    txtinisial.SetFocus
            Case 4:
                    txtnilai.Width = grid.ColWidth(grid.Col) - 40
                    txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                    txtnilai.Left = grid.Left + grid.CellLeft
                    txtnilai.Top = grid.Top + grid.CellTop + 20
                    txtnilai.Visible = True
                    txtnilai.SetFocus
            Case 5:
                    If .TextMatrix(.Row, 1) = "" Then Exit Sub
                    carisql1 = "select kodesatuan, namasatuan, initial from am_apunit"
                    namatabel = "Satuan"
                    frmsearch.Show vbModal
            Case 8:
                    txtnilai.Width = grid.ColWidth(grid.Col) - 40
                    txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                    txtnilai.Left = grid.Left + grid.CellLeft
                    txtnilai.Top = grid.Top + grid.CellTop + 20
                    txtnilai.Visible = True
                    txtnilai.SetFocus
            Case 9: 'PERIKSA KATEGORI BAHAN BAKU
                    If .TextMatrix(.Row, 1) = "" Then Exit Sub
                    cmbstatus.Width = grid.ColWidth(grid.Col) - 40
                    cmbstatus = grid.TextMatrix(grid.Row, grid.Col)
                    cmbstatus.Left = grid.Left + grid.CellLeft
                    cmbstatus.Top = grid.Top + grid.CellTop + 20
                    cmbstatus.Visible = True
                    cmbstatus.SetFocus
        End Select
    End With
End Sub

Private Sub grid_GotFocus()
    Dim stokbahan As Double
    Dim kode_bahan As String
    Dim nama_bahan As String
    Dim nama_satuan As String
    
    If hasil = "" Then Exit Sub
    Select Case grid.Col
        Case 1:
            kode_bahan = hasil
            
            GetStokBarang Format(Date, "yyyyMMdd"), kode_bahan, nama_bahan, nama_satuan, stokbahan
            With grid
                .TextMatrix(.Row, 1) = hasil
                .TextMatrix(.Row, 2) = hasil1
                .TextMatrix(.Row, 4) = "0.000"
                .TextMatrix(.Row, 7) = stokbahan
                .Col = 0
                Set .CellPicture = uncheck
                SetAlternatingGrid grid.Row
                .Rows = .Rows + 1
                hasil = ""
                hasil1 = ""
                carisql1 = ""
                namatabel = ""
            End With
        Case 5:
            grid.TextMatrix(grid.Row, 5) = hasil
            grid.TextMatrix(grid.Row, 6) = hasil1
            carisql1 = ""
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub txtinisial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtinisial = "" Then Exit Sub
        grid.TextMatrix(grid.Row, 3) = txtinisial
        grid.SetFocus
    End If
End Sub

Private Sub txtinisial_LostFocus()
    txtinisial.Visible = False
End Sub
Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim nama_bahan As String
    Dim nama_satuan As String
    Dim d As Date
    
    If KeyAscii = 13 Then
        If grid.Col = 4 Then
        d = DateAdd("d", 1, datebahan)
        GetStokBarang Format(Date, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), , , stokbahan
        
        grid.TextMatrix(grid.Row, 4) = txtnilai.text
        grid.TextMatrix(grid.Row, 7) = Format(stokbahan, "##,###,###,##0.00")
        grid.SetFocus
            
        Else
        grid.TextMatrix(grid.Row, 8) = txtnilai.Value
        
        grid.SetFocus
        End If
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Function getkdbb() As String    'BB2110001
'On Error GoTo Err_handler:
    Dim strformat As String
    strformat = Format(Date, "yymm")
    
    Dim str99 As String
    Dim no As String
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select top 1 No_RCNBB from am_rcnbb"
    SQL = SQL + " where No_RCNBB like 'BB' + '" + strformat + "%' order by No_RCNBB desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!No_RCNBB, 3)
    Else
        str99 = 0
    End If
        str99 = str99 + 1
        
    If Len(str99) = 1 Then no = "BB" & strformat & "00" & str99
    If Len(str99) = 2 Then no = "BB" & strformat & "0" & str99
    If Len(str99) = 3 Then no = "BB" & strformat & str99
        
    getkdbb = no
    OBJ.Close
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function
