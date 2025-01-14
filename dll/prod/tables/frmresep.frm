VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmresep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATUR FORMULA LEM DAN KARET"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4635
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   10410
      _Version        =   851970
      _ExtentX        =   18362
      _ExtentY        =   8176
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Bahan Baku"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "TabControlPage1"
      Item(0).Control(1)=   "page1"
      Item(1).Caption =   "Barang Jadi"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4275
         Left            =   30
         TabIndex        =   19
         Top             =   330
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         Page            =   2
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
            Height          =   4125
            Left            =   60
            TabIndex        =   21
            Top             =   45
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   7276
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
      End
      Begin XtremeSuiteControls.TabControlPage page1 
         Height          =   4275
         Left            =   -69970
         TabIndex        =   18
         Top             =   330
         Visible         =   0   'False
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         Page            =   1
         Begin TDBNumber6Ctl.TDBNumber txtnilai 
            Height          =   255
            Left            =   7605
            TabIndex        =   22
            Top             =   -720
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            Calculator      =   "frmresep.frx":0000
            Caption         =   "frmresep.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmresep.frx":008C
            Keys            =   "frmresep.frx":00AA
            Spin            =   "frmresep.frx":00EC
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.000;(###,###,###,##0.000);0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.000;(###,###,###,##0.000)"
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
         Begin VB.TextBox txtinisial 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   8235
            TabIndex        =   23
            Top             =   135
            Visible         =   0   'False
            Width           =   2010
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   4035
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   7117
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   4275
         Left            =   -69970
         TabIndex        =   17
         Top             =   330
         Visible         =   0   'False
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   0
      End
   End
   Begin VB.TextBox txtnoproduk 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   14
      Top             =   1290
      Width           =   2010
   End
   Begin VB.ComboBox cmbklasifikasi 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   930
      Width           =   2040
   End
   Begin VB.TextBox txtKlasifika 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   30
      TabIndex        =   2
      Top             =   7170
      Visible         =   0   'False
      Width           =   2010
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
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   180
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
      Left            =   10080
      Picture         =   "frmresep.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   255
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
      Left            =   9840
      Picture         =   "frmresep.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtKdProduk 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   195
      Width           =   2010
   End
   Begin VB.TextBox txtnmprod 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   555
      Width           =   5580
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   9405
      TabIndex        =   3
      Top             =   6660
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   -2147483633
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
      Left            =   7305
      TabIndex        =   4
      Top             =   6660
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Save"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton btnDelete 
      Height          =   465
      Left            =   8340
      TabIndex        =   5
      Top             =   6660
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Delete"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton btnNew 
      Height          =   465
      Left            =   6240
      TabIndex        =   6
      Top             =   6660
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "New"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton btnKodeProduk 
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1140
      _Version        =   851970
      _ExtentX        =   2011
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Kode Produk : "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   0
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   30
      TabIndex        =   15
      Top             =   1320
      Width           =   1320
      _Version        =   851970
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nomor Produk : "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   0
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   450
      Left            =   5175
      TabIndex        =   24
      Top             =   6660
      Visible         =   0   'False
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Cetak"
      BackColor       =   -2147483633
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
   Begin Crystal.CrystalReport crystal 
      Left            =   1215
      Top             =   6675
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
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   450
      Left            =   10560
      TabIndex        =   25
      Top             =   5805
      Width           =   3360
      _Version        =   851970
      _ExtentX        =   5927
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "UPLOAD GAMBAR SOP"
      BackColor       =   -2147483633
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
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   3960
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.PushButton cmdstok 
      Height          =   465
      Left            =   90
      TabIndex        =   26
      Top             =   6645
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Cek STOK"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton cmdotoritas 
      Height          =   465
      Left            =   1680
      TabIndex        =   27
      Top             =   6645
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Open Formula"
      BackColor       =   -2147483633
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
   Begin VB.Image image_photo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5565
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   210
      Width           =   3360
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Klasifika          :"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1005
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nama Produk  :"
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
      Left            =   90
      TabIndex        =   8
      Top             =   630
      Width           =   1200
   End
End
Attribute VB_Name = "frmResep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private RSStream As ADODB.Stream
Private SQL As String
Private flname As String
Private poscol As Integer
Private posrow As Integer
Private editmode As Boolean
Private datebahan As Date

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    On Error GoTo Err_handler:
    If MsgBox("Apakah anda yakin akan menghapus produk tersebut ..?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "delete from list_produk_master where kode_produk ='" & txtKdProduk & "'"
    OBJ.Execute SQL
    SQL = "delete from list_produk_child where kode_produk = '" & txtKdProduk & "'"
    OBJ.Execute SQL
    OBJ.Close
    MsgBox "Proses hapus berhasil...", vbInformation, AppName
    btnNew_Click
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Proses hapus tidak berhasil...!!, " + Err.Description, vbCritical, AppName
End Sub

Private Sub btnKodeProduk_Click()
    carisql1 = "select * from am_itemcode where (lev =3 or lev =4)"
    namatabel = "Produk"
    frmsearch.Show vbModal
End Sub

Private Sub btnKodeProduk_GotFocus()
    If hasil = "" Then Exit Sub
    txtKdProduk = hasil1
    txtnmprod = hasil2
    carisql1 = ""
    namatabel = ""
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hapusgrid
    opendata
End Sub

Private Sub opendata()
   ' On Error GoTo Err_handler:
    
    If OBJ.State = 0 Then
        OBJ.Open dsn
    End If
    
    SQL = "select * from list_produk_master where kode_produk='" & txtKdProduk & "'"
    Set RS = OBJ.Execute(SQL)
    If RS.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    editmode = True
    
    If Not RS.EOF Then
        cmbklasifikasi.text = RS!klasifikasi
        txtnoproduk.text = RS!nomor
        flname = AppPath & "\temp\" & RS!nomor & ".jpg"
        If Not IsNull(RS!gambar) Then
            Set RSStream = New ADODB.Stream
            RSStream.Type = adTypeBinary
            RSStream.Open
            RSStream.Write RS!gambar
            RSStream.SaveToFile flname, adSaveCreateOverWrite
            RSStream.Close
            Set RSStream = Nothing
            image_photo.Picture = LoadPicture(flname)
        End If
    End If
    
    SQL = "select a.*, b.namasatuan from list_produk_child a "
    SQL = SQL + "inner join am_apunit b on a.kode_satuan = b.kodesatuan "
    SQL = SQL + " where kode_produk='" & txtKdProduk & "' order by a.line "
    
    Set RS = OBJ.Execute(SQL)
    If RS.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    Do While Not RS.EOF
        grid.TextMatrix(grid.Row, 1) = RS!kode_bahan
        grid.TextMatrix(grid.Row, 2) = RS!nama_bahan
        grid.TextMatrix(grid.Row, 3) = RS!Inisial
        grid.TextMatrix(grid.Row, 4) = Format(RS!Qty, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 5) = RS!KODE_SATUAN
        grid.TextMatrix(grid.Row, 6) = RS!namasatuan
        grid.TextMatrix(grid.Row, 7) = RS!Line
        grid.Col = 0
        Set grid.CellPicture = uncheck
        SetAlternatingGrid grid.Row
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RS.MoveNext
    Loop
    
    SQL = "select  a.kode_produk,a.kode_barang_jadi,a.kode_satuan,"
    SQL = SQL + "b.namabarang ,c.namasatuan  "
    SQL = SQL + "from list_produk_hasil a "
    SQL = SQL + "inner join am_itemdtl b on a.kode_barang_jadi= b.kodebarang and a.kode_satuan=b.kodesatuan "
    SQL = SQL + "inner join am_unit c on a.kode_satuan = c.kodesatuan "
    SQL = SQL + "where a.kode_produk ='" & txtKdProduk & "'"
    Set RS = OBJ.Execute(SQL)
    hapusgrid1
    grid1.Row = 1
    Do While Not RS.EOF
        grid1.TextMatrix(grid1.Row, 1) = RS!Kode_barang_jadi
        grid1.TextMatrix(grid1.Row, 2) = RS!namabarang
        grid1.TextMatrix(grid1.Row, 3) = RS!KODE_SATUAN
        grid1.TextMatrix(grid1.Row, 4) = RS!namasatuan
        grid1.Col = 0
        grid1.Col = 0
        Set grid1.CellPicture = uncheck
        setAlternatingGrid1 grid1.Row
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RS.MoveNext
    Loop
    'periksa wip jadi
    SQL = "Select a.kode_produk,a.kode_barang_jadi,a.kode_satuan,b.NamaBarang,c.NamaSatuan"
    SQL = SQL + " From list_produk_hasil a"
    SQL = SQL + " inner join am_apitemmst b on a.kode_barang_jadi = b.KodeBarang"
    SQL = SQL + " inner join am_apunit c on a.kode_satuan = c.KodeSatuan"
    SQL = SQL + " where a.kode_produk ='" & txtKdProduk & "'"
    Set RS = OBJ.Execute(SQL)
    Do While Not RS.EOF
        grid1.TextMatrix(grid1.Row, 1) = RS!Kode_barang_jadi
        grid1.TextMatrix(grid1.Row, 2) = RS!namabarang
        grid1.TextMatrix(grid1.Row, 3) = RS!KODE_SATUAN
        grid1.TextMatrix(grid1.Row, 4) = RS!namasatuan
        grid1.Col = 0
        grid1.Col = 0
        Set grid1.CellPicture = uncheck
        setAlternatingGrid1 grid1.Row
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RS.MoveNext
    Loop
    
    If Not nmuser = "martsanto" Or nmuser = "Creator" Or nmuser = "angelgunawan" Or nmuser = "Angle" Or nmuser = "kimlie" Or nmuser = "putri" Or nmuser = "bina" Then
        'Periksa Otoritas
        SQL = "Select * From list_produk_masterkey"
        SQL = SQL + " Where kode_produk ='" & txtKdProduk & "' and otoritas = '1'"
        Set RS = OBJ.Execute(SQL)
        
        If RS.EOF Then
            MsgBox "Maaf Anda tidak memiliki otoritas, Silahkan hubungi Administrator Anda", vbCritical, AppName
            btnNew_Click
        End If
    End If
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Proses open data tidak berhasil...! " + Err.Description, vbCritical, AppName
End Sub

Private Sub btnNew_Click()
    txtKdProduk = ""
    txtnmprod = ""
    txtnoproduk = ""
    hapusgrid
    hapusgrid1
    editmode = False
    TabControl1.Item(0).Selected = True
    flname = ""
    image_photo.Picture = LoadPicture(flname)
End Sub

Private Sub btnSave_Click()
    On Error GoTo err_msg
    
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Akses ditolak...!", vbCritical, AppName
        Exit Sub
    End If
    
proses:

    'cek grid bahan baku
    grid.Row = 1
    If grid.TextMatrix(grid.Row, 1) = "" Then
        MsgBox "Bahan baku belum terisi...", vbCritical
        Exit Sub
    End If
    
    'cek barang jadi
    grid1.Row = 1
    If grid1.TextMatrix(grid1.Row, 1) = "" Then
        MsgBox "Barang Jadi/Perolehan belum terisi....!", vbCritical
        Exit Sub
    End If
    
    'cek gambar
    If flname = "" Then
        MsgBox "File gambar tidak ditemukan...!", vbCritical, AppName
        Exit Sub
    End If

    If txtKdProduk = "" Then
        MsgBox "Data Not Completed", vbCritical, AppName
        Exit Sub
    End If
    If txtnmprod = "" Then
        MsgBox "Data Not Complete", vbCritical, AppName
        Exit Sub
    End If
    If grid.TextMatrix(1, 1) = "" Then
        MsgBox "Data Not Completed", vbCritical, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    
    If editmode = True Then
        If MsgBox("Apakah anda yakin akan merubah data...?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
        SQL = "delete from list_produk_master where kode_produk='" & txtKdProduk & "'"
        OBJ.Execute SQL
        SQL = "delete from list_produk_child where kode_produk ='" & txtKdProduk & "'"
        OBJ.Execute SQL
        SQL = "delete from list_produk_hasil where kode_produk ='" & txtKdProduk & "'"
        OBJ.Execute SQL
        'save to produk

        SQL = "select * from list_produk_master where 0=1"
        Set RS = New ADODB.Recordset
        RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        With RS
            .AddNew
            !KODE_PRODUK = txtKdProduk
            !NAMA_PRODUK = txtnmprod
            !flag_status = "1"
            !klasifikasi = cmbklasifikasi.text
            !nomor = txtnoproduk
            !UserName = nmuser
            If flname <> "" Then
                Set RSStream = New ADODB.Stream
                RSStream.Type = adTypeBinary
                RSStream.Open
                RSStream.LoadFromFile flname
                !gambar = RSStream.Read
                RSStream.Close
            End If
            .Update
        End With
    
        'save to produk child
        SQL = "select * from list_produk_child where 0=1 "
        Set RS = New ADODB.Recordset
        RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            RS.AddNew
            RS!KODE_PRODUK = txtKdProduk
            RS!kode_bahan = grid.TextMatrix(grid.Row, 1)
            RS!nama_bahan = grid.TextMatrix(grid.Row, 2)
            RS!Inisial = grid.TextMatrix(grid.Row, 3)
            RS!KODE_SATUAN = grid.TextMatrix(grid.Row, 5)
            RS!Qty = Format(grid.TextMatrix(grid.Row, 4), "general number")
            RS!Line = Format(grid.TextMatrix(grid.Row, 7), "general number")
            RS.Update
            grid.Row = grid.Row + 1
        Loop
    
        'save to produk hasil
        SQL = "select * from list_produk_hasil where 0=1"
        Set RS = New ADODB.Recordset
        RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid1.Row = 1
        Do While True
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            RS.AddNew
            RS!KODE_PRODUK = txtKdProduk
            RS!Kode_barang_jadi = grid1.TextMatrix(grid1.Row, 1)
            RS!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 3)
            RS!baris = grid1.Row
            RS.Update
            grid1.Row = grid1.Row + 1
        Loop
        
        'Kunci Kembali Otoritas Update
        SQL = "Update list_produk_masterkey set otoritas = '0',keterangan='" & nmuser & "'"
        SQL = SQL + " Where kode_produk='" & txtKdProduk & "' and otoritas='1'"
        Set RS = OBJ.Execute(SQL)
        
        OBJ.Close
        MsgBox "Data is save....", vbInformation, AppName
        btnNew_Click
        Exit Sub
    End If
    
    
    'validasi produk
    SQL = "select * from list_produk_master where kode_produk ='" & txtKdProduk & "'"
    Set RS = OBJ.Execute(SQL)
    If Not RS.EOF Then
        OBJ.Close
        MsgBox "Data produk telah ada", vbInformation, AppName
        Exit Sub
    End If
    
    'save to produk

    SQL = "select * from list_produk_master where 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RS
        .AddNew
        !KODE_PRODUK = txtKdProduk
        !NAMA_PRODUK = txtnmprod
        !flag_status = "1"
        !klasifikasi = cmbklasifikasi.text
        !nomor = txtnoproduk
        !UserName = nmuser
        If flname <> "" Then
            Set RSStream = New ADODB.Stream
            RSStream.Type = adTypeBinary
            RSStream.Open
            RSStream.LoadFromFile flname
            !gambar = RSStream.Read
            RSStream.Close
        End If
        
        .Update
    End With
    
    'save to produk child
    SQL = "select * from list_produk_child where 0=1 "
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        RS.AddNew
        RS!KODE_PRODUK = txtKdProduk
        RS!kode_bahan = grid.TextMatrix(grid.Row, 1)
        RS!nama_bahan = grid.TextMatrix(grid.Row, 2)
        RS!Inisial = grid.TextMatrix(grid.Row, 3)
        RS!KODE_SATUAN = grid.TextMatrix(grid.Row, 5)
        RS!Qty = Format(grid.TextMatrix(grid.Row, 4), "general number")
        RS!Line = Format(grid.TextMatrix(grid.Row, 7), "general number")
        RS.Update
        grid.Row = grid.Row + 1
    Loop
    
    'save to produk hasil
    SQL = "select * from list_produk_hasil where 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RS.AddNew
        RS!KODE_PRODUK = txtKdProduk
        RS!Kode_barang_jadi = grid1.TextMatrix(grid1.Row, 1)
        RS!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 3)
        RS!baris = grid1.Row
        RS.Update
        grid1.Row = grid1.Row + 1
    Loop
    
    OBJ.Close
    MsgBox "Data is save....", vbInformation, AppName
    btnNew_Click
    Exit Sub
err_msg:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdotoritas_Click()
    If nmuser = "martsanto" Or nmuser = "Creator" Or nmuser = "angelgunawan" Or nmuser = "Angle" Or nmuser = "kimlie" Or nmuser = "putri" Or nmuser = "bina" Then
        frmotoritas.Show vbModal
    End If
End Sub

Private Sub cmdstok_Click()
    Dim stokbahan As Double
    Dim namabahan As String
    Dim namasatuan As String
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        setAlternatingGrid1Yelow grid.Row
        GetStokBarang Format(Date, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), namabahan, namasatuan, stokbahan
        If stokbahan <= 0 Or stokbahan < Format(grid.TextMatrix(grid.Row, 4), "general number") Then
            MsgBox "Nama Bahan Baku : " & namabahan & Chr(13) & _
            " Stok Terkahir : " & stokbahan & " " & namasatuan, vbCritical, "Peringatan Stok"
            setAlternatingGrid1Red grid.Row
        Else
            SetAlternatingGrid grid.Row
        End If
        grid.Row = grid.Row + 1
    Loop
    MsgBox "Proses cek stok bahan baku selesai...!", vbInformation, AppName
End Sub

Private Sub Form_Load()
    'init form
    txtKdProduk.TabIndex = 0
    txtnmprod.TabIndex = 1
    cmbklasifikasi.TabIndex = 2
    txtnoproduk.TabIndex = 3
    datebahan = Date
    'set grid
    With grid
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode Bahan"
        .TextMatrix(0, 2) = "Nama Bahan"
        .TextMatrix(0, 3) = "Inisial"
        .TextMatrix(0, 4) = "Qty"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "Baris"
        
    End With
    setGrid
    cmbklasifikasi.AddItem "PU Base"
    cmbklasifikasi.AddItem "CR"
    cmbklasifikasi.AddItem "CR-GRAFT"
    cmbklasifikasi.AddItem "NR"
    cmbklasifikasi.AddItem "PU"
    cmbklasifikasi.AddItem "Primer"
    cmbklasifikasi.AddItem "PVC"
    cmbklasifikasi.AddItem "KARET"
    cmbklasifikasi.AddItem "WB"
    
    'set grid
    With grid1
        .Cols = 5
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode Barang"
        .TextMatrix(0, 2) = "Barang"
        .TextMatrix(0, 3) = "Kode Satuan"
        .TextMatrix(0, 4) = "Satuan"
    End With
    setgrid1
    
    TabControl1.Item(0).Selected = True
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 250
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(7) = 500
    End With
End Sub

Private Sub setgrid1()
    With grid1
        .ColWidth(0) = 250
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 800
        .ColWidth(4) = 1200
    End With
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtKdProduk = "" Then Exit Sub
    
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
            Case 7:
                    txtnilai.Width = grid.ColWidth(grid.Col) - 40
                    txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                    txtnilai.Left = grid.Left + grid.CellLeft
                    txtnilai.Top = grid.Top + grid.CellTop + 20
                    txtnilai.Visible = True
                    txtnilai.SetFocus
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
            
            If stokbahan <= 0 Then
                MsgBox "Nama Bahan Baku : " & hasil1 & Chr(13) & _
                "Stok Terakhir : " & Format(stokbahan, "##,###,###,##0.000") & " " & nama_satuan
                hasil = ""
                hasil1 = ""
                carisql1 = ""
                namatabel = ""
                Exit Sub
            End If
            With grid
                '.Row = 1
                'Do While True
                '    If .TextMatrix(.Row, 1) = "" Then Exit Do
                '    If .TextMatrix(.Row, 1) = hasil Then
                '        MsgBox "All ready exist...!", vbCritical, AppName
                '        hasil = ""
                '        hasil1 = ""
                '        Exit Sub
                '    End If
                '    .Row = .Row + 1
                'Loop
                .TextMatrix(.Row, 1) = hasil
                .TextMatrix(.Row, 2) = hasil1
                .TextMatrix(.Row, 4) = "0.000"
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
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
             grid.TextMatrix(grid.Row, 6) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtKdProduk = "" Then Exit Sub
    
    poscol = grid1.Col
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 0:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                If grid1.CellPicture = uncheck Then
                Set grid1.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid1.CellPicture = uncheck
                    hapusrow1
                    Exit Sub
                End If
                Set grid1.CellPicture = uncheck
                End If
        Case 1:
            'If Check1.Value = 1 Then
                'If grid1.TextMatrix(grid1.Row, 1) <> "" Then Exit Sub
                'carisql1 = "select a.KodeBarang,a.NamaBarang,a.KodeSatuan,b.NamaSatuan from am_apitemmst a "
                'carisql1 = carisql1 + "inner join am_apunit b on a.KodeSatuan = b.KodeSatuan and KodeBarang like 'L11.%'"
                'namatabel = "WIP Jadi"
            'ElseIf Check1.Value = 0 Then
            If grid1.TextMatrix(grid1.Row, 1) <> "" Then Exit Sub
            carisql1 = "select a.kodebarang,a.namabarang,a.kodesatuan,b.namasatuan from am_itemdtl a "
            carisql1 = carisql1 + "inner join am_unit b on a.kodesatuan= b.kodesatuan "
            namatabel = "Barang Jadi"
            'End If
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    grid1.TextMatrix(grid1.Row, 1) = hasil
    grid1.TextMatrix(grid1.Row, 2) = hasil1
    grid1.TextMatrix(grid1.Row, 3) = hasil2
    'If Check1.Value = 1 Then grid1.TextMatrix(grid1.Row, 4) = hasil3
    carisql1 = ""
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hasil3 = ""
    
    OBJ.Open dsn
    SQL = "select a.*,b.namasatuan from am_itemdtl a inner join am_unit b on a.kodesatuan = b.kodesatuan "
    SQL = SQL + "where a.kodebarang ='" & grid1.TextMatrix(grid1.Row, 1) & "' and  a.kodesatuan='" & grid1.TextMatrix(grid1.Row, 3) + "'"
    Set RS = OBJ.Execute(SQL)
    If Not RS.EOF Then
        grid1.TextMatrix(grid1.Row, 4) = RS!namasatuan
    End If
    grid1.Col = 0
    Set grid1.CellPicture = uncheck
    setAlternatingGrid1 grid1.Row
    grid1.Rows = grid1.Rows + 1
    grid1.Row = grid1.Row + 1
    OBJ.Close
End Sub

Private Sub PushButton2_Click()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetaksop)"
    crystal.ReportFileName = AppPath & "\reports\produksi\1014.rpt"
    crystal.ParameterFields(0) = "@kode_produk;" & txtKdProduk & ";true"
    crystal.ParameterFields(1) = "@username;" & nmuser & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub PushButton3_Click()
    On Error GoTo err_msg
    With cmndlg
        .CancelError = False
        .DialogTitle = "Gambar SOP"
        .Filter = "File Gambar (*.jpg)|*.jpg"
        .ShowOpen
        If .FileName <> "" Then
            flname = .FileName
            image_photo.Picture = LoadPicture(flname)
        Else
            image_photo.Picture = LoadPicture(flname)
            Exit Sub
        End If
    End With
    Exit Sub
err_msg:
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
    Dim stokbahanbaku As Double
    Dim nama_bahan As String
    Dim nama_satuan As String
    Dim d As Date
    
    If KeyAscii = 13 Then
        If grid.Col = 4 Then
        d = DateAdd("d", 1, datebahan)
        'GetStokBarang Format(d, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), , , stokbahanbaku
        GetStokBarang Format(d, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), nama_bahan, nama_satuan, stokbahanbaku
        If stokbahanbaku <= 0 Or stokbahanbaku < Format(txtnilai.Value, "general number") Then
            MsgBox "Nama Bahan Baku : " & nama_bahan & Chr(13) & _
             "Stok Terakhir : " & Format(stokbahanbaku, "##,###,###,##0.0000") & " " & nama_satuan, vbCritical, "Peringatan Stok"
            'Exit Sub
        End If
        grid.TextMatrix(grid.Row, 4) = txtnilai.text
        grid.SetFocus
            
        Else
        grid.TextMatrix(grid.Row, 7) = txtnilai.Value
        
        grid.SetFocus
        End If
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
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

Private Function setAlternatingGrid1(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If (i Mod 2) = 0 Then
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HFFFFFF
        Next
    End If
End Function


Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.Col = 0
        Set grid1.CellPicture = blank
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    setGrid
End Sub

Private Sub hapusrow1()
    grid1.TextMatrix(grid1.Row, 1) = ""
    grid1.TextMatrix(grid1.Row, 2) = ""
    grid1.TextMatrix(grid1.Row, 3) = ""
    grid1.TextMatrix(grid1.Row, 4) = ""
    
    Do While True
        If grid1.TextMatrix(grid1.Row + 1, 1) = "" Then
            grid1.TextMatrix(grid1.Row, 1) = ""
            grid1.TextMatrix(grid1.Row, 2) = ""
            grid1.TextMatrix(grid1.Row, 3) = ""
            grid1.TextMatrix(grid1.Row, 4) = ""
            Exit Do
        End If
        grid1.TextMatrix(grid1.Row, 1) = grid1.TextMatrix(grid1.Row + 1, 1)
        grid1.TextMatrix(grid1.Row, 2) = grid1.TextMatrix(grid1.Row + 1, 2)
        grid1.TextMatrix(grid1.Row, 3) = grid1.TextMatrix(grid1.Row + 1, 3)
        grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row + 1, 4)
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = grid1.Rows - 1
    grid1.Col = 0
    Set grid1.CellPicture = blank
End Sub


Private Function setAlternatingGrid1Red(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = vbRed
    Next
End Function

Private Function setAlternatingGrid1Yelow(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = vbYellow
    Next
End Function

