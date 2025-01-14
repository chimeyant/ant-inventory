VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmrcnprod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rencana Produksi"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   14925
   StartUpPosition =   1  'CenterOwner
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmrcnprod.frx":0000
      Caption         =   "frmrcnprod.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrcnprod.frx":008C
      Keys            =   "frmrcnprod.frx":00AA
      Spin            =   "frmrcnprod.frx":00EC
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
      Left            =   735
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   6000
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
      Left            =   375
      Picture         =   "frmrcnprod.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   6000
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
      Left            =   120
      Picture         =   "frmrcnprod.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   13800
      TabIndex        =   0
      Top             =   5640
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4920
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   8678
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   8421504
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin XtremeSuiteControls.DateTimePicker dtpfrom 
      Height          =   315
      Left            =   735
      TabIndex        =   6
      Top             =   240
      Width           =   1515
      _Version        =   851970
      _ExtentX        =   2672
      _ExtentY        =   556
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
      Format          =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpto 
      Height          =   315
      Left            =   2895
      TabIndex        =   8
      Top             =   240
      Width           =   1515
      _Version        =   851970
      _ExtentX        =   2672
      _ExtentY        =   556
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
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   465
      Left            =   12840
      TabIndex        =   12
      Top             =   5640
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11520
      TabIndex        =   14
      Top             =   5760
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9240
      TabIndex        =   13
      Top             =   5760
      Width           =   675
   End
   Begin VB.Label lblkdrcn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13080
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbltotalkg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "To  :"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   255
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "From  :"
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
      TabIndex        =   7
      Top             =   240
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   8880
      Shape           =   4  'Rounded Rectangle
      Top             =   5715
      Width           =   2895
   End
End
Attribute VB_Name = "frmrcnprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private SQL As String
Dim bln As String
Dim thn As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 900
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
        .ColWidth(4) = 900
        .ColWidth(5) = 3500
        .ColWidth(6) = 0
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 500
        .ColWidth(10) = 500
        .ColWidth(11) = 1500
        .ColWidth(12) = 1500
        .ColWidth(13) = 1500
    End With
End Sub

Private Sub initGrid()
    With grid
        .Cols = 14
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"          'KODE PRODUK
        .Col = 1
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 2) = "NAMA PRODUK"
        .TextMatrix(0, 3) = "KODE"          'SATUAN
        .TextMatrix(0, 4) = "KODE"          'KODE BARANG
        .Col = 4
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 5) = "NAMA BARANG"
        .TextMatrix(0, 6) = "KODE"          'KODE SATUAN
        .Col = 7
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 7) = "Qty"
        .TextMatrix(0, 8) = "SATUAN"        'SATUAN
        .TextMatrix(0, 9) = "Kg"
        .TextMatrix(0, 10) = "ISI"
        .TextMatrix(0, 11) = "Total Kg"
        .Col = 12
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 12) = "No RcnBB"
        .Col = 13
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 13) = "No RcnPack"
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .ColAlignmentFixed(12) = flexAlignCenterCenter
        .ColAlignmentFixed(13) = flexAlignCenterCenter
    End With
End Sub

Private Sub btnSave_Click()
    If lblkdrcn = "" Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
    If MsgBox("Apakah tanggal rencana produksi sudah benar", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If MsgBox("Apakah Data Qty,No RcnBB dan No RcnPack Plan sudah diisi dengan benar", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    'SIMPAN
    OBJ.Open dsn
    SQL = "Select * From am_rcnprod WHERE 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        RS.AddNew
        RS!Kd_RCN = lblkdrcn
        RS!KODE_PRODUK = grid.TextMatrix(grid.Row, 1)
        RS!PRODUK = grid.TextMatrix(grid.Row, 2)
        RS!KODE_BRG = grid.TextMatrix(grid.Row, 4)
        RS!NAMA_BARANG = grid.TextMatrix(grid.Row, 5)
        RS!Qty = grid.TextMatrix(grid.Row, 7)
        RS!KODE_SATUAN = grid.TextMatrix(grid.Row, 6)
        RS!SATUAN = grid.TextMatrix(grid.Row, 8)
        RS!KG = grid.TextMatrix(grid.Row, 9)
        RS!ISI = grid.TextMatrix(grid.Row, 10)
        RS!Totalkg = grid.TextMatrix(grid.Row, 11)
        RS!TGL1 = Format(dtpfrom, "yyyy/MM/dd")
        RS!TGL2 = Format(dtpto, "yyyy/MM/dd")
        RS!KD_PACK = grid.TextMatrix(grid.Row, 13)
        RS!No_RCNBB = grid.TextMatrix(grid.Row, 12)
        RS!FLAG = "0"
        RS.Update
        If grid.Rows = grid.Row + 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    MsgBox "Production Plan is saved", vbInformation, AppName
    
    hapusgrid
End Sub

Private Sub dtpfrom_Change()
    bln = Month(dtpfrom)
    thn = Year(dtpfrom)
End Sub

Private Sub Form_Load()
    setGrid
    initGrid
    dtpfrom = Date
    dtpto = Date
    bln = Month(dtpfrom)
    thn = Year(dtpfrom)
    lblkdrcn = getkdrcn
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
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
        grid.TextMatrix(grid.Row, 11) = ""
        grid.TextMatrix(grid.Row, 12) = ""
        grid.TextMatrix(grid.Row, 13) = ""
        
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
    initGrid
    dtpfrom = Date
    dtpto = Date
    lbltotalkg = "0"
    bln = Month(dtpfrom)
    thn = Year(dtpfrom)
    lblkdrcn = getkdrcn
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
    grid.TextMatrix(grid.Row, 11) = ""
    grid.TextMatrix(grid.Row, 12) = ""
    grid.TextMatrix(grid.Row, 13) = ""
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
            grid.TextMatrix(grid.Row, 11) = ""
            grid.TextMatrix(grid.Row, 12) = ""
            grid.TextMatrix(grid.Row, 13) = ""
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
        grid.TextMatrix(grid.Row, 11) = grid.TextMatrix(grid.Row + 1, 11)
        grid.TextMatrix(grid.Row, 12) = grid.TextMatrix(grid.Row + 1, 12)
        grid.TextMatrix(grid.Row, 13) = grid.TextMatrix(grid.Row + 1, 13)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If grid.TextMatrix(grid.Row, 13) = "" Then
                    If MsgBox("Hapus baris ini ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                        Set grid.CellPicture = uncheck
                        hapusrow
                        Call Totalkg
                        Exit Sub
                    End If
                Else
                    If MsgBox("Menghapus baris ini akan menghapus data pada Nomor Packaging :" & grid.TextMatrix(grid.Row, 13) & vbLf & "Hapus data ini ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                        Set grid.CellPicture = uncheck
                        'hapus data packaging
                        Call hapuspack
                        hapusrow
                        Call Totalkg
                        Exit Sub
                    End If
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1:
            'CARI PRODUK
            If lblkdrcn = "" Then lblkdrcn = getkdrcn
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            carisql1 = "select * from am_itemcode where (lev =3 or lev =4)"
            namatabel = "Produk"
            frmsearch.Show vbModal
        Case 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            'CARI BARANG JADI BY PRODUK
            carisql1 = "Select distinct a.kode_barang_jadi,b.NamaBarang,a.kode_satuan,c.NamaSatuan"
            carisql1 = carisql1 + " From list_produk_hasil a inner join am_itemdtl b on a.kode_barang_jadi = b.KodeBarang"
            carisql1 = carisql1 + " inner join am_unit c on a.kode_satuan = c.KodeSatuan"
            carisql1 = carisql1 + " where a.kode_produk='" & grid.TextMatrix(grid.Row, 1) & "'"
            namatabel = "Barang Jadi."
            frmsearch.Show vbModal
        Case 7:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 4) = "" Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 20
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 12: 'NOMOR KODE BAHAN BAKU
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 4) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 7) = "" Or grid.TextMatrix(grid.Row, 7) = "0" Then
                MsgBox "Kolom Qty tidak boleh bernilai 0", vbExclamation, AppName
                Exit Sub
            End If
            If grid.TextMatrix(grid.Row, 12) <> "" Then
                'JIKA NOMOR BB SUDAH ADA
                hasil4 = grid.TextMatrix(grid.Row, 12)
                frmrcnbb.lblnobb = grid.TextMatrix(grid.Row, 12)
            End If
            frmrcnbb.lblkdproduk = grid.TextMatrix(grid.Row, 1)
            frmrcnbb.lblnamaproduk = grid.TextMatrix(grid.Row, 2)
            frmrcnbb.lblkode = grid.TextMatrix(grid.Row, 4)
            frmrcnbb.lblnamabrg = grid.TextMatrix(grid.Row, 5)
            frmrcnbb.lblqty = grid.TextMatrix(grid.Row, 7)
            frmrcnbb.lblsatuan = grid.TextMatrix(grid.Row, 8)
            frmrcnbb.lblnorcn = lblkdrcn
            frmrcnbb.Show vbModal
        Case 13: 'NOMOR KODE PACKAGING
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 4) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 7) = "" Or grid.TextMatrix(grid.Row, 7) = "0" Then
                MsgBox "Kolom Qty tidak boleh bernilai 0", vbExclamation, AppName
                Exit Sub
            End If
            If grid.TextMatrix(grid.Row, 13) <> "" Then
                'JIKA NOMOR PACKAGING SUDAH ADA
                hasil4 = grid.TextMatrix(grid.Row, 13)
                frmrcnpack.lblkdpack = grid.TextMatrix(grid.Row, 13)
            End If
            frmrcnpack.lblproduk = grid.TextMatrix(grid.Row, 1)
            frmrcnpack.lblnamaproduk = grid.TextMatrix(grid.Row, 2)
            frmrcnpack.lblkodebrg = grid.TextMatrix(grid.Row, 4)
            frmrcnpack.lblnamabarang = grid.TextMatrix(grid.Row, 5)
            frmrcnpack.lblkdsatuan = grid.TextMatrix(grid.Row, 6)
            frmrcnpack.lblqty = grid.TextMatrix(grid.Row, 7)
            frmrcnpack.lblsatuan = grid.TextMatrix(grid.Row, 8)
            frmrcnpack.lblkg = grid.TextMatrix(grid.Row, 9)
            frmrcnpack.lblisi = grid.TextMatrix(grid.Row, 10)
            frmrcnpack.lbltotal = grid.TextMatrix(grid.Row, 11)
            frmrcnpack.lblkdrcn = lblkdrcn
            frmrcnpack.Show vbModal
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid.Col
        Case 1:
            grid.TextMatrix(grid.Row, 1) = hasil1
            grid.TextMatrix(grid.Row, 2) = hasil2
            carisql1 = ""
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            
            grid.Col = 0
            Set grid.CellPicture = uncheck
            SetAlternatingGrid grid.Row
            grid.Rows = grid.Row + 2
            
        Case 4:
            grid.TextMatrix(grid.Row, 4) = hasil
            grid.TextMatrix(grid.Row, 5) = hasil1
            grid.TextMatrix(grid.Row, 6) = hasil2
            grid.TextMatrix(grid.Row, 8) = hasil3
            'OPEN KG BASE UNIT
            OBJ.Open dsn
            SQL = "Select * From am_itemkg Where kodebarang='" & hasil & "' and kodesatuan='" & hasil2 & "' and tahun='" & thn & "'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                MsgBox "Kg Base Unit tidak ditemukan", vbCritical, AppName
            Else
                If bln = "1" Then grid.TextMatrix(grid.Row, 9) = RS!kg1
                If bln = "2" Then grid.TextMatrix(grid.Row, 9) = RS!kg2
                If bln = "3" Then grid.TextMatrix(grid.Row, 9) = RS!kg3
                If bln = "4" Then grid.TextMatrix(grid.Row, 9) = RS!kg4
                If bln = "5" Then grid.TextMatrix(grid.Row, 9) = RS!kg5
                If bln = "6" Then grid.TextMatrix(grid.Row, 9) = RS!kg6
                If bln = "7" Then grid.TextMatrix(grid.Row, 9) = RS!kg7
                If bln = "8" Then grid.TextMatrix(grid.Row, 9) = RS!kg8
                If bln = "9" Then grid.TextMatrix(grid.Row, 9) = RS!kg9
                If bln = "10" Then grid.TextMatrix(grid.Row, 9) = RS!kg10
                If bln = "11" Then grid.TextMatrix(grid.Row, 9) = RS!kg11
                If bln = "12" Then grid.TextMatrix(grid.Row, 9) = RS!kg12
            End If
            
            SQL = "Select * From am_itemdtl where KodeBarang='" & hasil & "' and KodeSatuan='" & hasil2 & "'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                MsgBox "Konversi Item tidak ditemukan", vbCritical, AppName
            Else
                grid.TextMatrix(grid.Row, 10) = RS!konversi
            End If
            OBJ.Close
            
            carisql1 = ""
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
        Case 12:
            If hasil = "hapus" And grid.TextMatrix(grid.Row, 12) <> "" Then
                grid.TextMatrix(grid.Row, 12) = ""
                hasil4 = ""
            Else
                grid.TextMatrix(grid.Row, 12) = hasil
            End If
            hasil = ""
        Case 13:
            If hasil = "hapus" And grid.TextMatrix(grid.Row, 13) <> "" Then
                grid.TextMatrix(grid.Row, 13) = ""
                hasil4 = ""
            Else
                grid.TextMatrix(grid.Row, 13) = hasil
            End If
            hasil = ""
    End Select
End Sub

Sub hapuspack()
    OBJ.Open dsn
    SQL = "DELETE FROM am_rcnpack WHERE KD_PACK='" & grid.TextMatrix(grid.Row, 13) & "'"
    Set RS = OBJ.Execute(SQL)
    OBJ.Close
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
    grid.Col = 1
    grid.CellBackColor = &H80FFFF
    grid.Col = 4
    grid.CellBackColor = &H80FFFF
    grid.Col = 7
    grid.CellBackColor = &H80FFFF
    grid.Col = 12
    grid.CellBackColor = &H80FFFF
    grid.Col = 13
    grid.CellBackColor = &H80FFFF
End Function

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim konvToKg As Double
    If KeyAscii = 13 Then
        If IsNull(txtnilai.Value) Or txtnilai = 0 Then
            grid.TextMatrix(grid.Row, 7) = "0"
        Else
            grid.TextMatrix(grid.Row, 7) = txtnilai.text
        End If
        grid.TextMatrix(grid.Row, 11) = grid.TextMatrix(grid.Row, 7) * CDbl(grid.TextMatrix(grid.Row, 9) * grid.TextMatrix(grid.Row, 10))
        grid.TextMatrix(grid.Row, 11) = Format(grid.TextMatrix(grid.Row, 11), "##,###,##0.00")
        Call Totalkg
        
        grid.SetFocus
    End If
End Sub

Private Sub Totalkg()
On Error Resume Next
    grid.Row = 1
    tkg = 0
    Do While True
        DoEvents
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        tkg = CDbl(Format(grid.TextMatrix(grid.Row, 11), "general number") + CDbl(tkg))
        grid.Row = grid.Row + 1
    Loop
    tkg = Format(tkg, "##,###,##0.00")
    lbltotalkg = tkg
End Sub

Private Function getkdrcn() As String
On Error GoTo Err_handler:
    Dim strformat As String
    strformat = Format(Date, "yymmdd")
    
    Dim str99 As String
    Dim no As String
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select top 1 kd_rcn from am_rcnprod "
    SQL = SQL + "where kd_rcn like 'RCN' + '" + strformat + "%' order by kd_rcn desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!Kd_RCN, 3)
    Else
        str99 = 0
    End If
        str99 = str99 + 1
        
    If Len(str99) = 1 Then no = "RCN" & strformat & "00" & str99
    If Len(str99) = 2 Then no = "RCN" & strformat & "0" & str99
    If Len(str99) = 3 Then no = "RCN" & strformat & str99
        
    getkdrcn = no
    OBJ.Close
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function
