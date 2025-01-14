VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmkonversi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konversi Kemasan"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3420
      Left            =   105
      TabIndex        =   9
      Top             =   3465
      Width           =   11190
      _Version        =   851970
      _ExtentX        =   19738
      _ExtentY        =   6032
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
      Color           =   8
      ItemCount       =   1
      Item(0).Caption =   "Item konversi"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "grid2"
      Item(0).Control(1)=   "txtnilai"
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calculator      =   "frmkonversi.frx":0000
         Caption         =   "frmkonversi.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmkonversi.frx":008C
         Keys            =   "frmkonversi.frx":00AA
         Spin            =   "frmkonversi.frx":00EC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   3015
         Left            =   45
         TabIndex        =   10
         Top             =   345
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   5318
         _Version        =   393216
         BackColor       =   -2147483628
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   -2147483642
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
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
      Left            =   7095
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   270
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
      Left            =   6735
      Picture         =   "frmkonversi.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   270
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
      Left            =   6480
      Picture         =   "frmkonversi.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtkodeproduk 
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
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   1050
   End
   Begin VB.TextBox txtproduk 
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
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   210
      Width           =   3060
   End
   Begin XtremeSuiteControls.PushButton cmdproduksi 
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   240
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "PRODUK :"
      BackColor       =   -2147483644
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
      TextAlignment   =   1
      Appearance      =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2835
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Pilih item untuk konversi kemasan"
      Top             =   585
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   -2147483628
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
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
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   420
      Left            =   9900
      TabIndex        =   7
      Top             =   6930
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "CLOSE"
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
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   420
      Left            =   8535
      TabIndex        =   8
      Top             =   6930
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "SAVE"
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
   Begin XtremeSuiteControls.PushButton btnclear 
      Height          =   420
      Left            =   7155
      TabIndex        =   12
      Top             =   6930
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "CLEAR"
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
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   420
      Left            =   120
      TabIndex        =   13
      Top             =   6945
      Width           =   1845
      _Version        =   851970
      _ExtentX        =   3254
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "View List Konversi"
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
      Left            =   2070
      Top             =   6960
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
   Begin VB.Image Image1 
      Height          =   1665
      Left            =   7605
      Picture         =   "frmkonversi.frx":0744
      Stretch         =   -1  'True
      Top             =   1155
      Width           =   2385
   End
End
Attribute VB_Name = "frmkonversi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private RST1 As ADODB.Recordset
Private SQL As String
Private poscol As Integer
Private posrow As Integer
Dim i As Integer

Private Sub btnClear_Click()
    txtkodeproduk = ""
    txtproduk = ""
    hapusgrid
    hapusgrid2
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If txtkodeproduk = "" Then Exit Sub
    If grid2.TextMatrix(grid2.Row, 1) = "" Then
        MsgBox "Item konversi belum terisi...", vbCritical
        Exit Sub
    End If
    
    OBJ.Open dsn
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            
        SQL = "Select * From list_konversi Where kodebarang = '" & grid2.TextMatrix(grid2.Row, 1) & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        If Not RST.EOF Then
            If MsgBox("Item sudah terkonversi." + vbCrLf + "Klik YES untuk update ?" + vbCrLf + _
            grid2.TextMatrix(grid2.Row, 1) + grid2.TextMatrix(grid2.Row, 2) + _
            " 1 " + grid2.TextMatrix(grid2.Row, 5) + " = " + _
            grid2.TextMatrix(grid2.Row, 6) + " " + grid2.TextMatrix(grid2.Row, 8), vbYesNo + vbQuestion, "Update Mode") = vbNo Then GoTo lanjut:
            'UPDATE KONVERSI KEMASAN
            RST!konversi = grid2.TextMatrix(grid2.Row, 6)
            RST!kodekemasan = grid2.TextMatrix(grid2.Row, 7)
            RST!namakemasan = grid2.TextMatrix(grid2.Row, 8)
            RST.Update
            'UPDATE AM_ITEMDTL
            updatekonv
lanjut:
        ElseIf RST.EOF Then
            RST.AddNew
            RST!KodeBarang = grid2.TextMatrix(grid2.Row, 1)
            RST!namabarang = grid2.TextMatrix(grid2.Row, 2)
            RST!kodesatuan = grid2.TextMatrix(grid2.Row, 4)
            RST!konversi = grid2.TextMatrix(grid2.Row, 6)
            RST!kodekemasan = grid2.TextMatrix(grid2.Row, 7)
            RST!namakemasan = grid2.TextMatrix(grid2.Row, 8)
            RST.Update
        End If
        grid2.Row = grid2.Row + 1
    Loop
    
    OBJ.Close
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    btnClear_Click
End Sub
Private Sub updatekonv()
    SQL = "Select * From am_itemdtl Where kodebarang = '" & grid2.TextMatrix(grid2.Row, 1) & "' "
    SQL = SQL + "and kodesatuan ='" & grid2.TextMatrix(grid2.Row, 4) & "'"
    
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
    If Not RST.EOF Then
        RST!konversi = Format(grid2.TextMatrix(grid2.Row, 6), "general number")
        RST.Update
    End If
End Sub

Private Sub btnview_Click()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.ReportFileName = AppPath & "\reports\produksi\list_konversi.rpt"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdproduksi_Click()
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Anda tidak memiliki akses..! ", vbCritical, AppName
        Exit Sub
    End If
proses:
    namatabel = "produk"
    carisql1 = "select kode_produk,nama_produk from list_produk_master"
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduksi_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeproduk = hasil
    txtproduk = hasil1
    hasil = ""
    hasil1 = ""
    carisql1 = ""
    findbrgjadi
End Sub

Private Sub findbrgjadi()
    SQL = "select a.kodebarang,a.namabarang,b.kode_satuan,c.namasatuan "
    SQL = SQL + "from am_itemmst a inner join list_produk_hasil b "
    SQL = SQL + "on a.kodebarang=b.kode_barang_jadi inner join am_unit c "
    SQL = SQL + "on b.kode_satuan=c.kodesatuan "
    SQL = SQL + "and b.kode_produk='" & txtkodeproduk & "' "
    
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    hapusgrid
    hapusgrid2
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!KodeBarang
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = RST!kode_satuan
        grid.TextMatrix(grid.Row, 4) = RST!namasatuan
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
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

        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    SetGrid
End Sub
Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.TextMatrix(grid2.Row, 8) = ""
        
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    SetGrid2
End Sub
Private Sub SetGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1000
    End With
End Sub
Private Sub SetGrid2()
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 400
        .ColWidth(4) = 0
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 3000
    End With
End Sub

Private Sub initGrid()
    poscol = grid2.Col
    posrow = grid2.Row
    With grid
        .Cols = 5
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "KODE"
        .TextMatrix(0, 4) = "SATUAN"
    End With
    With grid2
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "KODE"
        .TextMatrix(0, 5) = "SATUAN"
        .TextMatrix(0, 6) = "KONVERSI"
        .TextMatrix(0, 7) = "KODE"
        .TextMatrix(0, 8) = "KEMASAN"

        For i = 6 To 8
            grid2.Col = i
            grid2.CellBackColor = &HE0E0E0
        Next
    End With
End Sub

Private Sub Form_Load()
    initGrid
    SetGrid
    SetGrid2
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub

    grid2.Col = 1
    grid2.Row = 1
    'PERIKSA GRID2
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        With grid2
            If .TextMatrix(.Row, 1) = "" Then GoTo isigrid2:
            If .TextMatrix(.Row, 1) = grid.TextMatrix(grid.Row, 1) Then Exit Sub
                .Row = .Row + 1
        End With
    Loop
    
    
isigrid2:
    grid2.Col = 0
    Set grid2.CellPicture = uncheck
    grid2.TextMatrix(grid2.Row, 1) = grid.TextMatrix(grid.Row, 1)
    grid2.TextMatrix(grid2.Row, 2) = grid.TextMatrix(grid.Row, 2)
    grid2.TextMatrix(grid2.Row, 3) = "1"
    grid2.TextMatrix(grid2.Row, 4) = grid.TextMatrix(grid.Row, 3)
    grid2.TextMatrix(grid2.Row, 5) = grid.TextMatrix(grid.Row, 4)
    caribrgjadi
    grid2.Rows = grid2.Rows + 1
    grid2.Row = grid2.Row + 1

    For i = 6 To 8
        grid2.Col = i
        grid2.CellBackColor = &HE0E0E0
    Next
End Sub

Private Sub caribrgjadi()
    OBJ.Open dsn
    SQL = "select * from am_itemdtl where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan ='" & grid.TextMatrix(grid.Row, 3) & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid2.TextMatrix(grid2.Row, 6) = RST!konversi
    
    SQL = "Select * From list_konversi Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        grid2.TextMatrix(grid2.Row, 7) = RST!kodekemasan
        grid2.TextMatrix(grid2.Row, 8) = RST!namakemasan
    Else
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.TextMatrix(grid2.Row, 8) = ""
    End If
    
    OBJ.Close
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If grid2.TextMatrix(1, 1) = "" Then Exit Sub
    poscol = grid2.Col
    posrow = grid2.Row
    
    Select Case grid2.Col
        Case 0:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete this row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow2
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
                End If
        Case 6:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
            txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtnilai.Left = grid2.Left + grid2.CellLeft
            txtnilai.Top = grid2.Top + grid2.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 7:
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid2_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid2.Col
        Case 7:
            grid2.TextMatrix(grid2.Row, 7) = hasil
            grid2.TextMatrix(grid2.Row, 8) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, 6) = txtnilai.text
        grid2.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub hapusrow2()
    grid2.TextMatrix(grid2.Row, 1) = ""
    grid2.TextMatrix(grid2.Row, 2) = ""
    grid2.TextMatrix(grid2.Row, 3) = ""
    grid2.TextMatrix(grid2.Row, 4) = ""
    grid2.TextMatrix(grid2.Row, 5) = ""
    grid2.TextMatrix(grid2.Row, 6) = ""
    grid2.TextMatrix(grid2.Row, 7) = ""
    grid2.TextMatrix(grid2.Row, 8) = ""
    
    Do While True
        If grid2.TextMatrix(grid2.Row + 1, 1) = "" Then
            grid2.TextMatrix(grid2.Row, 1) = ""
            grid2.TextMatrix(grid2.Row, 2) = ""
            grid2.TextMatrix(grid2.Row, 3) = ""
            grid2.TextMatrix(grid2.Row, 4) = ""
            grid2.TextMatrix(grid2.Row, 5) = ""
            grid2.TextMatrix(grid2.Row, 6) = ""
            grid2.TextMatrix(grid2.Row, 7) = ""
            grid2.TextMatrix(grid2.Row, 8) = ""
            Exit Do
        End If
        grid2.TextMatrix(grid2.Row, 1) = grid2.TextMatrix(grid2.Row + 1, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row + 1, 2)
        grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row + 1, 3)
        grid2.TextMatrix(grid2.Row, 4) = grid2.TextMatrix(grid2.Row + 1, 4)
        grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row + 1, 5)
        grid2.TextMatrix(grid2.Row, 6) = grid2.TextMatrix(grid2.Row + 1, 6)
        grid2.TextMatrix(grid2.Row, 7) = grid2.TextMatrix(grid2.Row + 1, 7)
        grid2.TextMatrix(grid2.Row, 8) = grid2.TextMatrix(grid2.Row + 1, 8)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = blank
End Sub
