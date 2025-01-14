VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmaddlot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Lot Bahan"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9465
   StartUpPosition =   1  'CenterOwner
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmaddlot.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaddlot.frx":006C
      Key             =   "frmaddlot.frx":008A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Input lot manual"
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
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtnolot 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   3840
      Width           =   4215
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
      Left            =   5895
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   120
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
      Left            =   5655
      Picture         =   "frmaddlot.frx":00C6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   120
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
      Left            =   5400
      Picture         =   "frmaddlot.frx":03A8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmaddlot.frx":06F6
      Caption         =   "frmaddlot.frx":0716
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaddlot.frx":0782
      Keys            =   "frmaddlot.frx":07A0
      Spin            =   "frmaddlot.frx":07E2
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
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   435
      Left            =   7200
      TabIndex        =   0
      Top             =   3840
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton btnclose 
      Height          =   435
      Left            =   8280
      TabIndex        =   1
      Top             =   3840
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
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
      Appearance      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2820
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4974
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   -2147483632
      BackColorBkg    =   8421504
      AllowUserResizing=   1
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
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   435
      Left            =   6120
      TabIndex        =   13
      Top             =   3840
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
   End
   Begin VB.Label lbltotalHpp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7680
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblNoSJ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7680
      TabIndex        =   12
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblqtysop 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Qty Required"
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
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Total Available"
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
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblnmbahan 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4020
   End
   Begin VB.Label lblbahan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmaddlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset
Private SQL1 As String

Dim kodelot As String
Dim hpperkg As Double
Dim sisaqty As Double
Dim qtyrow As Double
Dim posrow As String
Dim lotdel As String
Dim kdsat As String

Private Sub btnclose_Click()
    Unload Me
End Sub
Private Sub initGrid()
    With grid
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Kg"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "Qty"
        .TextMatrix(0, 5) = "Hpp" 'Per Kg
        .TextMatrix(0, 6) = "Qty Use"
        .TextMatrix(0, 7) = "Satuan"
        .TextMatrix(0, 8) = "Sisa"
        .ColAlignmentFixed(4) = flexAlignRightCenter
        .Col = 6
        .CellBackColor = &H80FFFF
    End With
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 0 '1000
        .ColWidth(2) = 0 '2000
        .ColWidth(3) = 1800
        .ColWidth(4) = 1000
        .ColWidth(5) = 0
        .ColWidth(6) = 1000
        .ColWidth(7) = 1500
        .ColWidth(8) = 1000
    End With
End Sub

Private Sub btnsave_Click()
    If grid.TextMatrix(1, 1) = "" Then Exit Sub
    If lbltotal > lblqtysop Then
        MsgBox "Qty lot melebihi jumlah Qty SOP", vbExclamation, AppName
        Exit Sub
    ElseIf lbltotal < lblqtysop Then
        MsgBox "Qty lot tidak mencukupi", vbExclamation, AppName
        Exit Sub
    End If
    'Simpan ke am_stoklot
    OBJ.Open dsn
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "Insert into am_sjlot ("
        SQL = SQL + "nosj,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "satuan,"
        SQL = SQL + "kg,"
        SQL = SQL + "nolot,"
        SQL = SQL + "hpp,"
        SQL = SQL + "flag,"
        SQL = SQL + "kode,"
        SQL = SQL + "keterangan)"
        SQL = SQL + " Values ('" & lblNoSJ & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 3) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'0','" & kodelot & "',"
        SQL = SQL + "'" & txtnolot & "')"
        Set RST = OBJ.Execute(SQL)
        
         'menandai lot yang sudah kosong (Habis)
        If grid.TextMatrix(grid.Row, 8) = "0.00" Then
            SQL = "Update am_sjlot set flag='1' Where nolot='" & grid.TextMatrix(grid.Row, 3) & "'"
            SQL = SQL + " And kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    hasil = txtnolot
    hasil1 = kodelot
    lotbarang = ""
    lotbarang1 = ""
    lotbarang2 = ""
    lotbarang3 = ""
    lotbarang4 = ""
    cmdclear_Click
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, AppName
    Unload Me
End Sub

Private Sub cmdclear_Click()
    sisaqty = 0
    lbltotal = 0
    txtnilai = 0
    lbltotalHpp = 0
    txtnolot = ""
    hapusgrid
End Sub

Private Sub Form_Load()
    initGrid
    setGrid
    lblbahan = lotbarang
    lblnmbahan = lotbarang1
    lblqtysop = Format(lotbarang2, "#,##0.00")
    lblNoSJ = lotbarang3
    OBJ.Open dsn
    SQL = "Select NamaSatuan From am_unit Where KodeSatuan='" & lotbarang4 & "'"
    Set RST = OBJ.Execute(SQL)
    kdsat = RST!namasatuan
    OBJ.Close
    'kdsat = lotbarang4
    kodelot = getkode
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    
                    totalQgrid
                    'Hapus nolot barang
                    OBJ.Open dsn
                    SQL = "Delete From am_sjlot Where nosj='" & lblNoSJ & "' and kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
                    Set RST = OBJ.Execute(SQL)
                    OBJ.Close
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
                End If
        Case 3: 'nolot & Qty
            If lblqtysop = lbltotal Then
                MsgBox "Qty telah mencukupi", vbExclamation, AppName
                Exit Sub
            End If
            'Input lot manual
            If Check1.Value = Checked Then
                txtket.Width = grid.ColWidth(grid.Col) - 40
                txtket = grid.TextMatrix(grid.Row, grid.Col)
                txtket.Left = grid.Left + grid.CellLeft
                txtket.Top = grid.Top + grid.CellTop + 20
                txtket.Visible = True
                txtket.SetFocus
            Exit Sub
            End If
            grid.TextMatrix(grid.Row, 3) = ""
            
            posrow = grid.Row
            namatabel = "Lot Item"
            carisql1 = "Select a.nolot,a.kodebarang,b.namabarang,(SUM(a.kgperpalet/a.kg)-isnull(c.total,0))'qty',a.hppperkg,a.kg"
            carisql1 = carisql1 + " From list_hpp_produksi a inner join am_itemmst b on a.kodebarang = b.KodeBarang"
            carisql1 = carisql1 + " left join (Select nolot,kodebarang,SUM(qty)'total' from am_sjlot Where kodebarang = '" & lotbarang & "' group by nolot,kodebarang) c"
            carisql1 = carisql1 + " on a.kodebarang = c.kodebarang and a.nolot = c.nolot"
            carisql1 = carisql1 + " Where a.kodebarang = '" & lotbarang & "' and a.flag = '2'" ' group by a.nolot,a.kodebarang,a.hppperkg,a.kg,c.total
            
            frmsearch.Show vbModal
        Case 6: 'input qty use
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    Select Case grid.Col
        Case 3:
            If hasil = "" Then Exit Sub
            If hasil1 = ".00" Then Exit Sub
            'Cek Lot yang sudah terpakai
            With grid
                .Row = 1
                Do While True
                    If .TextMatrix(.Row, 3) = "" Then Exit Do
                    If .TextMatrix(.Row, 3) = hasil Then
                        MsgBox "Lot is already exist...!", vbCritical, AppName
                        grid.Row = posrow
                        hasil = ""
                        Exit Sub
                    End If
                    .Row = .Row + 1
                Loop
            
                
            grid.Row = posrow
            grid.TextMatrix(grid.Row, 6) = "0.00"
            totalQgrid
            grid.Row = posrow
            grid.Col = 0
            Set grid.CellPicture = uncheck
            grid.TextMatrix(grid.Row, 1) = lotbarang
            grid.TextMatrix(grid.Row, 2) = hasil3
            grid.TextMatrix(grid.Row, 3) = hasil
            grid.TextMatrix(grid.Row, 4) = Format(hasil1, "#,##0.00")
            grid.TextMatrix(grid.Row, 5) = Format(hasil2, "#,##0.00")
            qtyrow = CInt(lbltotal) + hasil3
            If hasil1 > sisaqty Then 'Jika jumlah qty lot lebih besar dari jumlah yang dibutuhkan
                If hasil1 < lotbarang2 Then
                    If lbltotal < lotbarang2 And qtyrow > lotbarang2 Then
                        grid.TextMatrix(grid.Row, 6) = Format(hasil1 - (CInt(qtyrow) - lotbarang2), "##0.00")
                    Else
                        grid.TextMatrix(grid.Row, 6) = Format(hasil1, "#,##0.00")
                    End If
                Else
                    grid.TextMatrix(grid.Row, 6) = Format(lotbarang2, "#,##0.00")
                End If
            Else
                grid.TextMatrix(grid.Row, 6) = Format(hasil1, "#,##0.00")
            End If

            grid.TextMatrix(grid.Row, 7) = kdsat
            grid.TextMatrix(grid.Row, 8) = Format(grid.TextMatrix(grid.Row, 4) - grid.TextMatrix(grid.Row, 6), "##0.00")
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            carisql1 = ""
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            If grid.Rows <> grid.Row + 1 Then grid.Rows = grid.Row + 1
            totalQgrid
            End With
    End Select
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
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""

        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Sub lbltotal_Change()
    If lbltotal > lblqtysop Then
        lbltotal.BackColor = vbRed
    ElseIf lbltotal = lblqtysop Then
        lbltotal.BackColor = vbGreen
    Else
        lbltotal.BackColor = &H80000005
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 3
                grid.SetFocus
                grid.Col = 3
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = lotbarang
                grid.TextMatrix(grid.Row, 2) = "0"
                grid.TextMatrix(grid.Row, 4) = "0"
                grid.TextMatrix(grid.Row, 5) = "0"
                grid.TextMatrix(grid.Row, 6) = Format(lotbarang2, "#,##0.00")
                grid.TextMatrix(grid.Row, 8) = "0"
                'OBJ.Open dsn
                'SQL = "Select NamaSatuan From am_unit Where KodeSatuan='" & lotbarang4 & "'"
                'Set RST = OBJ.Execute(SQL)
                grid.TextMatrix(grid.Row, 7) = kdsat
                'OBJ.Close
                grid.TextMatrix(grid.Row, 3) = txtket
                txtket = ""
                txtket.Visible = False
                grid.Col = 0
                Set grid.CellPicture = uncheck
                
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                If grid.Rows <> grid.Row + 1 Then grid.Rows = grid.Row + 1
                totalQgrid
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If grid.TextMatrix(grid.Row, 2) <> "0" Then
            If Format(txtnilai.Value, "#,##0.00") > CDbl(Format(grid.TextMatrix(grid.Row, 4), "#,##0.00")) Then
                MsgBox "Qty you entered (" & txtnilai & ") is greater than what is available (" & Format(grid.TextMatrix(grid.Row, 4), "##,##0") & ")", vbCritical, AppName
                txtnilai = ""
                Exit Sub
            End If
        End If
        grid.TextMatrix(grid.Row, 6) = Format(txtnilai, "#,##0.00")
        If grid.TextMatrix(grid.Row, 5) = "0" Then
            hpperkg = "0"
        Else
            hpperkg = Format(grid.TextMatrix(grid.Row, 5), "general number") / Format(grid.TextMatrix(grid.Row, 4), "general number")
        End If
        
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row, 4) - txtnilai
        grid.TextMatrix(grid.Row, 8) = Format(grid.TextMatrix(grid.Row, 8), "#,##0.00")
        txtnilai.Visible = False
        totalQgrid
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub totalQgrid()
On Error Resume Next
'TOTAL GRID
Dim lot As String
    grid.Row = 1
    tg = 0
    th = 0
    lot = ""
    Do While True
        DoEvents
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            tg = CDbl(Format(grid.TextMatrix(grid.Row, 6), "general number") + CDbl(tg))
            th = CDbl(Format(grid.TextMatrix(grid.Row, 7), "general number") + CDbl(th))
            lot = lot & grid.TextMatrix(grid.Row, 3) & ","
                grid.Row = grid.Row + 1
    Loop
        tg = Format(tg, "##,###,##0.00")
        lbltotal = tg
        tg = lotbarang2 - tg
        sisaqty = Format(tg, "##,###,##0.00")
        txtnolot = lot
        th = Format(th, "##,###,##0.00")
        lbltotalHpp = th
End Sub

Function getkode() As String    '230323001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kd As Long
    
    strnumber = Format(Date, "yymmdd")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(kode)as kr from am_sjlot where kode like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getkode = strnumber + "001"
    Else
        kd = CLng(Mid(RST!kr, 7, 3)) + 1
        
        If (Len(Trim(Str(kd))) = 1) Then
            tempkode = strnumber + "00" + Trim(Str(kd))
        End If
        If (Len(Trim(Str(kd))) = 2) Then
            tempkode = strnumber + "0" + Trim(Str(kd))
        End If
        If (Len(Trim(Str(kd))) = 3) Then
            tempkode = strnumber + Trim(Str(kd))
        End If
        getkode = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getkode = strnumber + "001"
End Function

Function tanggalpakaibase()
    tanggalpakaibase = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
