VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmaddsop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add SOP"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbpass 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmAddSOP.frx":0000
      Left            =   11685
      List            =   "frmAddSOP.frx":000A
      TabIndex        =   24
      Text            =   "PASS"
      Top             =   975
      Visible         =   0   'False
      Width           =   1590
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
      Left            =   7305
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   240
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
      Left            =   7785
      Picture         =   "frmAddSOP.frx":001B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   240
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
      Left            =   7545
      Picture         =   "frmAddSOP.frx":02FD
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai2 
      Height          =   255
      Left            =   9330
      TabIndex        =   20
      Top             =   885
      Visible         =   0   'False
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   450
      Calculator      =   "frmAddSOP.frx":064B
      Caption         =   "frmAddSOP.frx":066B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddSOP.frx":06D7
      Keys            =   "frmAddSOP.frx":06F5
      Spin            =   "frmAddSOP.frx":0737
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   9345
      TabIndex        =   16
      Top             =   450
      Visible         =   0   'False
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   450
      Calculator      =   "frmAddSOP.frx":075F
      Caption         =   "frmAddSOP.frx":077F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddSOP.frx":07EB
      Keys            =   "frmAddSOP.frx":0809
      Spin            =   "frmAddSOP.frx":084B
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   420
      Left            =   10155
      TabIndex        =   15
      Top             =   7920
      Width           =   1050
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      Height          =   420
      Left            =   11220
      TabIndex        =   14
      Top             =   7920
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   1890
      Left            =   75
      TabIndex        =   10
      Top             =   5955
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      WordWrap        =   -1  'True
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtnolot 
      Height          =   315
      Left            =   1170
      TabIndex        =   9
      Top             =   960
      Width           =   5070
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   3330
      Left            =   90
      TabIndex        =   6
      Top             =   1860
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   5874
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Top             =   600
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   88866817
      CurrentDate     =   41357
   End
   Begin VB.TextBox txtnmproduk 
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Top             =   210
      Width           =   3810
   End
   Begin VB.TextBox txtkdproduk 
      Height          =   330
      Left            =   1170
      TabIndex        =   2
      Top             =   210
      Width           =   1245
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   420
      Left            =   12285
      TabIndex        =   0
      Top             =   7920
      Width           =   1050
   End
   Begin Chameleon.chameleonButton cmdproduk 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   255
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Produk"
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
      MICON           =   "frmAddSOP.frx":0873
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtjumlah 
      Height          =   285
      Left            =   10080
      TabIndex        =   13
      Top             =   5250
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   503
      Calculator      =   "frmAddSOP.frx":0B8D
      Caption         =   "frmAddSOP.frx":0BAD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddSOP.frx":0C19
      Keys            =   "frmAddSOP.frx":0C37
      Spin            =   "frmAddSOP.frx":0C81
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin VB.Line Line6 
      X1              =   8250
      X2              =   8250
      Y1              =   5685
      Y2              =   5925
   End
   Begin VB.Line Line5 
      X1              =   8250
      X2              =   3840
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line4 
      X1              =   3855
      X2              =   3855
      Y1              =   5685
      Y2              =   5910
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "KEMASAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3900
      TabIndex        =   19
      Top             =   5715
      Width           =   4320
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "KEMASAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7830
      TabIndex        =   18
      Top             =   1590
      Width           =   2235
   End
   Begin VB.Line Line3 
      X1              =   7800
      X2              =   7800
      Y1              =   1575
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   10095
      X2              =   7800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   10095
      X2              =   10095
      Y1              =   1575
      Y2              =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Kg"
      Height          =   255
      Left            =   11190
      TabIndex        =   17
      Top             =   5265
      Width           =   510
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   255
      Left            =   9615
      TabIndex        =   12
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label Label4 
      Caption         =   "Perolehan :"
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   5640
      Width           =   930
   End
   Begin VB.Label Label3 
      Caption         =   "Lot Number"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   1050
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Bahan Baku :"
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   1575
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   945
   End
End
Attribute VB_Name = "frmaddsop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String
Private kodestok As String
Private posrow As Integer

Private Sub cmbpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With grid2
            .TextMatrix(.Row, .Col) = cmbpass.text
            .SetFocus
        End With
    End If
    If KeyAscii = 27 Then
        cmbpass.Visible = False
    End If
End Sub

Private Sub cmbpass_LostFocus()
    cmbpass.Visible = False
End Sub

Private Sub cmdclear_Click()
    txtkdproduk = ""
    date1 = Date
    txtnolot = ""
    txtjumlah = 0
    hapusgrid1
    hapusgrid2
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdproduk_Click()
    carisql1 = "select * from am_itemcode where lev=3"
    namatabel = "Produk"
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduk_GotFocus()
    If hasil1 = "" Then Exit Sub
    txtkdproduk = hasil1
    txtnmproduk = hasil2
    carisql1 = ""
    namatabel = ""
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    OpenBahanBaku
End Sub

Private Sub cmdsimpan_Click()
    If txtkdproduk = "" Then
        MsgBox "Data Not Completed..", vbCritical, AppName
        Exit Sub
    End If
    If txtnolot = "" Then
        MsgBox "Data Not Completed..", vbCritical, AppName
    End If
    If grid1.TextMatrix(1, 1) = "" Then
        MsgBox "Data Not Completed..", vbCritical, AppName
        Exit Sub
    End If
    grid1.Row = 1
    Do While True
        With grid1
            If grid1.TextMatrix(.Row, 1) = "" Then Exit Do
            If grid1.TextMatrix(.Row, 6) = "0.00" Then
                MsgBox "Data Not Completed..", vbCritical, AppName
                Exit Sub
            End If
            .Row = .Row + 1
        End With
    Loop
    
    'Save To Table Produksi Header
    SQL = "INSERT INTO am_approduksihdr ("
    SQL = SQL + "kdproduksi,"
    SQL = SQL + "tanggal,"
    SQL = SQL + "kdproduk,"
    SQL = SQL + "lot,"
    SQL = SQL + "total_bahanbaku,"
    SQL = SQL + "total_perolehan,"
    SQL = SQL + "username) "
    SQL = SQL + "VALUES('"
    SQL = SQL + txtnolot + "',"
    SQL = SQL + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'"
    SQL = SQL + txtkdproduk + "','"
    SQL = SQL + txtnolot + "',"
    SQL = SQL + "convert(money,'" + Format(txtjumlah, "general number") + "'),"
    SQL = SQL + "convert(money,'0'),'"
    SQL = SQL + UserOnline + "')"
    
    OBJ.Open dsn
    Set RS = OBJ.Execute(SQL)
    OBJ.Close
    
    grid1.Row = 1
    OBJ.Open dsn
    Do While True
        With grid1
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "INSERT INTO am_approduksiin ("
            SQL = SQL + "kdproduksi,"
            SQL = SQL + "kdbahan,"
            SQL = SQL + "lot,"
            SQL = SQL + "kdpckg,"
            SQL = SQL + "qty,"
            SQL = SQL + "kdsatuan,line) "
            SQL = SQL + "VALUES('"
            SQL = SQL + txtnolot + "','"
            SQL = SQL + .TextMatrix(.Row, 1) + "','"
            SQL = SQL + .TextMatrix(.Row, 3) + "','"
            SQL = SQL + .TextMatrix(.Row, 4) + "',"
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 6)) + "'),'"
            SQL = SQL + .TextMatrix(.Row, 7) + "',"
            SQL = SQL + "convert(numeric,'" + Format(.Row, "general number") + "')"
            SQL = SQL + ")"
            Set RS = OBJ.Execute(SQL)
            .Row = .Row + 1
        End With
    Loop
    OBJ.Close
    
    'save to tabel perolehan
    grid2.Row = 2
    OBJ.Open dsn
    Do While True
        With grid2
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "INSERT INTO am_approduksihasil ("
            SQL = SQL + "kdproduksi,"
            SQL = SQL + "kdbarang,"
            SQL = SQL + "lot,"
            SQL = SQL + "kdpckg,"
            SQL = SQL + "qty,"
            SQL = SQL + "total,"
            SQL = SQL + "perolehan,"
            SQL = SQL + "waktu_reaksi,"
            SQL = SQL + "waktu_kemas,"
            SQL = SQL + "qc_cps,"
            SQL = SQL + "qc_solid,"
            SQL = SQL + "status,"
            SQL = SQL + "line) "
            SQL = SQL + "VALUES('"
            SQL = SQL + txtnolot + "','"
            SQL = SQL + .TextMatrix(.Row, 1) + "','"
            SQL = SQL + txtnolot + "','"
            SQL = SQL + .TextMatrix(.Row, 3) + "',"
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 5), "general number") + "')," 'qty
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "')," 'total
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 7), "general number") + "')," 'perolehan
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 8), "general number") + "')," 'waktu reaksi
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 9), "general number") + "')," 'waktu kemas
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 10), "general number") + "')," 'qc pcs
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 11), "general number") + "'),'" 'qc solid
            SQL = SQL + .TextMatrix(.Row, 12) + "',"
            SQL = SQL + "convert(numeric,'" + Format(.Row, "general number") + "')"
            SQL = SQL + ")"
            Set RS = OBJ.Execute(SQL)
            .Row = .Row + 1
        End With
        DoEvents
    Loop
    OBJ.Close
    
    'save to stok bahan baku
    'save to stok
    OBJ.Open dsn
    grid1.Row = 1
    kodestok = AmbilKodeStokBaru(Format(date1, "yy.MM"))
    Do While True
        With grid1
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "INSERT INTO am_stokbahan ("
            SQL = SQL + "kdstok,"
            SQL = SQL + "kdbarang,"
            SQL = SQL + "nolot,"
            SQL = SQL + "trans,"
            SQL = SQL + "ref,"
            SQL = SQL + "line,"
            SQL = SQL + "tgl,"
            SQL = SQL + "kdpckg,"
            SQL = SQL + "kdsatuan,"
            SQL = SQL + "awal,"
            SQL = SQL + "h_awal,"
            SQL = SQL + "masuk,"
            SQL = SQL + "h_masuk,"
            SQL = SQL + "keluar,"
            SQL = SQL + "h_keluar"
            SQL = SQL + ") "
            SQL = SQL + " Values('"
            SQL = SQL + kodestok + "','"  'kdstok
            SQL = SQL + .TextMatrix(.Row, 1) + "','" 'kdbarang
            SQL = SQL + .TextMatrix(.Row, 3) + "','"  'nolot
            SQL = SQL + "K" + "','"
            SQL = SQL + txtnolot + "',"  'ref
            SQL = SQL + "convert(numeric, '" + Format(.Row, "general number") + "')," 'line
            SQL = SQL + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'" 'tgl
            SQL = SQL + .TextMatrix(.Row, 4) + "','" 'satuan package
            SQL = SQL + .TextMatrix(.Row, 7) + "'," 'satuan
            SQL = SQL + "convert(money,'0')," 'awal
            SQL = SQL + "convert(money,'0')," 'h_awal
            SQL = SQL + "convert(money,'0')," 'masuk
            SQL = SQL + "convert(money,'0')," 'h_masuk
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "')," 'keluar
            SQL = SQL + "convert(money,'0')"  'h_keluar
            SQL = SQL + ")"
            
            Set RS = OBJ.Execute(SQL)
            .Row = .Row + 1
        End With
        DoEvents
    Loop
    OBJ.Close

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    
End Sub

Private Sub Form_Load()
    date1 = Date
    'grid1
    With grid1
        .Cols = 9
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "NO"
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "KD BAHAN"
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "BAHAN"
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .TextMatrix(0, 3) = "LOT NUMBER"
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "K/Pck"
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Packaging"
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Qty"
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .TextMatrix(0, 7) = "K/Sat"
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .TextMatrix(0, 8) = "Satuan"
        
    End With
    SetGrid1
    
    'grid2
    With grid2
        .Cols = 13
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "O"
        .TextMatrix(1, 0) = "O"
        .MergeCol(0) = True
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "Kode Barang"
        .TextMatrix(1, 1) = "Kode Barang"
        .MergeCol(1) = True
        
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Barang"
        .TextMatrix(1, 2) = "Barang"
        .MergeCol(2) = True
        
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Kode Kemasan"
        .TextMatrix(1, 3) = "Kode Kemasan"
        .MergeCol(3) = True
        
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Kemasan"
        .TextMatrix(1, 4) = "Kemasan"
        .MergeCol(4) = True
        
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Qty"
        .TextMatrix(1, 5) = "Qty"
        .MergeCol(5) = True
        
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Total (Kg)"
        .TextMatrix(1, 6) = "Total (Kg)"
        .MergeCol(6) = True
        
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .TextMatrix(0, 7) = "Perolehan (%)"
        .TextMatrix(1, 7) = "Perolehan (%)"
        .MergeCol(7) = True
        
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .TextMatrix(0, 8) = "Waktu Reaksi (Jam)"
        .TextMatrix(1, 8) = "Waktu Reaksi (Jam)"
        .MergeCol(8) = True
        
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .TextMatrix(0, 9) = "Waktu Kemasan (Jam)"
        .TextMatrix(1, 9) = "Waktu Kemasan (Jam)"
        .MergeCol(9) = True
        
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .TextMatrix(0, 10) = "QC"
        .TextMatrix(0, 11) = "QC"
        .MergeRow(0) = True
        
        .TextMatrix(1, 10) = "CPS"
        .TextMatrix(1, 11) = "Solid (%)"
        
        .ColAlignmentFixed(12) = flexAlignCenterCenter
        .TextMatrix(0, 12) = "STATUS"
        .TextMatrix(1, 12) = "PAS/TIDAK"
    End With
    
    SetGrid2
End Sub

Private Sub SetGrid1()
    With grid1
        .RowHeightMin = 300
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 4000
        .ColWidth(3) = 2000
        .ColWidth(4) = 800
        .ColWidth(5) = 1500
        .ColWidth(6) = 1000
        .ColWidth(7) = 800
        .ColWidth(8) = 1000
        
    End With
End Sub

Private Sub SetGrid2()
    With grid2
        .RowHeightMin = 300
        .ColWidth(0) = 250
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 800
        .ColWidth(4) = 2000
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 1000
    End With
End Sub

Private Sub OpenBahanBaku()
    SQL = "select a.*,b.namabarang,c.kodesatuan,c.namasatuan from am_apresepin a inner join am_apitemmst b on b.kodebarang=a.kdbahan "
    SQL = SQL + "inner join am_apunit c on c.kodesatuan = b.kodesatuanmutasi "
    SQL = SQL + " where a.kdproduk='" + txtkdproduk + "' order by a.line asc"
    OBJ.Open dsn
    
    Set RS = OBJ.Execute(SQL)
    Do While Not RS.EOF
        With grid1
            .TextMatrix(.Row, 0) = RS!Line
            .TextMatrix(.Row, 1) = RS!kdbahan
            .TextMatrix(.Row, 2) = RS!namabarang
            .TextMatrix(.Row, 7) = RS!kodesatuan
            .TextMatrix(.Row, 8) = RS!namasatuan
            .TextMatrix(.Row, 6) = "0.00"
            .Rows = .Rows + 1
            .Row = .Row + 1
        End With
        RS.MoveNext
        DoEvents
    Loop
    OBJ.Close
End Sub


Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtkdproduk = "" Then Exit Sub
    
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 3:
                With grid1
                If .TextMatrix(.Row, 1) = "" Then Exit Sub
                OBJ.Open dsn
                SQL = "exec am_stokbahanbaku"
                Set RS = OBJ.Execute(SQL)
                OBJ.Close
            
                carisql1 = "select * from am_tempstokbahanbaku where kdbarang='" + .TextMatrix(.Row, 1) + "'"
                namatabel = "Stock Bahan Baku Per Kode"
                
                frmsearch.Show vbModal
                    
                End With
        Case 6:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai.Left = grid1.Left + grid1.CellLeft
            txtnilai.Top = grid1.Top + grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
    
    
    
End Sub

Private Sub grid1_EnterCell()
    posrow = grid1.Row
    
    Select Case grid1.Col
      
        Case 6:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai.Left = grid1.Left + grid1.CellLeft
            txtnilai.Top = grid1.Top + grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    
    
    Select Case grid1.Col
        Case 3:
                With grid1
                    .TextMatrix(.Row, 3) = hasil2
                    .TextMatrix(.Row, 4) = hasil3
                    .TextMatrix(.Row, 5) = hasil4
                End With
                hasil = ""
                hasil1 = ""
                hasil2 = ""
                hasil3 = ""
                hasil4 = ""
                carisql1 = ""
                namatabel = ""
    End Select
    
    
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtkdproduk = "" Or txtnolot = "" Then Exit Sub
    
    posrow = grid2.Row
    Select Case grid2.Col
        Case 0:
        Case 1:
                With grid2
                    If .TextMatrix(.Row, 1) <> "" Then Exit Sub
                    carisql1 = "select * from am_itemcode where (lev=3 or lev=4)"
                    namatabel = "Produk"
                    frmsearch.Show vbModal
                End With
        Case 3:
              If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                carisql1 = "select * from am_appackaging"
                namatabel = "Packaging"
    
                frmsearch.Show vbModal
        Case 5, 6, 7, 8, 9, 10, 11:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            With txtnilai2
                .Width = grid2.ColWidth(grid2.Col) - 40
                .Value = grid2.TextMatrix(grid2.Row, grid2.Col)
                .Left = grid2.Left + grid2.CellLeft
                .Top = grid2.Top + grid2.CellTop + 20
                .Visible = True
                .SetFocus
            End With
        Case 12:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            With cmbpass
                .Width = grid2.ColWidth(grid2.Col) - 40
                .text = grid2.TextMatrix(grid2.Row, grid2.Col)
                .Left = grid2.Left + grid2.CellLeft
                .Top = grid2.Top + grid2.CellTop + 20
                .Visible = True
                .SetFocus
            End With
    End Select
End Sub

Private Sub grid2_EnterCell()
    posrow = grid2.Row
    Select Case grid2.Col
       
        Case 5, 6, 7, 8, 9, 10, 11:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            With txtnilai2
                .Width = grid2.ColWidth(grid2.Col) - 40
                .Value = grid2.TextMatrix(grid2.Row, grid2.Col)
                .Left = grid2.Left + grid2.CellLeft
                .Top = grid2.Top + grid2.CellTop + 20
                .Visible = True
                .SetFocus
            End With
        Case 12:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            With cmbpass
                .Width = grid2.ColWidth(grid2.Col) - 40
                .text = grid2.TextMatrix(grid2.Row, grid2.Col)
                .Left = grid2.Left + grid2.CellLeft
                .Top = grid2.Top + grid2.CellTop + 20
                .Visible = True
                .SetFocus
            End With
    End Select
End Sub

Private Sub grid2_GotFocus()
    If hasil1 = "" Then Exit Sub
    
    Select Case grid2.Col
        Case 1:
            With grid2
                .Col = 0
                Set .CellPicture = uncheck
                .MergeCol(12) = False
                .TextMatrix(.Row, 1) = hasil1
                .TextMatrix(.Row, 2) = hasil2
                
                .TextMatrix(.Row, 5) = "0.00"
                .TextMatrix(.Row, 6) = "0.00"
                .TextMatrix(.Row, 7) = "0.00"
                .TextMatrix(.Row, 8) = "0.00"
                .TextMatrix(.Row, 9) = "0.00"
                .TextMatrix(.Row, 10) = "0.00"
                .TextMatrix(.Row, 11) = "0.00"
                .MergeCells = flexMergeRestrictAll
                .Rows = .Rows + 1
                
            End With
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            namatabel = ""
            carisql1 = ""
        Case 3:
            With grid2
                .TextMatrix(.Row, 3) = hasil
                .TextMatrix(.Row, 4) = hasil1
            End With
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
        grid1.TextMatrix(grid1.Row, grid1.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        grid1.Row = 1
        txtjumlah.Value = 0
                Do While True
                    With grid1
                        If .TextMatrix(.Row, 1) = "" Then Exit Do
                        txtjumlah.Value = txtjumlah.Value + Format(.TextMatrix(.Row, 6), "general number")
                        .Row = .Row + 1
                    End With
                Loop

        txtnilai.Visible = False
        grid1.SetFocus
        grid1.Row = posrow
        
    ElseIf KeyAscii = 27 Then
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With grid2
           
            .TextMatrix(.Row, .Col) = Format(txtnilai2, "###,###,##0.00")
            If .Col = 6 Then
                 If txtnilai2 > txtjumlah Then
                    MsgBox "Pegisian data > dari jumlah total pemakaian bahan baku...!", vbCritical, AppName
                    txtnilai2 = 0
                    txtnilai2.SetFocus
                    Exit Sub
                End If
                If txtnilai2 = 0 Then Exit Sub
                .TextMatrix(.Row, 7) = Format(txtnilai2.Value / txtjumlah.Value, "###,###,##0.00")
            End If
            txtnilai2 = 0
            txtnilai2.Visible = False
        
            .SetFocus
            .Row = posrow
        End With
    ElseIf KeyAscii = 27 Then
        txtnilai2.Visible = False
    End If
End Sub

Private Sub txtnilai2_LostFocus()
    txtnilai2.Visible = False
End Sub

Private Sub hapusgrid1()
   grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
       grid1.TextMatrix(grid1.Row, 0) = ""
       grid1.TextMatrix(grid1.Row, 1) = ""
       grid1.TextMatrix(grid1.Row, 2) = ""
       grid1.TextMatrix(grid1.Row, 3) = ""
       grid1.TextMatrix(grid1.Row, 4) = ""
       grid1.TextMatrix(grid1.Row, 5) = ""
       grid1.TextMatrix(grid1.Row, 6) = ""
       grid1.TextMatrix(grid1.Row, 7) = ""
       grid1.TextMatrix(grid1.Row, 8) = ""
     
       grid1.Col = 0
        Set grid1.CellPicture = blank
       grid1.Row = grid1.Row + 1
    Loop
    
   grid1.Rows = 2
    SetGrid1
End Sub

Private Sub hapusrow1()
    grid1.TextMatrix(grid1.Row, 1) = ""
    grid1.TextMatrix(grid1.Row, 2) = ""
    grid1.TextMatrix(grid1.Row, 3) = ""
    grid1.TextMatrix(grid1.Row, 4) = ""
    grid1.TextMatrix(grid1.Row, 5) = ""
    grid1.TextMatrix(grid1.Row, 6) = ""
    grid1.TextMatrix(grid1.Row, 7) = ""
    grid1.TextMatrix(grid1.Row, 8) = ""
  
    Do While True
        If grid1.TextMatrix(grid1.Row + 1, 1) = "" Then
            grid1.TextMatrix(grid1.Row, 1) = ""
            grid1.TextMatrix(grid1.Row, 2) = ""
            grid1.TextMatrix(grid1.Row, 3) = ""
            grid1.TextMatrix(grid1.Row, 4) = ""
            grid1.TextMatrix(grid1.Row, 5) = ""
            grid1.TextMatrix(grid1.Row, 6) = ""
            grid1.TextMatrix(grid1.Row, 7) = ""
            grid1.TextMatrix(grid1.Row, 8) = ""
           
            Exit Do
        End If
        grid1.TextMatrix(grid1.Row, 1) = grid1.TextMatrix(grid1.Row + 1, 1)
        grid1.TextMatrix(grid1.Row, 2) = grid1.TextMatrix(grid1.Row + 1, 2)
        grid1.TextMatrix(grid1.Row, 3) = grid1.TextMatrix(grid1.Row + 1, 3)
        grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row + 1, 4)
        grid1.TextMatrix(grid1.Row, 5) = grid1.TextMatrix(grid1.Row + 1, 5)
        grid1.TextMatrix(grid1.Row, 6) = grid1.TextMatrix(grid1.Row + 1, 6)
        grid1.TextMatrix(grid1.Row, 7) = grid1.TextMatrix(grid1.Row + 1, 7)
        grid1.TextMatrix(grid1.Row, 8) = grid1.TextMatrix(grid1.Row + 1, 8)
        grid1.TextMatrix(grid1.Row, 9) = grid1.TextMatrix(grid1.Row + 1, 9)
       
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = grid1.Rows - 1
    grid1.Col = 0
    Set grid1.CellPicture = blank
End Sub

Private Sub hapusgrid2()
    grid2.Row = 2
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.TextMatrix(grid2.Row, 8) = ""
        grid2.TextMatrix(grid2.Row, 9) = ""
        grid2.TextMatrix(grid2.Row, 10) = ""
        grid2.TextMatrix(grid2.Row, 11) = ""
        grid2.TextMatrix(grid2.Row, 12) = ""
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    
    grid2.Rows = 3
    SetGrid2
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
    grid2.TextMatrix(grid2.Row, 9) = ""
    grid2.TextMatrix(grid2.Row, 10) = ""
    grid2.TextMatrix(grid2.Row, 11) = ""
    grid2.TextMatrix(grid2.Row, 12) = ""
   
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
            grid2.TextMatrix(grid2.Row, 9) = ""
            grid2.TextMatrix(grid2.Row, 10) = ""
            grid2.TextMatrix(grid2.Row, 11) = ""
            grid2.TextMatrix(grid2.Row, 12) = ""
           
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
        grid2.TextMatrix(grid2.Row, 9) = grid2.TextMatrix(grid2.Row + 1, 9)
        grid2.TextMatrix(grid2.Row, 10) = grid2.TextMatrix(grid2.Row + 1, 10)
        grid2.TextMatrix(grid2.Row, 11) = grid2.TextMatrix(grid2.Row + 1, 11)
        grid2.TextMatrix(grid2.Row, 12) = grid2.TextMatrix(grid2.Row + 1, 12)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = blank
End Sub


