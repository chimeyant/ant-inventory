VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "chameleon.ocx"
Begin VB.Form frmchangesop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubah SOP"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6855
      TabIndex        =   30
      Text            =   "Text4"
      Top             =   1740
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   9810
      TabIndex        =   29
      Top             =   1170
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   9735
      TabIndex        =   28
      Top             =   690
      Width           =   960
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   6855
      TabIndex        =   27
      Text            =   "Text3"
      Top             =   1170
      Width           =   2730
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   6885
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   690
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6870
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   165
      Width           =   2775
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   420
      Left            =   12210
      TabIndex        =   14
      Top             =   7710
      Width           =   1050
   End
   Begin VB.TextBox txtkdproduk 
      Height          =   330
      Left            =   1095
      TabIndex        =   13
      Top             =   0
      Width           =   1245
   End
   Begin VB.TextBox txtnmproduk 
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   0
      Width           =   3810
   End
   Begin VB.TextBox txtnolot 
      Height          =   315
      Left            =   1095
      TabIndex        =   9
      Top             =   750
      Width           =   5070
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      Height          =   420
      Left            =   11145
      TabIndex        =   7
      Top             =   7710
      Width           =   1050
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   420
      Left            =   10080
      TabIndex        =   6
      Top             =   7710
      Width           =   1050
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
      Left            =   12255
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   465
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
      Left            =   7710
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   30
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
      Left            =   7230
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbpass 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11610
      TabIndex        =   0
      Text            =   "PASS"
      Top             =   765
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai2 
      Height          =   255
      Left            =   11760
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   450
      Calculator      =   "frmchangesop.frx":0000
      Caption         =   "frmchangesop.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmchangesop.frx":008C
      Keys            =   "frmchangesop.frx":00AA
      Spin            =   "frmchangesop.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   12315
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   450
      Calculator      =   "frmchangesop.frx":011C
      Caption         =   "frmchangesop.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmchangesop.frx":01A8
      Keys            =   "frmchangesop.frx":01C6
      Spin            =   "frmchangesop.frx":0210
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   1890
      Left            =   0
      TabIndex        =   8
      Top             =   5745
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   1170
      Left            =   30
      TabIndex        =   10
      Top             =   3810
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   2064
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   1095
      TabIndex        =   11
      Top             =   390
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   130940929
      CurrentDate     =   41357
   End
   Begin Chameleon.chameleonButton cmdproduk 
      Height          =   285
      Left            =   60
      TabIndex        =   15
      Top             =   45
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
      MICON           =   "frmchangesop.frx":0238
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
      Left            =   10005
      TabIndex        =   16
      Top             =   5040
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   503
      Calculator      =   "frmchangesop.frx":0552
      Caption         =   "frmchangesop.frx":0572
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmchangesop.frx":05DE
      Keys            =   "frmchangesop.frx":05FC
      Spin            =   "frmchangesop.frx":0646
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.000;(#,###,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.000;(#,###,###,##0.000)"
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   225
      Left            =   45
      TabIndex        =   24
      Top             =   450
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Bahan Baku :"
      Height          =   255
      Left            =   30
      TabIndex        =   23
      Top             =   1365
      Width           =   1050
   End
   Begin VB.Label Label3 
      Caption         =   "Lot Number"
      Height          =   225
      Left            =   45
      TabIndex        =   22
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "Perolehan :"
      Height          =   255
      Left            =   15
      TabIndex        =   21
      Top             =   5430
      Width           =   930
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   255
      Left            =   9540
      TabIndex        =   20
      Top             =   5070
      Width           =   420
   End
   Begin VB.Label Label6 
      Caption         =   "Kg"
      Height          =   255
      Left            =   11115
      TabIndex        =   19
      Top             =   5055
      Width           =   510
   End
   Begin VB.Line Line1 
      X1              =   10020
      X2              =   10020
      Y1              =   1365
      Y2              =   1605
   End
   Begin VB.Line Line2 
      X1              =   10320
      X2              =   8025
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Line Line3 
      X1              =   7725
      X2              =   7725
      Y1              =   1365
      Y2              =   1590
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
      Left            =   8070
      TabIndex        =   18
      Top             =   3870
      Width           =   2235
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
      Left            =   3825
      TabIndex        =   17
      Top             =   5505
      Width           =   4320
   End
   Begin VB.Line Line4 
      X1              =   3780
      X2              =   3780
      Y1              =   5475
      Y2              =   5700
   End
   Begin VB.Line Line5 
      X1              =   8175
      X2              =   3765
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line6 
      X1              =   8175
      X2              =   8175
      Y1              =   5475
      Y2              =   5715
   End
End
Attribute VB_Name = "frmchangesop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private obj As New ADODB.Connection
Private RS As ADODB.Recordset
Private sql As String
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
    txtnmproduk = ""
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
    carisql1 = "select * from am_itemcode where lev=2 or lev=3"
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
            If grid1.TextMatrix(.Row, 6) = "0.000" Then
                If (MsgBox("Bahan Baku ada yang bernilai 0...!," & Chr(13) & " Apakah anda yakin akan melanjutkan proses penyimpanan ..?", vbQuestion + vbYesNo) = vbNo) Then
                    Exit Sub
                End If
            End If
            .Row = .Row + 1
        End With
    Loop
    
    'Save To Table Produksi Header
    sql = "INSERT INTO am_approduksihdr ("
    sql = sql + "kdproduksi,"
    sql = sql + "tanggal,"
    sql = sql + "kdproduk,"
    sql = sql + "lot,"
    sql = sql + "total_bahanbaku,"
    sql = sql + "total_perolehan,"
    sql = sql + "username) "
    sql = sql + "VALUES('"
    sql = sql + txtnolot + "',"
    sql = sql + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'"
    sql = sql + txtkdproduk + "','"
    sql = sql + txtnolot + "',"
    sql = sql + "convert(money,'" + Format(txtjumlah, "general number") + "'),"
    sql = sql + "convert(money,'0'),'"
    sql = sql + UserOnline + "')"
    
    obj.Open dsn
    Set RS = obj.Execute(sql)
    obj.Close
    
    grid1.Row = 1
    obj.Open dsn
    Do While True
        With grid1
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            sql = "INSERT INTO am_approduksiin ("
            sql = sql + "kdproduksi,"
            sql = sql + "kdbahan,"
            sql = sql + "lot,"
            sql = sql + "kdpckg,"
            sql = sql + "qty,"
            sql = sql + "kdsatuan,line) "
            sql = sql + "VALUES('"
            sql = sql + txtnolot + "','"
            sql = sql + .TextMatrix(.Row, 1) + "','"
            sql = sql + .TextMatrix(.Row, 3) + "','"
            sql = sql + .TextMatrix(.Row, 4) + "',"
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 6)) + "'),'"
            sql = sql + .TextMatrix(.Row, 7) + "',"
            sql = sql + "convert(numeric,'" + Format(.Row, "general number") + "')"
            sql = sql + ")"
            Set RS = obj.Execute(sql)
            .Row = .Row + 1
        End With
    Loop
    obj.Close
    
    'save to tabel perolehan
    grid2.Row = 2
    obj.Open dsn
    Do While True
        With grid2
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            sql = "INSERT INTO am_approduksihasil ("
            sql = sql + "kdproduksi,"
            sql = sql + "kdbarang,"
            sql = sql + "lot,"
            sql = sql + "kdpckg,"
            sql = sql + "qty,"
            sql = sql + "total,"
            sql = sql + "perolehan,"
            sql = sql + "waktu_reaksi,"
            sql = sql + "waktu_kemas,"
            sql = sql + "qc_cps,"
            sql = sql + "qc_solid,"
            sql = sql + "status,"
            sql = sql + "line) "
            sql = sql + "VALUES('"
            sql = sql + txtnolot + "','"
            sql = sql + .TextMatrix(.Row, 1) + "','"
            sql = sql + txtnolot + "','"
            sql = sql + .TextMatrix(.Row, 3) + "',"
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 5), "general number") + "')," 'qty
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "')," 'total
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 7), "general number") + "')," 'perolehan
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 8), "general number") + "')," 'waktu reaksi
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 9), "general number") + "')," 'waktu kemas
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 10), "general number") + "')," 'qc pcs
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 11), "general number") + "'),'" 'qc solid
            sql = sql + .TextMatrix(.Row, 12) + "',"
            sql = sql + "convert(numeric,'" + Format(.Row, "general number") + "')"
            sql = sql + ")"
            Set RS = obj.Execute(sql)
            .Row = .Row + 1
        End With
        DoEvents
    Loop
    obj.Close
    
    'save to stok bahan baku
    'save to stok
    obj.Open dsn
    grid1.Row = 1
    kodestok = AmbilKodeStokBaru(Format(date1, "yy.MM"))
    Do While True
        With grid1
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            sql = "INSERT INTO am_stokbahan ("
            sql = sql + "kdstok,"
            sql = sql + "kdbarang,"
            sql = sql + "nolot,"
            sql = sql + "trans,"
            sql = sql + "ref,"
            sql = sql + "line,"
            sql = sql + "tgl,"
            sql = sql + "kdpckg,"
            sql = sql + "kdsatuan,"
            sql = sql + "awal,"
            sql = sql + "h_awal,"
            sql = sql + "masuk,"
            sql = sql + "h_masuk,"
            sql = sql + "keluar,"
            sql = sql + "h_keluar"
            sql = sql + ") "
            sql = sql + " Values('"
            sql = sql + kodestok + "','"  'kdstok
            sql = sql + .TextMatrix(.Row, 1) + "','" 'kdbarang
            sql = sql + .TextMatrix(.Row, 3) + "','"  'nolot
            sql = sql + "K" + "','"
            sql = sql + txtnolot + "',"  'ref
            sql = sql + "convert(numeric, '" + Format(.Row, "general number") + "')," 'line
            sql = sql + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'" 'tgl
            sql = sql + .TextMatrix(.Row, 4) + "','" 'satuan package
            sql = sql + .TextMatrix(.Row, 7) + "'," 'satuan
            sql = sql + "convert(money,'0')," 'awal
            sql = sql + "convert(money,'0')," 'h_awal
            sql = sql + "convert(money,'0')," 'masuk
            sql = sql + "convert(money,'0')," 'h_masuk
            sql = sql + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "')," 'keluar
            sql = sql + "convert(money,'0')"  'h_keluar
            sql = sql + ")"
            
            Set RS = obj.Execute(sql)
            .Row = .Row + 1
        End With
        DoEvents
    Loop
    obj.Close

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    
End Sub

Private Sub Command1_Click()
    Dim namabarang As String
    Dim namasatuan As String
    Dim qty_stok As Double
    GetStokBarang Format(Date, "yyyyMMdd"), Text1.text, namabarang, namasatuan, qty_stok
    Text2 = qty_stok
    
End Sub

Private Sub Command2_Click()
 Text4 = getHPP(Text1, Text2, Text3)
End Sub

Private Sub Form_Load()
    date1 = Date
    
    With cmbpass
        .AddItem "PAS"
        .AddItem "TIDAK"
        .text = "PAS"
    End With
    
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
    sql = "select a.*,b.namabarang,c.kodesatuan,c.namasatuan from am_apresepin a inner join am_apitemmst b on b.kodebarang=a.kdbahan "
    sql = sql + "inner join am_apunit c on c.kodesatuan = b.kodesatuanmutasi "
    sql = sql + " where a.kdproduk='" + txtkdproduk + "' order by a.line asc"
    obj.Open dsn
    
    Set RS = obj.Execute(sql)
    Do While Not RS.EOF
        With grid1
            .TextMatrix(.Row, 0) = RS!Line
            .TextMatrix(.Row, 1) = RS!kdbahan
            .TextMatrix(.Row, 2) = RS!namabarang
            .TextMatrix(.Row, 7) = RS!kodesatuan
            .TextMatrix(.Row, 8) = RS!namasatuan
            .TextMatrix(.Row, 6) = "0.000"
            .Rows = .Rows + 1
            .Row = .Row + 1
        End With
        RS.MoveNext
        DoEvents
    Loop
    obj.Close
End Sub


Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtkdproduk = "" Then Exit Sub
    
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 3:
                With grid1
                If .TextMatrix(.Row, 1) = "" Then Exit Sub
                obj.Open dsn
                sql = "exec am_stokbahanbaku"
                Set RS = obj.Execute(sql)
                obj.Close
            
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
                If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow2
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
            End If
        Case 1:
                With grid2
                    If .TextMatrix(.Row, 1) <> "" Then Exit Sub
                    carisql1 = "select * from am_itemcode where (lev=2 or lev=3 or lev=4)"
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
                
                .TextMatrix(.Row, 5) = "0.000"
                .TextMatrix(.Row, 6) = "0.000"
                .TextMatrix(.Row, 7) = "0.000"
                .TextMatrix(.Row, 8) = "0.000"
                .TextMatrix(.Row, 9) = "0.000"
                .TextMatrix(.Row, 10) = "0.000"
                .TextMatrix(.Row, 11) = "0.000"
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
        grid1.TextMatrix(grid1.Row, grid1.Col) = Format(txtnilai, "###,###,##0.000")
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
           
            .TextMatrix(.Row, .Col) = Format(txtnilai2, "###,###,##0.000")
            If .Col = 6 Then
                 If txtnilai2 > txtjumlah Then
                    MsgBox "Pegisian data > dari jumlah total pemakaian bahan baku...!", vbCritical, AppName
                    txtnilai2 = 0
                    txtnilai2.SetFocus
                    Exit Sub
                End If
                If txtnilai2 = 0 Then Exit Sub
                .TextMatrix(.Row, 7) = Format((txtnilai2.Value / txtjumlah.Value) * 100, "###,###,##0.000")
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



