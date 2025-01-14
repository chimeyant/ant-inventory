VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmeditpalet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT LOT/PALET"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl2 
      Height          =   3105
      Left            =   75
      TabIndex        =   9
      Top             =   3480
      Width           =   10410
      _Version        =   851970
      _ExtentX        =   18362
      _ExtentY        =   5477
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
      ItemCount       =   1
      Item(0).Caption =   "Pemakaian Kaleng dan Kemasan"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "grid2"
      Item(0).Control(1)=   "btnupdate_package"
      Item(0).Control(2)=   "txtqty"
      Begin XtremeSuiteControls.PushButton btnupdate_package 
         Height          =   540
         Left            =   9390
         TabIndex        =   16
         Top             =   420
         Width           =   930
         _Version        =   851970
         _ExtentX        =   1640
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Update"
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
      Begin TDBNumber6Ctl.TDBNumber txtqty 
         Height          =   255
         Left            =   7980
         TabIndex        =   19
         Top             =   555
         Visible         =   0   'False
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   450
         Calculator      =   "frmeditpalet.frx":0000
         Caption         =   "frmeditpalet.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmeditpalet.frx":008C
         Keys            =   "frmeditpalet.frx":00AA
         Spin            =   "frmeditpalet.frx":00EC
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   2520
         Left            =   120
         TabIndex        =   10
         Top             =   435
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4445
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   12632256
         BackColorBkg    =   4210752
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
      Left            =   3975
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   165
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
      Left            =   3720
      Picture         =   "frmeditpalet.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   165
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
      Left            =   3465
      Picture         =   "frmeditpalet.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2445
      Left            =   75
      TabIndex        =   2
      Top             =   1035
      Width           =   10410
      _Version        =   851970
      _ExtentX        =   18362
      _ExtentY        =   4313
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
      ItemCount       =   1
      Item(0).Caption =   "Perolehan Produksi"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "grid1"
      Item(0).Control(1)=   "btnupdate_hasil"
      Item(0).Control(2)=   "txthasil"
      Item(0).Control(3)=   "lblflag"
      Item(0).Control(4)=   "lblproses"
      Begin XtremeSuiteControls.PushButton btnupdate_hasil 
         Height          =   540
         Left            =   9360
         TabIndex        =   15
         Top             =   420
         Width           =   930
         _Version        =   851970
         _ExtentX        =   1640
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Update"
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
      Begin TDBNumber6Ctl.TDBNumber txthasil 
         Height          =   255
         Left            =   7980
         TabIndex        =   18
         Top             =   510
         Visible         =   0   'False
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   450
         Calculator      =   "frmeditpalet.frx":0744
         Caption         =   "frmeditpalet.frx":0764
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmeditpalet.frx":07D0
         Keys            =   "frmeditpalet.frx":07EE
         Spin            =   "frmeditpalet.frx":0830
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   1875
         Left            =   120
         TabIndex        =   5
         Top             =   435
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3307
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   12632256
         BackColorBkg    =   4210752
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblproses 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9480
         TabIndex        =   23
         Top             =   1455
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblflag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9465
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.TextBox txtpalet 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   435
      Left            =   9465
      TabIndex        =   1
      Top             =   6660
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "CLOSE"
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
   Begin XtremeSuiteControls.DateTimePicker Dtptglscan 
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   1515
      _Version        =   851970
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.PushButton cmdpalet 
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "PALET : "
      BackColor       =   16777215
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnclear 
      Height          =   435
      Left            =   8385
      TabIndex        =   20
      Top             =   6660
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "CLEAR"
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
   Begin XtremeSuiteControls.PushButton cmdupdatetgl 
      Height          =   420
      Left            =   6000
      TabIndex        =   26
      Top             =   120
      Width           =   1410
      _Version        =   851970
      _ExtentX        =   2487
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Update tgl scan"
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
      Alignment       =   1  'Right Justify
      Caption         =   "GUDANG :"
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
      Left            =   7425
      TabIndex        =   25
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label lblgudang 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8880
      TabIndex        =   24
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label lbllot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4170
      TabIndex        =   21
      Top             =   150
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      Height          =   285
      Left            =   1875
      TabIndex        =   17
      Top             =   675
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TANGGAL SCAN :"
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
      Left            =   7425
      TabIndex        =   14
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "PRODUK :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   255
      TabIndex        =   13
      Top             =   660
      Width           =   795
   End
   Begin VB.Label lblproduk 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   2085
      TabIndex        =   12
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label lblkode 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1290
      TabIndex        =   11
      Top             =   660
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   3285
   End
End
Attribute VB_Name = "frmeditpalet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset
Private SQL1 As String
Private akses As Boolean


Private Sub btnclear_Click()
    Dim i As Integer
    txtpalet = ""
    lblkode = ""
    lblproduk = ""
    lbllot = ""
    Dtptglscan = Date
    hapusgrid1
    hapusgrid2
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnupdate_hasil_Click()
'list_produksi_hasil
'list_mutasi_produksi_details
'am_bpblin
'am_stok
    If grid1.TextMatrix(1, 1) = "" Then Exit Sub
    If MsgBox("Anda yakin ingin merubah data ini", vbQuestion + vbYesNo, "KONFIRMASI EDIT DATA PEROLEHAN") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "Update list_produksi_hasil SET"
    SQL = SQL + " kode_bahan='" & grid1.TextMatrix(1, 1) & "',"
    SQL = SQL + " qty_bahan= convert(money,'" & Format(grid1.TextMatrix(1, 4), "general number") & "'),"
    SQL = SQL + " kode_satuan='" & grid1.TextMatrix(1, 5) & "'"
    SQL = SQL + " WHERE noref='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "Update am_stok SET"
    SQL = SQL + " kodebarang='" & grid1.TextMatrix(1, 1) & "',"
    SQL = SQL + " namabarang='" & grid1.TextMatrix(1, 2) & "',"
    SQL = SQL + " qtyin= convert(money,'" & Format(grid1.TextMatrix(1, 4), "general number") & "'),"
    SQL = SQL + " kodesatuan='" & grid1.TextMatrix(1, 5) & "',"
    SQL = SQL + " kg='" & grid1.TextMatrix(1, 7) & "',"
    SQL = SQL + " isi='" & grid1.TextMatrix(1, 8) & "'"
    SQL = SQL + " WHERE palet='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)

    SQL = "Update list_mutasi_produksi_details SET"
    SQL = SQL + " kode_barang='" & grid1.TextMatrix(1, 1) & "',"
    SQL = SQL + " kode_satuan='" & grid1.TextMatrix(1, 5) & "',"
    SQL = SQL + " qty= convert(money,'" & Format(grid1.TextMatrix(1, 4), "general number") & "')"
    SQL = SQL + " WHERE kode_palet='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)

    SQL = "Update am_bpblin SET"
    SQL = SQL + " kodebarang='" & grid1.TextMatrix(1, 1) & "',"
    SQL = SQL + " qty= convert(money,'" & Format(grid1.TextMatrix(1, 4), "general number") & "'),"
    SQL = SQL + " kodesatuan='" & grid1.TextMatrix(1, 5) & "'"
    SQL = SQL + " WHERE keterangan='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)

    SQL = "Update am_bpblin SET qty= convert(money,'" & Format(grid1.TextMatrix(1, 4), "general number") * -1 & "')"
    SQL = SQL + " WHERE keterangan='" & txtpalet & "' and type='99'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Perolehan data is successfuly updated", vbInformation, AppName
    OBJ.Close
    btnclear_Click
End Sub

Private Sub btnupdate_package_Click()
'list_produksi_kemasan
'am_uselin
    If grid2.TextMatrix(1, 1) = "" Then Exit Sub
    If MsgBox("Anda yakin ingin merubah data ini", vbQuestion + vbYesNo, "KONFIRMASI EDIT DATA KEMASAN") = vbNo Then Exit Sub
        
        OBJ.Open dsn
        
        'UPDATE LIST_PRODUKSI_KEMASAN
        SQL = "Delete From list_produksi_kemasan Where noref='" & txtpalet & "'"
        OBJ.Execute SQL
        
        SQL = "Select * From list_produksi_kemasan Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        grid2.Row = 1
        Do While True
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!kode_produk = lblkode
            RST!nolot = lbllot
            RST!kode_bahan = grid2.TextMatrix(grid2.Row, 1)
            RST!Lot_bahan = ""
            RST!qty_bahan = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
            RST!KODE_SATUAN = grid2.TextMatrix(grid2.Row, 5)
            RST!flag_tambahan = "0"
            RST!hpp = Format(grid2.TextMatrix(grid2.Row, 7), "general number")
            RST!tanggal = Format(Dtptglscan, "yyyy/MM/dd")
            RST!noref = txtpalet
            RST!proses_ke = "0"
            RST.Update
            grid2.Row = grid2.Row + 1
        Loop
        
        'UPDATE AM_USEHDR
        SQL = "Delete From am_usehdr Where nobpb='" & txtpalet & "'"
        OBJ.Execute SQL
        
        SQL = "Select * From am_usehdr Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        With RST
            .AddNew
            !nobpb = txtpalet
            !tglbpb = Format(Dtptglscan, "yyyy/MM/dd")
            !noorder = txtpalet
            .Update
        End With
        'UPDATE AM_USELIN
        SQL = "Delete From am_uselin Where nobpb='" & txtpalet & "'"
        OBJ.Execute SQL
        
        SQL = "Select * From am_uselin Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        grid2.Row = 1
        Do While True
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!nobpb = txtpalet
            RST!kodebarang = grid2.TextMatrix(grid2.Row, 1)
            RST!qty = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
            RST!kodesatuan = grid2.TextMatrix(grid2.Row, 5)
            RST!lineitem = grid2.Row
            RST.Update
            grid2.Row = grid2.Row + 1
        Loop
        
    MsgBox "Packaging data was successfully updated", vbInformation, AppName
    OBJ.Close
    btnclear_Click
End Sub

Private Sub cmdupdatetgl_Click()
    If txtpalet = "" Then Exit Sub
    If MsgBox("Yakin ingin merubah tanggal scan ?", vbQuestion + vbYesNo, "Confirm Update") = vbNo Then Exit Sub
        OBJ.Open dsn
        SQL = "Update list_produksi_kemasan SET tanggal='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE noref='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update list_produksi_hasil SET tanggal='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE noref='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update list_mutasi_produksi_header SET tanggal='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE kode_palet='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_usehdr SET tglbpb='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE nobpb='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_bpbhdr SET tglbpb='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE noref='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_bpblin SET tglbpb='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " WHERE keterangan='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_stok SET tanggal='" & Format(Dtptglscan, "yyyyMMdd") & "'"
        SQL = SQL + " Where palet='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        
        MsgBox "Tanggal scan berhasil diupdate", vbInformation, AppName
        OBJ.Close
        btnclear_Click
End Sub

Private Sub Form_Load()
    'Periksa hak akses hpp
    OBJ.Open dsn
    SQL = "Select * From LIST_USERS Where username = '" & nmuser & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!gl = "1" Then
            akses = True
        Else
            akses = False
        End If
    Else
        If nmuser = "Creator" Then akses = True
    End If
    OBJ.Close
    
    Dtptglscan = Date
    setGrid1
    setGrid2
    initGrid1
    initGrid2
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    Select Case grid1.Col
        Case 0:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                If grid1.CellPicture = uncheck Then
                Set grid1.CellPicture = check
                If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid1.CellPicture = uncheck
                    hapusrow1
                    Exit Sub
                End If
                Set grid1.CellPicture = uncheck
            End If
        Case 1:     'ADD BARANG
            If grid1.TextMatrix(grid1.Row, 1) <> "" Then Exit Sub
            carisql1 = "select am_itemdtl.kodebarang, am_itemdtl.namabarang,list_produk_hasil.kode_satuan,am_unit.namasatuan from am_itemdtl  "
            carisql1 = carisql1 + " inner join list_produk_hasil on am_itemdtl.kodebarang=list_produk_hasil.kode_barang_jadi "
            carisql1 = carisql1 + " and am_itemdtl.kodesatuan = list_produk_hasil.kode_satuan and list_produk_hasil.kode_produk='" & lblkode & "' "
            carisql1 = carisql1 + " inner join am_unit on list_produk_hasil.kode_satuan= am_unit.kodesatuan "
            namatabel = "Barang Jadi"
            frmsearch.Show vbModal
        Case 2:     'UPDATE BARANG
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            carisql1 = "select am_itemdtl.kodebarang, am_itemdtl.namabarang,list_produk_hasil.kode_satuan,am_unit.namasatuan from am_itemdtl  "
            carisql1 = carisql1 + " inner join list_produk_hasil on am_itemdtl.kodebarang=list_produk_hasil.kode_barang_jadi "
            carisql1 = carisql1 + " and am_itemdtl.kodesatuan = list_produk_hasil.kode_satuan and list_produk_hasil.kode_produk='" & lblkode & "' "
            carisql1 = carisql1 + " inner join am_unit on list_produk_hasil.kode_satuan= am_unit.kodesatuan "
            namatabel = "Barang Jadi"
            frmsearch.Show vbModal
        Case 4:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            txthasil.Width = grid1.ColWidth(grid1.Col) - 40
            txthasil = grid1.TextMatrix(grid1.Row, grid1.Col)
            txthasil.Left = grid1.Left + grid1.CellLeft
            txthasil.Top = grid1.Top + grid1.CellTop + 20
            txthasil.Visible = True
            txthasil.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid1.Col
        Case 1:
            grid1.TextMatrix(grid1.Row, 1) = hasil
            grid1.TextMatrix(grid1.Row, 2) = hasil1
            grid1.TextMatrix(grid1.Row, 5) = hasil2
            OBJ.Open dsn
            SQL = "Select NamaSatuan From am_unit Where KodeSatuan= '" & hasil2 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
            
            'cari kg base unit
            SQL = "Select * From am_itemkg Where kodebarang = '" & hasil & "'"
            SQL = SQL + " and kodesatuan= '" & hasil2 & "' and tahun= '" & Right(Dtptglscan, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If Mid(Dtptglscan, 4, 2) = "01" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg1
                ElseIf Mid(Dtptglscan, 4, 2) = "02" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg2
                ElseIf Mid(Dtptglscan, 4, 2) = "03" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg3
                ElseIf Mid(Dtptglscan, 4, 2) = "04" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg4
                ElseIf Mid(Dtptglscan, 4, 2) = "05" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg5
                ElseIf Mid(Dtptglscan, 4, 2) = "06" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg6
                ElseIf Mid(Dtptglscan, 4, 2) = "07" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg7
                ElseIf Mid(Dtptglscan, 4, 2) = "08" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg8
                ElseIf Mid(Dtptglscan, 4, 2) = "09" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg9
                ElseIf Mid(Dtptglscan, 4, 2) = "10" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg10
                ElseIf Mid(Dtptglscan, 4, 2) = "11" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg11
                ElseIf Mid(Dtptglscan, 4, 2) = "12" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg12
                End If
            End If
            'cari konversi item
            SQL = "Select * From am_itemdtl Where KodeBarang = '" & hasil & "' and KodeSatuan = '" & hasil2 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then grid1.TextMatrix(grid1.Row, 8) = RST!konversi
            
            OBJ.Close
            hasil = "": hasil1 = "": hasil2 = ""
            grid1.Col = 0
            Set grid1.CellPicture = uncheck

            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
        Case 2:
            grid1.TextMatrix(grid1.Row, 1) = hasil
            grid1.TextMatrix(grid1.Row, 2) = hasil1
            grid1.TextMatrix(grid1.Row, 5) = hasil2
            OBJ.Open dsn
            SQL = "Select NamaSatuan From am_unit Where KodeSatuan= '" & hasil2 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
            
            'cari kg base unit
            SQL = "Select * From am_itemkg Where kodebarang = '" & hasil & "'"
            SQL = SQL + " and kodesatuan= '" & hasil2 & "' and tahun= '" & Right(Dtptglscan, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If Mid(Dtptglscan, 4, 2) = "01" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg1
                ElseIf Mid(Dtptglscan, 4, 2) = "02" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg2
                ElseIf Mid(Dtptglscan, 4, 2) = "03" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg3
                ElseIf Mid(Dtptglscan, 4, 2) = "04" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg4
                ElseIf Mid(Dtptglscan, 4, 2) = "05" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg5
                ElseIf Mid(Dtptglscan, 4, 2) = "06" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg6
                ElseIf Mid(Dtptglscan, 4, 2) = "07" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg7
                ElseIf Mid(Dtptglscan, 4, 2) = "08" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg8
                ElseIf Mid(Dtptglscan, 4, 2) = "09" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg9
                ElseIf Mid(Dtptglscan, 4, 2) = "10" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg10
                ElseIf Mid(Dtptglscan, 4, 2) = "11" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg11
                ElseIf Mid(Dtptglscan, 4, 2) = "12" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST!kg12
                End If
            End If
            'cari konversi item
            SQL = "Select * From am_itemdtl Where KodeBarang = '" & hasil & "' and KodeSatuan = '" & hasil2 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then grid1.TextMatrix(grid1.Row, 8) = RST!konversi
            
            OBJ.Close
            hasil = "": hasil1 = "": hasil2 = ""
    End Select
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    Select Case grid2.Col
        Case 0:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow2
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
            End If
        Case 1:     'ADD KEMASAN
            If grid2.TextMatrix(grid2.Row, 2) <> "" Then Exit Sub
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
        Case 2:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
        Case 4:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            txtqty.Width = grid2.ColWidth(grid2.Col) - 40
            txtqty = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtqty.Left = grid2.Left + grid2.CellLeft
            txtqty.Top = grid2.Top + grid2.CellTop + 20
            txtqty.Visible = True
            txtqty.SetFocus
    End Select
End Sub

Private Sub grid2_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid2.Col
        Case 1:
            grid2.TextMatrix(grid2.Row, 1) = hasil
            grid2.TextMatrix(grid2.Row, 2) = hasil1
            'cari satuan
            SQL = "select kodesatuan from am_apitemmst  "
            SQL = SQL + "where kodebarang='" & hasil & "'"
            OBJ.Open dsn
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid2.TextMatrix(grid2.Row, 4) = "0"
            grid2.TextMatrix(grid2.Row, 5) = RST!kodesatuan
                    
            'cari nama satuan
            SQL = "select * from am_apunit where kodesatuan ='" & grid2.TextMatrix(grid2.Row, 5) & "'"
            Set RST = OBJ.Execute(SQL)
            grid2.TextMatrix(grid2.Row, 6) = RST!namasatuan
            OBJ.Close
            
            hasil = "": hasil1 = ""
            grid2.Col = 0
            Set grid2.CellPicture = uncheck

            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
        Case 2:
            grid2.TextMatrix(grid2.Row, 1) = hasil
            grid2.TextMatrix(grid2.Row, 2) = hasil1
            'cari satuan
            SQL = "select kodesatuan from am_apitemmst  "
            SQL = SQL + "where kodebarang='" & hasil & "'"
            OBJ.Open dsn
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid2.TextMatrix(grid2.Row, 4) = "0"
            grid2.TextMatrix(grid2.Row, 5) = RST!kodesatuan
                    
            'cari nama satuan
            SQL = "select * from am_apunit where kodesatuan ='" & grid2.TextMatrix(grid2.Row, 5) & "'"
            Set RST = OBJ.Execute(SQL)
            grid2.TextMatrix(grid2.Row, 6) = RST!namasatuan
            OBJ.Close
            
            hasil = "": hasil1 = ""
            grid2.Col = 0
            Set grid2.CellPicture = uncheck
    End Select
End Sub

Private Sub txthasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, 4) = txthasil.text
        grid1.SetFocus
    End If
End Sub

Private Sub txthasil_LostFocus()
    txthasil.Visible = False
End Sub

Private Sub txtpalet_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        OBJ.Open dsn
        'OPEN PEROLEHAN PRODUKSI
        SQL = "Select a.*,e.nama_produk,b.kode_bahan,c.NamaBarang,b.kode_satuan,d.NamaSatuan,b.qty_bahan,f.Konversi From list_mutasi_produksi_header a "
        SQL = SQL + " inner join list_produksi_hasil b on a.kode_palet=b.noref"
        SQL = SQL + " left join am_itemmst c on b.kode_bahan =c.KodeBarang"
        SQL = SQL + " left join am_unit d on d.KodeSatuan = b.kode_satuan"
        SQL = SQL + " inner join list_produk_master e on a.kode_produk = e.kode_produk"
        SQL = SQL + " left join am_itemdtl f on b.kode_bahan = f.KodeBarang and b.kode_satuan = f.KodeSatuan"
        SQL = SQL + " Where a.kode_palet= '" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Kode Palet pada hasil produksi tidak ditemukan", vbExclamation, AppName
            OBJ.Close
            Exit Sub
        End If
        
        Dtptglscan = RST!tanggal
        lblkode = RST!kode_produk
        lblproduk = RST!nama_produk
        lbllot = RST!nomor_lot
        hapusgrid1
        grid1.Row = 1
        Do While Not RST.EOF
            grid1.Col = 0
            Set grid1.CellPicture = uncheck
            grid1.TextMatrix(grid1.Row, 1) = RST!kode_bahan
            grid1.TextMatrix(grid1.Row, 2) = RST!namabarang
            grid1.TextMatrix(grid1.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00")
            grid1.TextMatrix(grid1.Row, 5) = RST!KODE_SATUAN
            grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
            grid1.TextMatrix(grid1.Row, 8) = RST!konversi
            OBJ1.Open dsn
            SQL1 = "Select * From am_itemkg Where kodebarang = '" & RST!kode_bahan & "'"
            SQL1 = SQL1 + " and kodesatuan= '" & RST!KODE_SATUAN & "' and tahun= '" & Right(RST!tanggal, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                If Mid(RST!tanggal, 4, 2) = "01" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg1
                ElseIf Mid(RST!tanggal, 4, 2) = "02" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg2
                ElseIf Mid(RST!tanggal, 4, 2) = "03" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg3
                ElseIf Mid(RST!tanggal, 4, 2) = "04" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg4
                ElseIf Mid(RST!tanggal, 4, 2) = "05" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg5
                ElseIf Mid(RST!tanggal, 4, 2) = "06" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg6
                ElseIf Mid(RST!tanggal, 4, 2) = "07" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg7
                ElseIf Mid(RST!tanggal, 4, 2) = "08" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg8
                ElseIf Mid(RST!tanggal, 4, 2) = "09" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg9
                ElseIf Mid(RST!tanggal, 4, 2) = "10" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg10
                ElseIf Mid(RST!tanggal, 4, 2) = "11" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg11
                ElseIf Mid(RST!tanggal, 4, 2) = "12" Then
                    grid1.TextMatrix(grid1.Row, 7) = RST1!kg12
                End If
            End If
            OBJ1.Close
            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
            RST.MoveNext
        Loop
        
        'OPEN PEMAKAIAN KEMASAN
        SQL = "Select a.kode_bahan,b.NamaBarang,a.lot_bahan,a.qty_bahan,a.kode_satuan,c.NamaSatuan,a.hpp"
        SQL = SQL + " From list_produksi_kemasan a inner join am_apitemmst b on a.kode_bahan = b.KodeBarang"
        SQL = SQL + " left join am_apunit c on a.kode_satuan = c.KodeSatuan"
        SQL = SQL + " Where a.noref = '" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Kode Palet pada kemasan tidak ditemukan", vbExclamation, AppName
            OBJ.Close
            Exit Sub
        End If
        
        hapusgrid2
        grid2.Row = 1
        Do While Not RST.EOF
            grid2.Col = 0
            Set grid2.CellPicture = uncheck
            grid2.TextMatrix(grid2.Row, 1) = RST!kode_bahan
            grid2.TextMatrix(grid2.Row, 2) = RST!namabarang
            grid2.TextMatrix(grid2.Row, 3) = RST!Lot_bahan
            
            'cek konversi (untuk revisi nilai qty & hpp)
                'OBJ1.Open dsn
                'SQL1 = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                'Set RST1 = OBJ1.Execute(SQL1)
                'If Not RST1.EOF Then
                    'konversi balik
                    'grid2.TextMatrix(grid2.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00") / Format(RST1!nilai, "##,###,##0.00")
                    'grid2.TextMatrix(grid2.Row, 4) = Format(grid2.TextMatrix(grid2.Row, 4), "##,###,##0.00")
                'Else
                    'grid2.TextMatrix(grid2.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00")
                'End If
                'OBJ1.Close
            grid2.TextMatrix(grid2.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00")
            grid2.TextMatrix(grid2.Row, 5) = RST!KODE_SATUAN
            grid2.TextMatrix(grid2.Row, 6) = RST!namasatuan
            grid2.TextMatrix(grid2.Row, 7) = Format(RST!hpp, "##,###,##0.00")
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST.MoveNext
        Loop
        
        
        SQL = "Select * From am_bpblin Where keterangan='" & txtpalet & "' and Type='99'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            lblgudang = "WIP"
        Else
            lblgudang = "Gudang Pusat"
        End If
        OBJ.Close
    End If
End Sub
Private Sub initGrid1()
    With grid1
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Produk"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "Kg"
        .TextMatrix(0, 8) = "isi"
        .ColAlignmentFixed(4) = flexAlignRightCenter
    End With
End Sub

Private Sub setGrid1()
    With grid1
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 4000
        .ColWidth(3) = 0
        .ColWidth(4) = 900
        .ColWidth(5) = 550
        .ColWidth(6) = 1200
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
    End With
End Sub
Private Sub initGrid2()
    With grid2
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Kemasan"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "HPP"
        .ColAlignmentFixed(4) = flexAlignRightCenter
        .ColAlignmentFixed(7) = flexAlignRightCenter
    End With
End Sub

Private Sub setGrid2()
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 4000
        .ColWidth(3) = 0
        .ColWidth(4) = 900
        .ColWidth(5) = 500
        .ColWidth(6) = 1000
        If akses = True Then
            .ColWidth(7) = 1400
        Else
            .ColWidth(7) = 0
        End If
    End With
End Sub

Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
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
    setGrid1
End Sub
Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    setGrid2
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
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = grid1.Rows - 1
    grid1.Col = 0
    Set grid1.CellPicture = blank
End Sub
Private Sub hapusrow2()
    grid2.TextMatrix(grid2.Row, 1) = ""
    grid2.TextMatrix(grid2.Row, 2) = ""
    grid2.TextMatrix(grid2.Row, 3) = ""
    grid2.TextMatrix(grid2.Row, 4) = ""
    grid2.TextMatrix(grid2.Row, 5) = ""
    grid2.TextMatrix(grid2.Row, 6) = ""
    grid2.TextMatrix(grid2.Row, 7) = ""
    Do While True
        If grid2.TextMatrix(grid2.Row + 1, 1) = "" Then
            grid2.TextMatrix(grid2.Row, 1) = ""
            grid2.TextMatrix(grid2.Row, 2) = ""
            grid2.TextMatrix(grid2.Row, 3) = ""
            grid2.TextMatrix(grid2.Row, 4) = ""
            grid2.TextMatrix(grid2.Row, 5) = ""
            grid2.TextMatrix(grid2.Row, 6) = ""
            grid2.TextMatrix(grid2.Row, 7) = ""
            Exit Do
        End If
        grid2.TextMatrix(grid2.Row, 1) = grid2.TextMatrix(grid2.Row + 1, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row + 1, 2)
        grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row + 1, 3)
        grid2.TextMatrix(grid2.Row, 4) = grid2.TextMatrix(grid2.Row + 1, 4)
        grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row + 1, 5)
        grid2.TextMatrix(grid2.Row, 6) = grid2.TextMatrix(grid2.Row + 1, 6)
        grid2.TextMatrix(grid2.Row, 7) = grid2.TextMatrix(grid2.Row + 1, 7)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = blank
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    If KeyAscii = 13 Then
        GetStokBarang Format(Dtptglscan, "yyyyMMdd"), grid2.TextMatrix(grid2.Row, 1), , , stokbahan
        
        'If stokbahan <= 0 Or stokbahan <= txtqty.Value Then
        '    MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
        '    Exit Sub
        'End If

        grid2.TextMatrix(grid2.Row, 4) = txtqty.text
        grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, txtqty.Value), "##,###,###,##0.00")
        grid2.SetFocus
    End If

End Sub

Private Sub txtqty_LostFocus()
    txtqty.Visible = False
End Sub
