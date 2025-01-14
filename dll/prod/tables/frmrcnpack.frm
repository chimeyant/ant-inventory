VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmrcnpack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packaging Plan"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4800
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8467
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
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
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   7005
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
      Left            =   600
      Picture         =   "frmrcnpack.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   6960
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
      Picture         =   "frmrcnpack.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   5520
      TabIndex        =   0
      Top             =   6600
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
   Begin XtremeSuiteControls.PushButton btnpack 
      Height          =   465
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
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
   Begin VB.PictureBox Picgrid 
      BackColor       =   &H00404040&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   6345
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   6405
      Begin VB.TextBox txtinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmrcnpack.frx":0630
         Top             =   60
         Width           =   6210
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   5715
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   6315
      _Version        =   851970
      _ExtentX        =   11139
      _ExtentY        =   10081
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
      Appearance      =   2
      PaintManager.Position=   2
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Produk"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "grid1"
      Item(0).Control(1)=   "txtnilai"
      Item(1).Caption =   "Kemasan"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "grid2"
      Item(1).Control(1)=   "txtqty"
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   465
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmrcnpack.frx":0636
         Caption         =   "frmrcnpack.frx":0656
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmrcnpack.frx":06C2
         Keys            =   "frmrcnpack.frx":06E0
         Spin            =   "frmrcnpack.frx":0722
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8454143
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   4305
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7594
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   8421504
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
         _Band(0).Cols   =   6
      End
      Begin TDBNumber6Ctl.TDBNumber txtqty 
         Height          =   255
         Left            =   -69790
         TabIndex        =   13
         Top             =   450
         Visible         =   0   'False
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmrcnpack.frx":074A
         Caption         =   "frmrcnpack.frx":076A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmrcnpack.frx":07D6
         Keys            =   "frmrcnpack.frx":07F4
         Spin            =   "frmrcnpack.frx":0836
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8454143
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,##0;(##,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,##0;(##,###,##0)"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   4305
         Left            =   -69880
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7594
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   8421504
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
         _Band(0).Cols   =   8
      End
   End
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   465
      Left            =   4560
      TabIndex        =   26
      Top             =   6600
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
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR        RENCANA PRODUKSI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   30
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblkdrcn 
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
      Left            =   1320
      TabIndex        =   29
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perhatian : Pada Mode Edit ini, data packaging plan sebelumnya telah dihapus"
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
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblnamaproduk 
      Caption         =   "Label4"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblkdpack 
      Alignment       =   1  'Right Justify
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
      Left            =   4440
      TabIndex        =   25
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "KODE KEMASAN RENCANA PRODUKSI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   24
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblkdsatuan 
      Caption         =   "Label4"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblisi 
      Caption         =   "Label4"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblkg 
      Caption         =   "Label4"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kg"
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
      Left            =   5640
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblsatuan 
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
      Left            =   5400
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblqty 
      Alignment       =   1  'Right Justify
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
      Left            =   5040
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblkodebrg 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblnamabarang 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblproduk 
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode Produk :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmrcnpack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event doexit(ByVal s_exit As Boolean)

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String
Private package, kaleng, etiket, gridon As Boolean
Dim str2, str3, str4 As String
Dim Totalkg As Double

Private Sub btnClose_Click()
    If hasil4 = "" Then
        Unload Me
    Else
        If MsgBox("Apakah Kode kemasan : " & hasil4 & " akan dihapus ?" & vbLf & "Klik Yes untuk hapus kode Dan klik No untuk lanjut input packaging", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            hasil = "hapus"
            Unload Me
        End If
    End If
End Sub

Private Sub btnpack_Click()
    Call showpack
    gridon = True
    btnpack.Enabled = False
End Sub

Private Sub btnSave_Click()
    If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
    'SIMPAN KE TABEL am_rcnpack
    OBJ.Open dsn
    SQL = "Select * From am_rcnpack Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
   
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!KD_PACK = lblkdpack
        RST!Kd_RCN = lblkdrcn
        RST!KODE_PRODUK = lblproduk
        RST!NAMA_PRODUK = lblnamaproduk
        RST!KODE_BRG = lblkodebrg
        RST!NAMA_BARANG = lblnamabarang
        RST!QTY_ITEM = lblqty
        RST!KODE_SATUAN = lblkdsatuan
        RST!SATUAN = lblsatuan
        RST!KG = lblkg
        RST!ISI = lblisi
        RST!Totalkg = CDbl(lblqty * CDbl(lblkg * lblisi))
        RST!TGL1 = frmrcnprod.dtpfrom
        RST!TGL2 = frmrcnprod.dtpto
        RST!KODE_PACK = grid2.TextMatrix(grid2.Row, 1)
        RST!KEMASAN = grid2.TextMatrix(grid2.Row, 2)
        RST!QTY_PACK = grid2.TextMatrix(grid2.Row, 4)
        RST!FLAG = "0"
        RST!hpp = "0.00"
        RST.Update
        If grid2.Rows = grid2.Row + 1 Then Exit Do
        grid2.Row = grid2.Row + 1
    Loop
    OBJ.Close
    If grid1.TextMatrix(grid1.Row, 1) = "" Then
        hasil = ""
    Else
        hasil = lblkdpack
        lblkdpack = ""
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    initGrid
    setGrid
    package = False: kaleng = False: etiket = False
    str2 = "": str3 = "": str4 = ""
    txtnilai.Visible = False
    txtqty.Visible = False
    If hasil4 <> "" Then
        'EDIT MODE (Data dihapus dulu)
        'hasil4 berisikan data kode pack sebagai parameter jika edit mode dicancel
        Call hapusdata
        lblinfo.Visible = True
    Else
        'Add Mode
        lblkdpack = getkdpack
        lblinfo.Visible = False
    End If
    
End Sub

Sub hapusdata()
    OBJ.Open dsn
    SQL = "DELETE FROM am_rcnpack WHERE KD_PACK='" & hasil4 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub grid_Click()
    Dim stokbahan As Double
    Dim hppbahan As Double
    Dim konvToKg As Double
    If grid.MouseRow = 0 Then Exit Sub
    If grid.Row = 2 And gridon = True Then Exit Sub
    
    Select Case grid.Col
        Case 0, 1, 2, 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then grid.Visible = False: Picgrid.Visible = False: Exit Sub
                If package = True Then
'===================PILIH KEMASAN
                    package = False
                    grid2.Row = 1
                    Do While True
                        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
                        grid2.Rows = grid2.Rows + 1
                        grid2.Row = grid2.Row + 1
                        grid2.Row = grid2.Rows - 1
                    Loop
                    grid2.Col = 0
                    Set grid2.CellPicture = uncheck
                    grid2.TextMatrix(grid2.Row, 1) = grid.TextMatrix(grid.Row, 1)
                    grid2.TextMatrix(grid2.Row, 2) = grid.TextMatrix(grid.Row, 2)
                    grid2.TextMatrix(grid2.Row, 3) = grid.TextMatrix(grid.Row, 4)
                    grid2.TextMatrix(grid2.Row, 5) = grid.TextMatrix(grid.Row, 3)
                    grid2.TextMatrix(grid2.Row, 6) = grid.TextMatrix(grid.Row, 5)
                    If kaleng = True Then
                        'KEMASAN CHILD
                        str4 = str4 * grid.TextMatrix(grid.Row, 4)
                        grid2.TextMatrix(grid2.Row, 4) = str4
                        grid2.Col = 4
                        grid2.CellBackColor = &H80FFFF
        
                        'cek konversi to kg unit (TOTAL QTY KEMASAN)
                        OBJ.Open dsn
                        SQL = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                        Set RST = OBJ.Execute(SQL)
                        If Not RST.EOF Then
                            konvToKg = Format(str4 / RST!nilai, "##,###,##0.00")
                        Else
                            konvToKg = str4
                        End If
                        OBJ.Close
                    Else
                        'KEMASAN HEADER
                        str3 = lblqty * grid.TextMatrix(grid.Row, 4)
                        'cek konversi to kg unit
                        OBJ.Open dsn
                        SQL = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                        Set RST = OBJ.Execute(SQL)
                        If Not RST.EOF Then
                            konvToKg = Format(str3 / RST!nilai, "##,###,##0.00")
                        Else
                            konvToKg = Format(str3, "##,###,##0.00")
                        End If
                        OBJ.Close

                        grid2.Col = 4
                        txtqty.Width = grid2.ColWidth(grid2.Col) - 40
                        txtqty = grid2.TextMatrix(grid2.Row, grid2.Col)
                        txtqty.Left = grid2.Left + grid2.CellLeft
                        txtqty.Top = grid2.Top + grid2.CellTop + 20
                        txtqty.Visible = True
                        txtqty.SetFocus
                    End If
                    TabControl.SelectedItem = 1
                Else
'===================PILIH PRODUK (BARANG JADI)
                    grid1.Row = 1
                    grid1.Col = 0
                    Set grid1.CellPicture = uncheck
                    grid1.TextMatrix(grid1.Row, 1) = grid.TextMatrix(grid.Row, 1)
                    grid1.TextMatrix(grid1.Row, 2) = grid.TextMatrix(grid.Row, 2)
                    grid1.TextMatrix(grid1.Row, 3) = grid.TextMatrix(grid.Row, 4)
                    grid1.TextMatrix(grid1.Row, 5) = grid.TextMatrix(grid.Row, 3)
                    grid1.TextMatrix(grid1.Row, 4) = lblqty
                    grid1.TextMatrix(1, 7) = CDbl(lblqty * CDbl(lblkg * lblisi))
                    TabControl.SelectedItem = 0
                    gridon = False
                End If
                Picgrid.Visible = False
                grid.Visible = False
    End Select
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    Select Case grid1.Col
        Case 2
            If grid1.TextMatrix(1, 1) = "" Then Exit Sub
            If grid1.TextMatrix(1, 4) = "" Then Exit Sub
            Picgrid.Visible = True
            grid.Visible = True
            str2 = grid1.TextMatrix(grid1.Row, 1)
            hapusgrid
            package = True
            
            OBJ.Open dsn
            SQL = "Select a.kode_kemasan,b.NamaBarang,b.kodesatuan,a.lev,a.konversi,a.id,a.id_root From list_konversilevel a"
            SQL = SQL + " inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang"
            SQL = SQL + " Where a.kode_barang_jadi = '" & str2 & "' and a.id_root = ''"
            Set RST = OBJ.Execute(SQL)
            initGrid
            grid.Row = 1
            Do While Not RST.EOF
                grid.TextMatrix(grid.Row, 1) = RST!kode_kemasan
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = RST!Id
                grid.TextMatrix(grid.Row, 4) = RST!konversi
                grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                grid.Col = 0
                Set grid.CellPicture = uncheck
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                RST.MoveNext
            Loop
            OBJ.Close
            txtinfo = "Pilih Kemasan" & vbCrLf & "Produk : " & grid1.TextMatrix(grid1.Row, 2)
        Case 4
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai.Left = grid1.Left + grid1.CellLeft
            txtnilai.Top = grid1.Top + grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    Select Case grid2.Col
        Case 0
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
            End If
        Case 2, 3
            If grid2.TextMatrix(grid2.Row, 4) = "" Then
                MsgBox "qty belum diisi", vbCritical, AppName
                Exit Sub
            End If
            If grid2.TextMatrix(1, 1) = "" Then Exit Sub
            Picgrid.Visible = True
            grid.Visible = True
            str3 = grid2.TextMatrix(grid2.Row, 5)
            str4 = grid2.TextMatrix(grid2.Row, 4)
            txtinfo = "Pilih konversi kemasan" & vbCrLf & "Packaging : " & grid2.TextMatrix(grid2.Row, 2)
            
            hapusgrid
            package = True: kaleng = True
            OBJ.Open dsn
            SQL = "Select a.kode_kemasan,b.NamaBarang,b.kodesatuan,a.lev,a.konversi,a.id,a.id_root From list_konversilevel a"
            SQL = SQL + " inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang"
            SQL = SQL + " Where a.kode_barang_jadi = '" & str2 & "' and a.id_root = '" & str3 & "'"
            Set RST = OBJ.Execute(SQL)
            
            initGrid
            grid.Row = 1
            Do While Not RST.EOF
                grid.TextMatrix(grid.Row, 1) = RST!kode_kemasan
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = RST!Id
                grid.TextMatrix(grid.Row, 4) = RST!konversi
                grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                grid.Col = 0
                Set grid.CellPicture = uncheck
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                RST.MoveNext
            Loop
            OBJ.Close
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

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    txtnilai.Visible = False
    txtqty.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case grid1.Col
            Case 4:
                txtnilai.Visible = False
                grid1.TextMatrix(grid1.Row, 4) = txtnilai
                etiket = True
                str3 = grid1.TextMatrix(grid1.Row, 4)
        End Select
        
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub showpack()
    grid.Visible = True
    txtnilai.Visible = False
    txtqty.Visible = False
    txtinfo = "Pilih barang jadi"
    
    OBJ.Open dsn
    SQL = "select distinct a.kodebarang, a.namabarang,a.KodeSatuan,c.namasatuan"
    SQL = SQL + " from am_itemdtl a inner join list_produk_hasil b on a.kodebarang=b.kode_barang_jadi"
    SQL = SQL + " inner join am_unit c on a.KodeSatuan = c.KodeSatuan"
    SQL = SQL + " and b.kode_produk='" & lblproduk & "' and a.kodebarang='" & lblkodebrg & "'"
    SQL = SQL + " and a.kodesatuan='" & lblkdsatuan & "' order by a.KodeBarang asc"
    Set RST = OBJ.Execute(SQL)
        
    hapusgrid
    Picgrid.Visible = True
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = RST!kodesatuan
        grid.TextMatrix(grid.Row, 4) = RST!namasatuan
        grid.Col = 0
        Set grid.CellPicture = uncheck
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
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
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
        grid1.Col = 0
        Set grid1.CellPicture = blank
        If grid1.Row = 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    txtnilai.Value = ""
    setGrid
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
        grid2.Col = 0
        Set grid2.CellPicture = blank
        If grid2.Row = 1 Then Exit Do
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    setGrid
End Sub
Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 900
        .ColWidth(2) = 3500
        .ColWidth(3) = 0
        .ColWidth(4) = 900
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
    With grid1
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 3500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 4500
        .ColWidth(3) = 0
        .ColWidth(4) = 1100
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
End Sub

Private Sub initGrid()
    If kaleng = True Or package = True Then
        With grid
            .Cols = 7
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "KODE"
            .TextMatrix(0, 2) = "NAMA BARANG"
            .TextMatrix(0, 3) = "KODE"
            .TextMatrix(0, 4) = "KONVERSI"
        End With
    Else
        With grid
            .Cols = 5
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "KODE"
            .TextMatrix(0, 2) = "NAMA BARANG"
            .TextMatrix(0, 3) = "KODE"
            .TextMatrix(0, 4) = "SATUAN"
        End With
    End If
    With grid1
        .Cols = 8 '6
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "SATUAN"
        .TextMatrix(0, 4) = "QTY"
        .Col = 4
        .CellBackColor = &H80FFFF
        '
        .TextMatrix(0, 5) = "KdSat"
        .TextMatrix(0, 6) = "kg"
        .TextMatrix(0, 7) = "hpp"
    End With
    With grid2
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "SATUAN"
        .TextMatrix(0, 4) = "QTY"
        .Col = 4
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 5) = "ID"
    End With
End Sub
Private Sub hapusrow()
    If grid2.Rows = 2 Then
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Rows = grid2.Rows - 1
        grid2.Col = 0
        Set grid2.CellPicture = blank
        If grid2.Rows = 1 Then grid2.Rows = grid2.Rows + 1: package = False: str4 = grid1.TextMatrix(grid1.Row, 4)
        Exit Sub
    Else
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
    End If
    Do While True
        If grid2.Row = grid2.Rows - 1 Then Exit Do
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
    Set grid2.CellPicture = uncheck
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    If KeyAscii = 13 Then
        Select Case grid2.Col
            Case 4:
                If IsNull(txtqty.Value) Or txtqty.Value = 0 Then
                    MsgBox "Qty tidak boleh kosong", vbCritical, AppName
                    Exit Sub
                End If
                txtqty.Visible = False
                grid2.TextMatrix(grid2.Row, 4) = txtqty
        End Select
    End If
End Sub

Private Sub txtqty_LostFocus()
    txtqty.Visible = False
End Sub

Function getkdpack() As String    '211007001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yymmdd")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(kd_pack)as pack from am_rcnpack where kd_pack like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!pack) = True Or RST!pack = "" Then
        getkdpack = strnumber + "001"
    Else
        kode = CLng(Mid(RST!pack, 7, 3)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getkdpack = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getkdpack = strnumber + "0001"
End Function
