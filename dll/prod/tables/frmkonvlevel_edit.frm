VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmkonvlevel_edit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubah Konversi Kemasan"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   6465
      ScaleHeight     =   1005
      ScaleWidth      =   5895
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   5925
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   4950
         Picture         =   "frmkonvlevel_edit.frx":0000
         Top             =   -15
         Width           =   1110
      End
      Begin VB.Label lblbaris 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "akan dihapus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Top             =   375
         Width           =   4995
      End
      Begin VB.Label lblwarn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar Packaging"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   29
         Top             =   135
         Width           =   4995
      End
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
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   165
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
      Left            =   2355
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   165
      Width           =   3555
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   3210
      Width           =   5745
      _Version        =   851970
      _ExtentX        =   10134
      _ExtentY        =   4260
      _StockProps     =   79
      Caption         =   "Konversi"
      ForeColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtvar 
         Height          =   330
         Left            =   3825
         TabIndex        =   27
         Top             =   1485
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtidroot 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1170
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtkdkemasan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtkemasan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   690
         Width           =   3000
      End
      Begin VB.ComboBox cmblevel 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         TabIndex        =   1
         Top             =   1095
         Width           =   2040
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   300
         Left            =   1380
         TabIndex        =   5
         Top             =   1530
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         Calculator      =   "frmkonvlevel_edit.frx":43A2
         Caption         =   "frmkonvlevel_edit.frx":43C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmkonvlevel_edit.frx":442E
         Keys            =   "frmkonvlevel_edit.frx":444C
         Spin            =   "frmkonvlevel_edit.frx":448E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
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
      Begin XtremeSuiteControls.PushButton cmdkemasan 
         Height          =   240
         Left            =   210
         TabIndex        =   6
         ToolTipText     =   "Click to search kemasan"
         Top             =   720
         Width           =   990
         _Version        =   851970
         _ExtentX        =   1746
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Kemasan"
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
         TextAlignment   =   0
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdRoot 
         Height          =   240
         Left            =   210
         TabIndex        =   7
         ToolTipText     =   "Click to search ID Root"
         Top             =   1980
         Width           =   990
         _Version        =   851970
         _ExtentX        =   1746
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "ID Root :"
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
         TextAlignment   =   0
         Appearance      =   6
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   1935
         Width           =   1065
      End
      Begin VB.Shape Shape2 
         Height          =   330
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1365
         TabIndex        =   14
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label lblRoot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2505
         TabIndex        =   13
         Top             =   1950
         Width           =   3000
      End
      Begin VB.Label lblbarang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2505
         TabIndex        =   12
         Top             =   315
         Width           =   3000
      End
      Begin VB.Label lblkode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1380
         TabIndex        =   11
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Konversi"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   255
         TabIndex        =   10
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   285
         TabIndex        =   9
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   270
         TabIndex        =   8
         Top             =   1155
         Width           =   1035
      End
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   420
      Left            =   11235
      TabIndex        =   17
      Top             =   5460
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
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
   Begin XtremeSuiteControls.PushButton cmdproduksi 
      Height          =   240
      Left            =   210
      TabIndex        =   18
      ToolTipText     =   "Click to search produk"
      Top             =   195
      Width           =   870
      _Version        =   851970
      _ExtentX        =   1535
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
      Height          =   2595
      Left            =   150
      TabIndex        =   19
      ToolTipText     =   "Pilih item untuk konversi kemasan"
      Top             =   585
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   -2147483628
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   4050
      Left            =   6330
      TabIndex        =   20
      ToolTipText     =   "Click here to update"
      Top             =   1320
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   7144
      _Version        =   393216
      BackColor       =   -2147483628
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   420
      Left            =   9915
      TabIndex        =   21
      Top             =   5460
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   555
      Left            =   6510
      TabIndex        =   22
      Top             =   240
      Width           =   5820
      _Version        =   851970
      _ExtentX        =   10266
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "UPDATE AND DELETE PACKAGING"
      ForeColor       =   8421504
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnClear 
      Height          =   420
      Left            =   8595
      TabIndex        =   23
      Top             =   5460
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
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
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnDel 
      Height          =   420
      Left            =   7275
      TabIndex        =   26
      Top             =   5460
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Delete"
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
   Begin VB.Shape Shape3 
      Height          =   330
      Left            =   165
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H00808080&
      Height          =   1035
      Left            =   6345
      Shape           =   4  'Rounded Rectangle
      Top             =   165
      Width           =   6135
   End
   Begin VB.Label LBLBRG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6405
      TabIndex        =   25
      Top             =   855
      Width           =   5985
   End
   Begin VB.Label lblket 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   225
      TabIndex        =   24
      Top             =   5685
      Width           =   6960
   End
End
Attribute VB_Name = "frmkonvlevel_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset

Private SQL As String
Dim baris As String

Private Sub btnClear_Click()
    txtproduk = ""
    txtkodeproduk = ""
    lblkode = ""
    lblbarang = ""
    txtkdkemasan = ""
    txtkemasan = ""
    cmblevel = ""
    txtnilai = "0"
    lblID = ""
    lblRoot = ""
    txtidroot = ""
    LBLBRG = ""
    hapusgrid
    hapusgrid2
End Sub

Private Sub btnClose_Click()
    Unload Me
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
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1000
    End With
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 950
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 800
        .ColWidth(5) = 400
        .ColWidth(6) = 500
    End With
End Sub

Private Sub initGrid()
    With grid
        .Cols = 5
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "KODE"
        .TextMatrix(0, 4) = "SATUAN"
    End With
    With grid2
        .Cols = 7
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "KONVERSI"
        .TextMatrix(0, 5) = "ID"
        .TextMatrix(0, 6) = "ROOT"
    End With
End Sub

Private Sub btnDel_Click()
    If txtkodeproduk = "" Then Exit Sub
    If lblkode = "" Then
        MsgBox "Data tidak lengkap", vbExclamation, AppName
        Exit Sub
    End If
    If grid2.TextMatrix(1, 1) = "" Then
        MsgBox "Data tidak ditemukan", vbInformation, AppName
        Exit Sub
    End If
    lblwarn = "Daftar Packaging " & LBLBRG
    lblbaris = baris & " Item akan dihapus !"
    DoEvents
    Picture1.Visible = True
    If MsgBox("Data konversi untuk kemasan " & LBLBRG & " akan dihapus" & vbCrLf _
    & "Klik OK untuk melanjutkan, Klik Cancel untuk membatalkan.", vbOKCancel + vbQuestion, "KONFIRMASI PENGHAPUSAN DATA") = vbCancel Then Picture1.Visible = False: Exit Sub
    
    OBJ.Open dsn
    SQL = "Delete From list_konversilevel Where kode_produk = '" & txtkodeproduk & "'"
    SQL = SQL + " and kode_barang_jadi='" & lblkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Konversi berhasil dihapus.", vbInformation, AppName
    lblkode = "": lblbarang = ""
    txtkdkemasan = "": txtkemasan = ""
    cmblevel = ""
    txtnilai = "0"
    lblID = "": lblRoot = "": txtidroot = ""
    Picture1.Visible = False
    LoadDataGrid2
End Sub

Private Sub btnSave_Click()
    If txtkodeproduk = "" Then Exit Sub
    If lblkode = "" Or txtkdkemasan = "" Then
        MsgBox "Data tidak lengkap", vbExclamation, AppName
        Exit Sub
    End If
    
    If MsgBox("Data konversi akan diupdate " & vbCrLf _
    & "Klik OK untuk melanjutkan, Klik Cancel untuk membatalkan.", vbOKCancel + vbQuestion, "KONFIRMASI PENGHAPUSAN DATA") = vbCancel Then Exit Sub
    
    OBJ.Open dsn
    SQL = "Select * From list_konversilevel Where kode_produk='" & txtkodeproduk & "'"
    SQL = SQL + " and kode_barang_jadi='" & lblkode & "' and kode_kemasan='" & txtvar & "'"
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    With RST
        !kode_kemasan = txtkdkemasan
        !lev = cmblevel
        !id_root = txtidroot
        !konversi = txtnilai
        .Update
    End With
    OBJ.Close
    MsgBox "Data konversi berhasil diupdated", vbInformation, AppName
    lblkode = "": lblbarang = ""
    txtkdkemasan = "": txtkemasan = ""
    cmblevel = ""
    txtnilai = "0"
    lblID = "": lblRoot = "": txtidroot = ""
    txtvar = ""
    LoadDataGrid2

End Sub

Private Sub cmdkemasan_Click()
    If lblkode = "" Then Exit Sub
    If txtkdkemasan = "" Then
        MsgBox "Silahkan klik tabel/grid kemasan terlebih dahulu", vbExclamation, "Update Procedure"
        Exit Sub
    End If
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkemasan_GotFocus()
    If hasil = "" Then Exit Sub
    txtkdkemasan = hasil
    txtkemasan = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdproduksi_Click()
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Anda tidak memiliki akses..! ", vbCritical, AppName
        Exit Sub
    End If
proses:
    namatabel = "Produk."
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
    OBJ.Open dsn
    SQL = "select a.kodebarang,a.namabarang,b.kode_satuan,c.namasatuan "
    SQL = SQL + "from am_itemmst a inner join list_produk_hasil b "
    SQL = SQL + "on a.kodebarang=b.kode_barang_jadi inner join am_unit c "
    SQL = SQL + "on b.kode_satuan=c.kodesatuan "
    SQL = SQL + "and b.kode_produk='" & txtkodeproduk & "' "
    Set RST = OBJ.Execute(SQL)
    hapusgrid
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = RST!KODE_SATUAN
        grid.TextMatrix(grid.Row, 4) = RST!namasatuan
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdRoot_Click()
    If txtkdkemasan = "" Then Exit Sub
    namatabel = "konversilevel"
    carisql1 = "Select a.kode_kemasan,b.NamaBarang,a.id From list_konversilevel a "
    carisql1 = carisql1 + "inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang "
    carisql1 = carisql1 + "Where a.kode_barang_jadi='" & lblkode & "'"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdRoot_GotFocus()
    If hasil = "" Then Exit Sub
    lblID = hasil
    lblRoot = hasil1
    txtidroot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub Form_Load()
    cmblevel.AddItem "Header"
    cmblevel.AddItem "Child"
    initGrid
    setGrid
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    lblkode = grid.TextMatrix(grid.Row, 1)
    lblbarang = grid.TextMatrix(grid.Row, 2)
    LBLBRG = grid.TextMatrix(grid.Row, 2)
    LoadDataGrid2
End Sub

Private Sub LoadDataGrid2()
    hapusgrid2
    OBJ.Open dsn
    SQL = "Select a.kode_kemasan,b.NamaBarang,a.lev,a.konversi,a.id,a.id_root From list_konversilevel a "
    SQL = SQL + "inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang "
    SQL = SQL + "Where a.kode_barang_jadi = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)

    grid2.Row = 1
    Do While Not RST.EOF
        grid2.TextMatrix(grid2.Row, 0) = grid2.Row
        grid2.TextMatrix(grid2.Row, 1) = RST!kode_kemasan
        grid2.TextMatrix(grid2.Row, 2) = RST!namabarang
        grid2.TextMatrix(grid2.Row, 3) = RST!lev
        grid2.TextMatrix(grid2.Row, 4) = RST!konversi
        grid2.TextMatrix(grid2.Row, 5) = RST!Id
        grid2.TextMatrix(grid2.Row, 6) = RST!id_root
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        RST.MoveNext
    Loop
    baris = grid2.Row - 1
    OBJ.Close
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
        grid2.Col = 0
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    setGrid
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    Dim strid As String
    txtkdkemasan = grid2.TextMatrix(grid2.Row, 1)
    txtvar = grid2.TextMatrix(grid2.Row, 1)
    txtkemasan = grid2.TextMatrix(grid2.Row, 2)
    cmblevel = grid2.TextMatrix(grid2.Row, 3)
    txtnilai = grid2.TextMatrix(grid2.Row, 4)
    strid = grid2.TextMatrix(grid2.Row, 5)
    txtidroot = grid2.TextMatrix(grid2.Row, 6)

    If txtidroot = "" Or IsNull(txtidroot) Then txtidroot = strid
    
    OBJ.Open dsn
    SQL = "Select a.kode_kemasan,b.namabarang From list_konversilevel a "
    SQL = SQL + "inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang "
    SQL = SQL + "Where a.id='" & txtidroot & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then lblID = RST!kode_kemasan: lblRoot = RST!namabarang
    OBJ.Close
End Sub

Private Sub lblRoot_Change()
    lblket = "1 " & lblRoot & " = " & txtnilai & " " & txtkemasan
End Sub

Private Sub txtkemasan_Change()
    lblket = "1 " & lblRoot & " = " & txtnilai & " " & txtkemasan
End Sub

Private Sub txtnilai_Change()
    lblket = "1 " & lblRoot & " = " & txtnilai & " " & txtkemasan
End Sub
