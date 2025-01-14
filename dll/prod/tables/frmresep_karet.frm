VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmresep_karet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ATUR FORMULA KARET"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmresep_karet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnmprod 
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
      Height          =   345
      Left            =   1380
      TabIndex        =   3
      Top             =   555
      Width           =   5580
   End
   Begin VB.TextBox txtKdProduk 
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
      Height          =   345
      Left            =   1380
      TabIndex        =   2
      Top             =   195
      Width           =   2010
   End
   Begin VB.ComboBox cmbklasifikasi 
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
      TabIndex        =   1
      Top             =   930
      Width           =   2040
   End
   Begin VB.TextBox txtnoproduk 
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
      Height          =   345
      Left            =   1380
      TabIndex        =   0
      Top             =   1290
      Width           =   2010
   End
   Begin XtremeSuiteControls.PushButton btnKodeProduk 
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   1230
      _Version        =   851970
      _ExtentX        =   2170
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Kode Resep    : "
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   0
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   1320
      Width           =   1230
      _Version        =   851970
      _ExtentX        =   2170
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nomor Produk : "
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   0
      Appearance      =   5
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4635
      Left            =   60
      TabIndex        =   8
      Top             =   1860
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
      Item(0).Caption =   "BAHAN BAKU"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "TabControlPage1"
      Item(0).Control(1)=   "page1"
      Item(1).Caption =   "BARANG JADI"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4275
         Left            =   -69970
         TabIndex        =   9
         Top             =   330
         Visible         =   0   'False
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   2
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
            Height          =   4035
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   7117
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
         Left            =   30
         TabIndex        =   11
         Top             =   330
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   1
         Begin VB.TextBox txtinisial 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   8235
            TabIndex        =   13
            Top             =   135
            Visible         =   0   'False
            Width           =   2010
         End
         Begin TDBNumber6Ctl.TDBNumber txtnilai 
            Height          =   255
            Left            =   7605
            TabIndex        =   12
            Top             =   -720
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            Calculator      =   "frmresep_karet.frx":000C
            Caption         =   "frmresep_karet.frx":002C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmresep_karet.frx":0098
            Keys            =   "frmresep_karet.frx":00B6
            Spin            =   "frmresep_karet.frx":00F8
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   4035
            Left            =   60
            TabIndex        =   14
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
         Left            =   30
         TabIndex        =   15
         Top             =   330
         Width           =   10350
         _Version        =   851970
         _ExtentX        =   18256
         _ExtentY        =   7541
         _StockProps     =   1
         Page            =   0
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpresep 
      Height          =   315
      Left            =   8955
      TabIndex        =   16
      Top             =   195
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
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   9420
      TabIndex        =   18
      Top             =   6555
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Resep  :"
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
      Left            =   7620
      TabIndex        =   17
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   150
      TabIndex        =   7
      Top             =   630
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   180
      TabIndex        =   6
      Top             =   1005
      Width           =   1110
   End
End
Attribute VB_Name = "frmresep_karet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbklasifikasi.AddItem "SIMPLEX"
    cmbklasifikasi.AddItem "VIBER"
    cmbklasifikasi.AddItem "KARPET"
    dtpresep = Date
End Sub

