VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmrcnprodedit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Rencana Produksi"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   11280
      TabIndex        =   0
      Top             =   6120
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
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   465
      Left            =   10320
      TabIndex        =   1
      Top             =   6120
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
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   5355
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   11955
      _Version        =   851970
      _ExtentX        =   21087
      _ExtentY        =   9446
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
      ItemCount       =   3
      Item(0).Caption =   "Produk"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "txtnilai"
      Item(0).Control(1)=   "grid"
      Item(1).Caption =   "Bahan Baku"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "MSHFlexGrid1"
      Item(1).Control(1)=   "lblrcnbb"
      Item(2).Caption =   "Kemasan"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "grid2"
      Item(2).Control(1)=   "txtqty"
      Item(2).Control(2)=   "lblrcnpack"
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   945
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmrcnprodedit.frx":0000
         Caption         =   "frmrcnprodedit.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmrcnprodedit.frx":008C
         Keys            =   "frmrcnprodedit.frx":00AA
         Spin            =   "frmrcnprodedit.frx":00EC
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   4305
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   11745
         _ExtentX        =   20717
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
         TabIndex        =   5
         Top             =   930
         Visible         =   0   'False
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmrcnprodedit.frx":0114
         Caption         =   "frmrcnprodedit.frx":0134
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmrcnprodedit.frx":01A0
         Keys            =   "frmrcnprodedit.frx":01BE
         Spin            =   "frmrcnprodedit.frx":0200
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
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   11745
         _ExtentX        =   20717
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4305
         Left            =   -69880
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   11745
         _ExtentX        =   20717
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
      Begin VB.Label lblrcnbb 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NoRcnBb"
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
         Left            =   -59800
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblrcnpack 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NoRcnPack"
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
         Left            =   -59800
         TabIndex        =   10
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.PushButton cmdrcn 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Rcn"
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
   Begin XtremeSuiteControls.DateTimePicker dtpfrom 
      Height          =   315
      Left            =   9015
      TabIndex        =   12
      Top             =   240
      Width           =   1275
      _Version        =   851970
      _ExtentX        =   2249
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
      Left            =   10935
      TabIndex        =   13
      Top             =   240
      Width           =   1275
      _Version        =   851970
      _ExtentX        =   2249
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
      Left            =   8400
      TabIndex        =   15
      Top             =   240
      Width           =   555
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
      Left            =   10560
      TabIndex        =   14
      Top             =   255
      Width           =   555
   End
   Begin VB.Label lblnorcn 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmrcnprodedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmdrcn_Click()
    carisql1 = "Select KD_RCN,TGL1,TGL2,SUM(TOTALKG)'Kg' From am_rcnprod"
    namatabel = "Rencana Produksi"
    frmsearch.Show vbModal
End Sub

Private Sub cmdrcn_GotFocus()
    lblnorcn = hasil
    dtpfrom = hasil1
    dtpto = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hasil3 = ""
    Call viewdata
End Sub

Private Sub viewdata()
    
End Sub
