VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmFArepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repair Data Fixed Asset"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComCtl2.DTPicker datetrx 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   134742019
      CurrentDate     =   38767
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5295
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   12975
      _Version        =   851970
      _ExtentX        =   22886
      _ExtentY        =   9340
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Color           =   16
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Jurnal Penjualan ( JJ )"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "grid"
      Item(0).Control(1)=   "cmdsavejj"
      Item(0).Control(2)=   "date1"
      Item(0).Control(3)=   "txtstring"
      Item(0).Control(4)=   "cmbDK"
      Item(0).Control(5)=   "txtnilai"
      Item(1).Caption =   "Bank Masuk ( BM )"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "grid1"
      Item(1).Control(1)=   "cmdAddBM"
      Item(1).Control(2)=   "date2"
      Item(1).Control(3)=   "txtnotrx"
      Item(1).Control(4)=   "lblinfo"
      Item(1).Control(5)=   "txtdesc"
      Item(1).Control(6)=   "cmbDK2"
      Item(1).Control(7)=   "txtamount"
      Item(1).Control(8)=   "Label5"
      Item(2).Caption =   "Jurnal Penyusutan ( JS )"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "grid2"
      Item(2).Control(1)=   "Label6"
      Item(2).Control(2)=   "txtsusut"
      Item(2).Control(3)=   "Label2"
      Item(2).Control(4)=   "txtsisa"
      Item(2).Control(5)=   "cmdprosesjs"
      Item(2).Control(6)=   "cmbflag"
      Item(2).Control(7)=   "txtjumlah"
      Begin TDBNumber6Ctl.TDBNumber txtamount 
         Height          =   255
         Left            =   -60760
         TabIndex        =   38
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmFArepair.frx":0000
         Caption         =   "frmFArepair.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFArepair.frx":008C
         Keys            =   "frmFArepair.frx":00AA
         Spin            =   "frmFArepair.frx":00EC
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.ComboBox cmbDK2 
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
         ItemData        =   "frmFArepair.frx":0114
         Left            =   -61720
         List            =   "frmFArepair.frx":0116
         TabIndex        =   37
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -63280
         TabIndex        =   36
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbflag 
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
         ItemData        =   "frmFArepair.frx":0118
         Left            =   7440
         List            =   "frmFArepair.frx":011A
         TabIndex        =   34
         Top             =   4800
         Width           =   855
      End
      Begin TDBNumber6Ctl.TDBNumber txtsusut 
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   4920
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calculator      =   "frmFArepair.frx":011C
         Caption         =   "frmFArepair.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFArepair.frx":01A8
         Keys            =   "frmFArepair.frx":01C6
         Spin            =   "frmFArepair.frx":0208
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.TextBox txtnotrx 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -64840
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   255
         Left            =   -66040
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   134742019
         CurrentDate     =   38767
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   -66040
         TabIndex        =   22
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmFArepair.frx":0230
         Caption         =   "frmFArepair.frx":0250
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFArepair.frx":02BC
         Keys            =   "frmFArepair.frx":02DA
         Spin            =   "frmFArepair.frx":031C
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.ComboBox cmbDK 
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
         ItemData        =   "frmFArepair.frx":0344
         Left            =   -67000
         List            =   "frmFArepair.frx":0346
         TabIndex        =   21
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtstring 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -68560
         TabIndex        =   20
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   255
         Left            =   -69760
         TabIndex        =   19
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   134742019
         CurrentDate     =   38767
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   4095
         Left            =   -69880
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   22
         FixedCols       =   0
         BackColorBkg    =   -2147483632
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   22
      End
      Begin XtremeSuiteControls.PushButton cmdsavejj 
         Height          =   375
         Left            =   -58120
         TabIndex        =   14
         Top             =   4800
         Visible         =   0   'False
         Width           =   945
         _Version        =   851970
         _ExtentX        =   1667
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Save JJ"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   4095
         Left            =   -69880
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         BackColorBkg    =   -2147483632
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   21
      End
      Begin XtremeSuiteControls.PushButton cmdAddBM 
         Height          =   375
         Left            =   -58120
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   945
         _Version        =   851970
         _ExtentX        =   1667
         _ExtentY        =   661
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
         Appearance      =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   4095
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         BackColorBkg    =   -2147483632
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   21
      End
      Begin TDBNumber6Ctl.TDBNumber txtsisa 
         Height          =   255
         Left            =   4800
         TabIndex        =   30
         Top             =   4920
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calculator      =   "frmFArepair.frx":0348
         Caption         =   "frmFArepair.frx":0368
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFArepair.frx":03D4
         Keys            =   "frmFArepair.frx":03F2
         Spin            =   "frmFArepair.frx":0434
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin XtremeSuiteControls.PushButton cmdprosesjs 
         Height          =   375
         Left            =   9480
         TabIndex        =   32
         Top             =   4800
         Width           =   3345
         _Version        =   851970
         _ExtentX        =   5900
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Proses Jurnal Penyusutan"
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
      Begin TDBNumber6Ctl.TDBNumber txtjumlah 
         Height          =   255
         Left            =   7920
         TabIndex        =   41
         Top             =   4920
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmFArepair.frx":045C
         Caption         =   "frmFArepair.frx":047C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFArepair.frx":04E8
         Keys            =   "frmFArepair.frx":0506
         Spin            =   "frmFArepair.frx":0548
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "* Wajib menyertakan kode Aktiva pada kolom keterangan"
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
         Height          =   255
         Left            =   -69760
         TabIndex        =   39
         Top             =   4800
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblinfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "(manual) BM=Bank Masuk (YYMM/zz/XXXXX)"
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
         Height          =   255
         Left            =   -69760
         TabIndex        =   35
         Top             =   5040
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Akumulasi Penyusutan"
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
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa Buku"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   4920
         Width           =   855
      End
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
      Left            =   10080
      Picture         =   "frmFArepair.frx":0570
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
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
      Left            =   9840
      Picture         =   "frmFArepair.frx":0926
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txthjual 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txthbeli 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtkdaktiva 
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
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   12000
      TabIndex        =   0
      Top             =   6840
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdkdaktiva 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1425
      _Version        =   851970
      _ExtentX        =   2514
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   "Kode Aktiva"
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
   Begin MSComCtl2.DTPicker datebeli 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   134742019
      CurrentDate     =   38767
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   12975
      _Version        =   851970
      _ExtentX        =   22886
      _ExtentY        =   450
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin VB.Label lblnoBM 
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   6840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa Buku dalam sistem"
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
      Left            =   9000
      TabIndex        =   31
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblsisa 
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
      Height          =   255
      Left            =   10920
      TabIndex        =   29
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblsusut 
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
      Height          =   255
      Left            =   10920
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Tanggal/Harga Jual"
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
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal/Harga Beli"
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
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblaktiva 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5595
   End
End
Attribute VB_Name = "frmFArepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim posrow As String

Private Sub cmbDK_Click()
    grid.Row = posrow
    
    grid.SetFocus
    grid.TextMatrix(grid.Row, 7) = cmbDK
    cmbDK.Visible = False
End Sub

Private Sub cmbDK_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmbDK.Visible = False
End Sub

Private Sub cmbDK_LostFocus()
    cmbDK.Visible = False
End Sub

Private Sub cmbDK2_Click()
    grid1.Row = posrow
    
    grid1.SetFocus
    grid1.TextMatrix(grid1.Row, 7) = cmbDK2
    cmbDK2.Visible = False
End Sub

Private Sub cmbDK2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmbDK2.Visible = False
End Sub

Private Sub cmbDK2_LostFocus()
    cmbDK2.Visible = False
End Sub

Private Sub cmbflag_Click()
    grid2.Row = posrow
    
    grid2.SetFocus
    grid2.TextMatrix(grid2.Row, 11) = cmbflag
    cmbflag.Visible = False
End Sub

Private Sub cmbflag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmbflag = False
End Sub

Private Sub cmbflag_LostFocus()
    cmbflag.Visible = False
End Sub

Private Sub cmdAddBM_Click()
    OBJ.Open dsn
    If cmdAddBM.Caption = "Save" Then
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
        grid1.Row = 1
        Do While True
            With grid1
                If .TextMatrix(.Row, 1) = "" Then Exit Do
                SQL = "insert into gl_transaksi ("
                SQL = SQL + "kdcomp,tgltrx,kdtrx,notrx,kurs,noactrx,desctrx,dbkrtrx,amounttrx,nilaitrx,currtrx,flag,flagprint,flagadjustment,lineitem,identry,idupdate,dateentry,dateupdate,cekbg,reconsil) VALUES('01',"
                SQL = SQL + "convert(datetime, '" + Format(date2, "MM/dd/yyyy") + "'),'BM','"
                SQL = SQL + grid1.TextMatrix(.Row, 3) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 4) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 5) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 6) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 7) + "',"
                SQL = SQL + "convert(money,'" + grid1.TextMatrix(.Row, 8) + "'),"
                SQL = SQL + "convert(money,'" + grid1.TextMatrix(.Row, 9) + "'),'"
                SQL = SQL + grid1.TextMatrix(.Row, 10) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 11) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 12) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 13) + "',"
                SQL = SQL + "convert(numeric,'" & grid1.Row & "'),"
                SQL = SQL + "'','',"
                SQL = SQL + "convert(datetime, '" + Format(grid1.TextMatrix(grid1.Row, 17), "MM/dd/yyyy") + "'),"
                SQL = SQL + "convert(datetime, '" + Format(Now, "MM/dd/yyyy") + "'),'"
                SQL = SQL + grid1.TextMatrix(.Row, 19) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 20) + "')"
                OBJ.Execute (SQL)
                .Row = .Row + 1
            End With
        Loop
    ElseIf cmdAddBM.Caption = "Update" Then
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
        
        SQL = "Delete From gl_transaksi Where kdtrx = 'BM' and notrx = '" & lblnoBM & "'"
        Set RST = OBJ.Execute(SQL)
        
        grid1.Row = 1
        Do While True
            With grid1
                If .TextMatrix(.Row, 1) = "" Then Exit Do
                SQL = "insert into gl_transaksi ("
                SQL = SQL + "kdcomp,tgltrx,kdtrx,notrx,kurs,noactrx,desctrx,dbkrtrx,amounttrx,nilaitrx,currtrx,flag,flagprint,flagadjustment,lineitem,identry,idupdate,dateentry,dateupdate,cekbg,reconsil) VALUES('01',"
                SQL = SQL + "convert(datetime, '" + Format(date2, "MM/dd/yyyy") + "'),'BM','"
                SQL = SQL + grid1.TextMatrix(.Row, 3) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 4) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 5) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 6) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 7) + "',"
                SQL = SQL + "convert(money,'" + grid1.TextMatrix(.Row, 8) + "'),"
                SQL = SQL + "convert(money,'" + grid1.TextMatrix(.Row, 9) + "'),'"
                SQL = SQL + grid1.TextMatrix(.Row, 10) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 11) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 12) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 13) + "',"
                SQL = SQL + "convert(numeric,'" & grid1.Row & "'),"
                SQL = SQL + "'','',"
                SQL = SQL + "convert(datetime, '" + Format(grid1.TextMatrix(grid1.Row, 17), "MM/dd/yyyy") + "'),"
                SQL = SQL + "convert(datetime, '" + Format(Now, "MM/dd/yyyy") + "'),'"
                SQL = SQL + grid1.TextMatrix(.Row, 19) + "','"
                SQL = SQL + grid1.TextMatrix(.Row, 20) + "')"
                OBJ.Execute (SQL)
                .Row = .Row + 1
            End With
        Loop
    End If
    OBJ.Close
    MsgBox "Data is successfully saved", vbInformation, AppName
    hapusgrid
    hapusgrid2
    Clearform
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdkdaktiva_Click()
    carisql1 = "Select distinct notrx,tgltrx From gl_transaksi Where kdtrx = 'JJ'"
    namatabel = "FA"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkdaktiva_GotFocus()
    MsgBox hasil
    If hasil = "" Then Exit Sub
    txtkdaktiva = hasil
    hasil = ""
    hasil1 = ""
    hapusgrid
    hapusgrid2
    Me.MousePointer = vbHourglass
    showFA
    Me.MousePointer = vbDefault
End Sub

Private Sub showFA()
    Dim j As Integer
    If txtkdaktiva = "" Then Exit Sub

    OBJ.Open dsn
    SQL = "Select * From gl_aktiva Where kdaktiva= '" & txtkdaktiva & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Aktiva tidak ditemukan", vbExclamation, AppName
    Else
        lblaktiva = RST!nmaktiva
        txthbeli = Format(RST!hargabeli, "#,##0.00")
        txthjual = Format(RST!hargajual, "#,##0.00")
        datebeli = RST!tglbeli
        datetrx = RST!tgljual
        lblsisa = Format(RST!nilaisisa, "#,##0.00")
    End If
    
'JURNAL PENJUALAN
    SQL = "Select * From gl_transaksi Where kdtrx = 'JJ' and notrx='" & txtkdaktiva & "'"
    SQL = SQL + " Order By tgltrx Asc"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Jurnal Jual tidak ditemukan", vbExclamation, AppName
    Else
        datetrx = RST!tgltrx
        Do Until RST.EOF
            grid.TextMatrix(grid.Row, 0) = RST!kdcomp
            grid.TextMatrix(grid.Row, 1) = RST!tgltrx
            grid.TextMatrix(grid.Row, 2) = RST!kdtrx
            grid.TextMatrix(grid.Row, 3) = RST!notrx
            grid.TextMatrix(grid.Row, 4) = Format(RST!kurs, "#,##0.00")
            grid.TextMatrix(grid.Row, 5) = RST!noactrx
            grid.TextMatrix(grid.Row, 6) = RST!desctrx
            grid.TextMatrix(grid.Row, 7) = RST!dbkrtrx
            grid.TextMatrix(grid.Row, 8) = Format(RST!amounttrx, "#,##0.00")
            grid.TextMatrix(grid.Row, 9) = Format(RST!nilaitrx, "#,##0.00")
            grid.TextMatrix(grid.Row, 10) = RST!currtrx
            grid.TextMatrix(grid.Row, 11) = RST!Flag
            grid.TextMatrix(grid.Row, 12) = RST!flagprint
            grid.TextMatrix(grid.Row, 13) = RST!flagadjustment
            grid.TextMatrix(grid.Row, 14) = RST!lineitem
            grid.TextMatrix(grid.Row, 15) = RST!identry
            grid.TextMatrix(grid.Row, 16) = RST!idupdate
            grid.TextMatrix(grid.Row, 17) = RST!dateentry
            grid.TextMatrix(grid.Row, 18) = RST!dateupdate
            grid.TextMatrix(grid.Row, 19) = RST!cekbg
            If IsNull(RST!reconsil) Then
                grid.TextMatrix(grid.Row, 20) = ""
            Else
                grid.TextMatrix(grid.Row, 20) = RST!reconsil
            End If
            grid.Col = 21
            Set grid.CellPicture = uncheck
            RST.MoveNext
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
        Loop
    End If
    
'BANK MASUK
    SQL = "Select * From gl_transaksi Where kdtrx = 'BM' and desctrx like '%" & txtkdaktiva & "%'"
    SQL = SQL + " Order By tgltrx Asc"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Bank Masuk tidak ditemukan", vbExclamation, AppName
        cmdAddBM.Caption = "Save"
    Else
        Do Until RST.EOF
            grid1.TextMatrix(grid1.Row, 0) = RST!kdcomp
            grid1.TextMatrix(grid1.Row, 1) = RST!tgltrx
            grid1.TextMatrix(grid1.Row, 2) = RST!kdtrx
            grid1.TextMatrix(grid1.Row, 3) = RST!notrx
            lblnoBM = RST!notrx
            grid1.TextMatrix(grid1.Row, 4) = Format(RST!kurs, "#,##0.00")
            grid1.TextMatrix(grid1.Row, 5) = RST!noactrx
            grid1.TextMatrix(grid1.Row, 6) = RST!desctrx
            grid1.TextMatrix(grid1.Row, 7) = RST!dbkrtrx
            grid1.TextMatrix(grid1.Row, 8) = Format(RST!amounttrx, "#,##0.00")
            grid1.TextMatrix(grid1.Row, 9) = Format(RST!nilaitrx, "#,##0.00")
            grid1.TextMatrix(grid1.Row, 10) = RST!currtrx
            grid1.TextMatrix(grid1.Row, 11) = RST!Flag
            grid1.TextMatrix(grid1.Row, 12) = RST!flagprint
            grid1.TextMatrix(grid1.Row, 13) = RST!flagadjustment
            grid1.TextMatrix(grid1.Row, 14) = RST!lineitem
            grid1.TextMatrix(grid1.Row, 15) = RST!identry
            grid1.TextMatrix(grid1.Row, 16) = RST!idupdate
            grid1.TextMatrix(grid1.Row, 17) = RST!dateentry
            grid1.TextMatrix(grid1.Row, 18) = RST!dateupdate
            grid1.TextMatrix(grid1.Row, 19) = RST!cekbg
            If IsNull(RST!reconsil) Then
                grid1.TextMatrix(grid1.Row, 20) = ""
            Else
                grid1.TextMatrix(grid1.Row, 20) = RST!reconsil
            End If
            RST.MoveNext
            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
        Loop
    End If
    
'JURNAL SUSUT
    SQL = "Select COUNT(kdcomp)'jml' From gl_transaksi Where kdtrx = 'JS' and notrx='" & txtkdaktiva & "'"
    Set RST = OBJ.Execute(SQL)
    Pg.Max = RST!jml
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select * From gl_transaksi Where kdtrx = 'JS' and notrx='" & txtkdaktiva & "'"
    SQL = SQL + " Order By tgltrx,dbkrtrx Asc"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Jurnal Susut tidak ditemukan", vbExclamation, AppName
        Exit Sub
    End If
    j = 0
    Do Until RST.EOF
        grid2.TextMatrix(grid2.Row, 0) = RST!kdcomp
        grid2.TextMatrix(grid2.Row, 1) = RST!tgltrx
        grid2.TextMatrix(grid2.Row, 2) = RST!kdtrx
        grid2.TextMatrix(grid2.Row, 3) = RST!notrx
        grid2.TextMatrix(grid2.Row, 4) = Format(RST!kurs, "#,##0.00")
        grid2.TextMatrix(grid2.Row, 5) = RST!noactrx
        grid2.TextMatrix(grid2.Row, 6) = RST!desctrx
        grid2.TextMatrix(grid2.Row, 7) = RST!dbkrtrx
        If grid2.TextMatrix(grid2.Row, 7) = "D" Then
            For j = 0 To grid2.Cols - 1
                grid2.Col = j
                grid2.CellBackColor = &H8000000C
            Next
        End If
        grid2.TextMatrix(grid2.Row, 8) = Format(RST!amounttrx, "#,##0.00")
        grid2.TextMatrix(grid2.Row, 9) = Format(RST!nilaitrx, "#,##0.00")
        grid2.TextMatrix(grid2.Row, 10) = RST!currtrx
        grid2.TextMatrix(grid2.Row, 11) = RST!Flag
        grid2.TextMatrix(grid2.Row, 12) = RST!flagprint
        grid2.TextMatrix(grid2.Row, 13) = RST!flagadjustment
        grid2.TextMatrix(grid2.Row, 14) = RST!lineitem
        grid2.TextMatrix(grid2.Row, 15) = RST!identry
        grid2.TextMatrix(grid2.Row, 16) = RST!idupdate
        grid2.TextMatrix(grid2.Row, 17) = RST!dateentry
        grid2.TextMatrix(grid2.Row, 18) = RST!dateupdate
        grid2.TextMatrix(grid2.Row, 19) = RST!cekbg
        If IsNull(RST!reconsil) Then
            grid2.TextMatrix(grid2.Row, 20) = ""
        Else
            grid2.TextMatrix(grid2.Row, 20) = RST!reconsil
        End If
        RST.MoveNext
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        Pg.Value = Pg.Value + 1
    Loop
    
    SQL = "Select SUM(amounttrx)'susut' From gl_transaksi Where kdtrx = 'JS' and notrx = '" & txtkdaktiva & "' and dbkrtrx = 'D' and flag = 'P'"
    Set RST = OBJ.Execute(SQL)
    If RST!susut = "0" Or IsNull(RST!susut) Then
        txtsusut = "0.00"
    Else
        txtsusut = RST!susut
    End If
    txtsisa = SpyRound(txthbeli - txtsusut)
    lblsisa = txtsisa
    Pg.Visible = False
    OBJ.Close
End Sub
Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function
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
        grid.TextMatrix(grid.Row, 14) = ""
        grid.TextMatrix(grid.Row, 15) = ""
        grid.TextMatrix(grid.Row, 16) = ""
        grid.TextMatrix(grid.Row, 17) = ""
        grid.TextMatrix(grid.Row, 18) = ""
        grid.TextMatrix(grid.Row, 19) = ""
        grid.TextMatrix(grid.Row, 20) = ""
        grid.Col = 21
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    
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
        grid1.TextMatrix(grid1.Row, 9) = ""
        grid1.TextMatrix(grid1.Row, 10) = ""
        grid1.TextMatrix(grid1.Row, 11) = ""
        grid1.TextMatrix(grid1.Row, 12) = ""
        grid1.TextMatrix(grid1.Row, 13) = ""
        grid1.TextMatrix(grid1.Row, 14) = ""
        grid1.TextMatrix(grid1.Row, 15) = ""
        grid1.TextMatrix(grid1.Row, 16) = ""
        grid1.TextMatrix(grid1.Row, 17) = ""
        grid1.TextMatrix(grid1.Row, 18) = ""
        grid1.TextMatrix(grid1.Row, 19) = ""
        grid1.TextMatrix(grid1.Row, 20) = ""
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
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
        grid2.TextMatrix(grid2.Row, 9) = ""
        grid2.TextMatrix(grid2.Row, 10) = ""
        grid2.TextMatrix(grid2.Row, 11) = ""
        grid2.TextMatrix(grid2.Row, 12) = ""
        grid2.TextMatrix(grid2.Row, 13) = ""
        grid2.TextMatrix(grid2.Row, 14) = ""
        grid2.TextMatrix(grid2.Row, 15) = ""
        grid2.TextMatrix(grid2.Row, 16) = ""
        grid2.TextMatrix(grid2.Row, 17) = ""
        grid2.TextMatrix(grid2.Row, 18) = ""
        grid2.TextMatrix(grid2.Row, 19) = ""
        grid2.TextMatrix(grid2.Row, 20) = ""
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
End Sub

Private Sub cmdprosesjs_Click()
    If MsgBox("Are you sure you want to change this data", _
    vbQuestion + vbYesNo, "Confirm Update Data") = vbNo Then Exit Sub
    
    grid2.Row = 1
    Do While True
        If Format(grid2.TextMatrix(grid2.Row, 1), "yyyy-MM-dd") > Format(datetrx, "yyyy-MM-dd") Then Exit Do
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 11) = "P"
        grid2.Row = grid2.Row + 1
    Loop
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        If Format(grid2.TextMatrix(grid2.Row, 1), "yyyy-MM-dd") > Format(datetrx, "yyyy-MM-dd") Then
            grid2.TextMatrix(grid2.Row, 11) = "J"
        End If
        grid2.Row = grid2.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "Update gl_transaksi Set flag='P' Where kdtrx = 'JS' and notrx = '" & txtkdaktiva & "' and tgltrx < convert(datetime,'" + Format(datetrx, "MM/dd/yyyy") + "')"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "Update gl_transaksi Set flag='J' Where kdtrx = 'JS' and notrx = '" & txtkdaktiva & "' and tgltrx > convert(datetime,'" + Format(datetrx, "MM/dd/yyyy") + "')"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "Select SUM(amounttrx)'susut' From gl_transaksi Where kdtrx = 'JS' and notrx = '" & txtkdaktiva & "' and dbkrtrx = 'D' and flag = 'P'"
    Set RST = OBJ.Execute(SQL)
    If RST!susut = "0" Or IsNull(RST!susut) Then
        txtsusut = "0.00"
    Else
        txtsusut = RST!susut
    End If
    txtsisa = txthbeli - txtsusut
    
    SQL = "Update gl_aktiva set nilaisisa = '" & txtsisa & "' Where kdaktiva = '" & txtkdaktiva & "'"
    Set RST = OBJ.Execute(SQL)
    
    OBJ.Close
    MsgBox "Depreciation Journal posted successfully", vbInformation, AppName

End Sub

Private Sub cmdsavejj_Click()
    If MsgBox("Are you sure you want to change this data" & _
    vbCrLf & "Changed data cannot be undo", vbQuestion + vbYesNo, "Confirm Update Data") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "Delete From gl_transaksi Where kdtrx = 'JJ' and notrx = '" & txtkdaktiva & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        With grid
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "insert into gl_transaksi ("
            SQL = SQL + "kdcomp,tgltrx,kdtrx,notrx,kurs,noactrx,desctrx,dbkrtrx,amounttrx,nilaitrx,currtrx,flag,flagprint,flagadjustment,lineitem,identry,idupdate,dateentry,dateupdate,cekbg,reconsil) VALUES('01',"
            SQL = SQL + "convert(datetime, '" + Format(datetrx, "MM/dd/yyyy") + "'),'JJ','"
            SQL = SQL + grid.TextMatrix(.Row, 3) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 4) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 5) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 6) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 7) + "',"
            SQL = SQL + "convert(money,'" + grid.TextMatrix(.Row, 8) + "'),"
            SQL = SQL + "convert(money,'" + grid.TextMatrix(.Row, 9) + "'),'"
            SQL = SQL + grid.TextMatrix(.Row, 10) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 11) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 12) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 13) + "',"
            SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
            SQL = SQL + "'','',"
            SQL = SQL + "convert(datetime, '" + Format(grid.TextMatrix(grid.Row, 17), "MM/dd/yyyy") + "'),"
            SQL = SQL + "convert(datetime, '" + Format(Now, "MM/dd/yyyy") + "'),'"
            SQL = SQL + grid.TextMatrix(.Row, 19) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 20) + "')"
            OBJ.Execute (SQL)
            .Row = .Row + 1
        End With
    Loop
    
    SQL = "Update gl_aktiva set tgljual = convert(datetime,'" & Format(datetrx, "MM/dd/yyyy") & "'),"
    SQL = SQL + " hargajual='" & txthjual & "',"
    SQL = SQL + " idupdate='" & UserOnline & "',"
    SQL = SQL + " dateupdate= convert(datetime,'" & Format(Now, "MM/dd/yyyy") & "'),"
    SQL = SQL + " nilaijual='" & txthjual & "' Where kdaktiva = '" & txtkdaktiva & "'"
    Set RST = OBJ.Execute(SQL)
    
    OBJ.Close
    MsgBox "Jurnal Penjualan " & txtkdaktiva & " successfuly update", vbInformation, AppName
    hapusgrid
    hapusgrid2
    Clearform
End Sub

Private Sub date1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date1.Visible = False
    
    If KeyCode = 13 Then
        grid.TextMatrix(posrow, 1) = Format(date1, "dd/MM/yyyy")
        grid.TextMatrix(posrow, 0) = "01"
        grid.TextMatrix(posrow, 2) = "JJ"
        grid.TextMatrix(posrow, 3) = txtkdaktiva
        grid.TextMatrix(posrow, 4) = grid.TextMatrix(1, 4)
        grid.TextMatrix(posrow, 10) = grid.TextMatrix(1, 10)
        grid.TextMatrix(posrow, 11) = grid.TextMatrix(1, 11)
        grid.TextMatrix(posrow, 12) = grid.TextMatrix(1, 12)
        grid.TextMatrix(posrow, 13) = grid.TextMatrix(1, 13)
        grid.TextMatrix(posrow, 14) = grid.Row
        grid.TextMatrix(posrow, 15) = grid.TextMatrix(1, 15)
        grid.TextMatrix(posrow, 16) = grid.TextMatrix(1, 16)
        grid.TextMatrix(posrow, 17) = grid.TextMatrix(1, 17)
        grid.TextMatrix(posrow, 18) = grid.TextMatrix(1, 18)
        grid.TextMatrix(posrow, 19) = grid.TextMatrix(1, 19)
        grid.TextMatrix(posrow, 20) = grid.TextMatrix(1, 20)
        grid.Col = 21
        Set grid.CellPicture = uncheck
            
        grid.SetFocus
        grid.Row = posrow
        date1.Visible = False

        If grid.Rows > grid.Row + 1 Then Exit Sub
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
    End If
End Sub

Private Sub date1_LostFocus()
    date1.Visible = False
End Sub

Private Sub date2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date2.Visible = False
    
    If KeyCode = 13 Then
        grid1.TextMatrix(posrow, 1) = Format(date2, "dd/MM/yyyy")
        grid1.TextMatrix(posrow, 0) = "01"
        grid1.TextMatrix(posrow, 2) = "BM"
        grid1.TextMatrix(posrow, 4) = grid.TextMatrix(1, 4)
        grid1.TextMatrix(posrow, 10) = grid.TextMatrix(1, 10)
        grid1.TextMatrix(posrow, 11) = "P"
        grid1.TextMatrix(posrow, 12) = grid.TextMatrix(1, 12)
        grid1.TextMatrix(posrow, 13) = grid.TextMatrix(1, 13)
        grid1.TextMatrix(posrow, 14) = "2"
        grid1.TextMatrix(posrow, 15) = grid.TextMatrix(1, 15)
        grid1.TextMatrix(posrow, 16) = grid.TextMatrix(1, 16)
        grid1.TextMatrix(posrow, 17) = grid.TextMatrix(1, 17)
        grid1.TextMatrix(posrow, 18) = grid.TextMatrix(1, 18)
        grid1.TextMatrix(posrow, 19) = grid.TextMatrix(1, 19)
        grid1.TextMatrix(posrow, 20) = grid.TextMatrix(1, 20)
        
        grid1.SetFocus
        grid1.Row = posrow
        date2.Visible = False
        If grid1.Rows > grid1.Row + 1 Then Exit Sub
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
    End If
End Sub

Private Sub date2_LostFocus()
    date2.Visible = False
End Sub

Private Sub Form_Load()
    cmbDK.additem "K"
    cmbDK.additem "D"
    cmbDK2.additem "K"
    cmbDK2.additem "D"
    cmbflag.additem "P"
    cmbflag.additem "J"
    cmbflag.additem "B"

    grid2.TextMatrix(0, 1) = "Tanggal"
    grid2.TextMatrix(0, 2) = "kode"
    grid2.TextMatrix(0, 3) = "No Transaksi"
    grid2.TextMatrix(0, 5) = "Account"
    grid2.TextMatrix(0, 6) = "Keterangan"
    grid2.TextMatrix(0, 7) = "D/K"
    grid2.TextMatrix(0, 8) = "Amount"
    grid2.TextMatrix(0, 11) = "Status"
    
    grid1.TextMatrix(0, 1) = "Tanggal"
    grid1.TextMatrix(0, 2) = "kode"
    grid1.TextMatrix(0, 3) = "No Transaksi"
    grid1.TextMatrix(0, 5) = "Account"
    grid1.TextMatrix(0, 6) = "Keterangan"
    grid1.TextMatrix(0, 7) = "D/K"
    grid1.TextMatrix(0, 8) = "Amount"
    grid1.TextMatrix(0, 11) = "Status"
    
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "kode"
    grid.TextMatrix(0, 3) = "No Transaksi"
    grid.TextMatrix(0, 5) = "Account"
    grid.TextMatrix(0, 6) = "Keterangan"
    grid.TextMatrix(0, 7) = "D/K"
    grid.TextMatrix(0, 8) = "Amount"
    grid.TextMatrix(0, 11) = "Status"
    grid.TextMatrix(0, 21) = " X"
    
    grid.Rows = 2
    grid.ColWidth(0) = 0 '900
    grid.ColWidth(1) = 1300
    grid.ColWidth(2) = 600
    grid.ColWidth(3) = 1200
    grid.ColWidth(4) = 0 '900
    grid.ColWidth(5) = 1300
    grid.ColWidth(6) = 4800
    grid.ColWidth(7) = 500
    grid.ColWidth(8) = 1300
    grid.ColWidth(9) = 0
    grid.ColWidth(10) = 0 '900
    grid.ColWidth(11) = 800
    grid.ColWidth(12) = 0
    grid.ColWidth(13) = 0
    grid.ColWidth(14) = 0
    grid.ColWidth(15) = 0
    grid.ColWidth(16) = 0
    grid.ColWidth(17) = 0
    grid.ColWidth(18) = 0
    grid.ColWidth(19) = 0
    grid.ColWidth(20) = 0
    grid.ColWidth(21) = 300
    
    grid1.Rows = 2
    grid1.ColWidth(0) = 0
    grid1.ColWidth(1) = 1300
    grid1.ColWidth(2) = 600
    grid1.ColWidth(3) = 1200
    grid1.ColWidth(4) = 0
    grid1.ColWidth(5) = 1300
    grid1.ColWidth(6) = 5100
    grid1.ColWidth(7) = 500
    grid1.ColWidth(8) = 1300
    grid1.ColWidth(9) = 0
    grid1.ColWidth(10) = 10
    grid1.ColWidth(11) = 800
    grid1.ColWidth(12) = 0
    grid1.ColWidth(13) = 0
    grid1.ColWidth(14) = 0
    grid1.ColWidth(15) = 0
    grid1.ColWidth(16) = 0
    grid1.ColWidth(17) = 0
    grid1.ColWidth(18) = 0
    grid1.ColWidth(19) = 0
    grid1.ColWidth(20) = 0
    
    grid2.Rows = 2
    grid2.ColWidth(0) = 0
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 600
    grid2.ColWidth(3) = 1200
    grid2.ColWidth(4) = 0
    grid2.ColWidth(5) = 1300
    grid2.ColWidth(6) = 5100
    grid2.ColWidth(7) = 500
    grid2.ColWidth(8) = 1300
    grid2.ColWidth(9) = 0
    grid2.ColWidth(10) = 0
    grid2.ColWidth(11) = 800
    grid2.ColWidth(12) = 0
    grid2.ColWidth(13) = 0
    grid2.ColWidth(14) = 0
    grid2.ColWidth(15) = 0
    grid2.ColWidth(16) = 0
    grid2.ColWidth(17) = 0
    grid2.ColWidth(18) = 0
    grid2.ColWidth(19) = 0
    grid2.ColWidth(20) = 0
    
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(5) = flexAlignRightCenter
    grid.ColAlignment(8) = flexAlignRightCenter
    
    grid1.ColAlignment(1) = flexAlignLeftCenter
    grid1.ColAlignment(5) = flexAlignRightCenter
    grid1.ColAlignment(8) = flexAlignRightCenter
    
    grid2.ColAlignment(1) = flexAlignLeftCenter
    grid2.ColAlignment(5) = flexAlignRightCenter
    grid2.ColAlignment(8) = flexAlignRightCenter
    
    TabControl1.SelectedItem = 0
    date1.Visible = False
    txtstring.Visible = False
    cmbDK.Visible = False
    txtnilai.Visible = False
End Sub
Private Sub Clearform()
    txtkdaktiva = ""
    lblaktiva = ""
    txthbeli = ""
    txthjual = ""
    datebeli = Date
    datetrx = Date
    lblsisa = ""
    lblsusut = ""
    txtsusut = 0
    txtsisa = 0
End Sub
Private Sub grid_Click()
    posrow = grid.Row
    Select Case grid.Col
        Case 1
            If date1.Visible = True Then Exit Sub
            
            date1.Width = grid.ColWidth(grid.Col) - 20
            date1.Height = 290
            'If grid.TextMatrix(grid.Row, grid.Col) <> "" Then date1 = grid.TextMatrix(grid.Row, 3)
            date1.Left = grid.Left + grid.CellLeft - 10
            date1.Top = grid.Top + grid.CellTop - 20
            date1.Visible = True
            date1 = Date
            date1.SetFocus
        Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            carisql1 = "Select noac,nmac From gl_masterac"
            namatabel = "Account"
            frmsearch.Show vbModal
        Case 6
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If txtstring.Visible = True Then Exit Sub
            
            txtstring.Width = grid.ColWidth(grid.Col) - 40
            txtstring = grid.TextMatrix(grid.Row, grid.Col)
            txtstring.Left = grid.Left + grid.CellLeft
            txtstring.Top = grid.Top + grid.CellTop - 30
            txtstring.Visible = True
            txtstring.SetFocus
            If grid.Col = 6 Then txtstring = grid.TextMatrix(1, 6)
        Case 7
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If cmbDK.Visible = True Then Exit Sub
            
            cmbDK.Width = grid.ColWidth(grid.Col) - 40
            cmbDK = grid.TextMatrix(grid.Row, grid.Col)
            cmbDK.Left = grid.Left + grid.CellLeft
            cmbDK.Top = grid.Top + grid.CellTop - 30
            cmbDK.Visible = True
            cmbDK.SetFocus
        Case 8
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 21
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
    End Select
End Sub

Private Sub grid_GotFocus()
    posrow = grid.Row
    Select Case grid.Col
        Case 5
            If hasil = "" Then Exit Sub
            grid.TextMatrix(grid.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub grid1_Click()
    posrow = grid1.Row
    Select Case grid1.Col
        Case 1
            If date2.Visible = True Then Exit Sub
            
            date2.Width = grid1.ColWidth(grid1.Col) - 20
            date2.Height = 290
            date2.Left = grid1.Left + grid1.CellLeft - 10
            date2.Top = grid1.Top + grid1.CellTop - 20
            date2.Visible = True
            date2 = Date
            date2.SetFocus
        Case 3
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            lblinfo.Visible = True
            If txtnotrx.Visible = True Then Exit Sub
            
            txtnotrx.Width = grid1.ColWidth(grid1.Col) - 40
            txtnotrx = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnotrx.Left = grid1.Left + grid1.CellLeft
            txtnotrx.Top = grid1.Top + grid1.CellTop - 30
            txtnotrx.Visible = True
            txtnotrx.SetFocus
        Case 5
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            carisql1 = "Select noac,nmac From gl_masterac"
            namatabel = "Account"
            frmsearch.Show vbModal
        Case 6
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If txtdesc.Visible = True Then Exit Sub
            
            txtdesc.Width = grid1.ColWidth(grid1.Col) - 40
            txtdesc = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtdesc.Left = grid1.Left + grid1.CellLeft
            txtdesc.Top = grid1.Top + grid1.CellTop - 30
            txtdesc.Visible = True
            txtdesc.SetFocus
            If grid1.Col = 6 Then txtdesc = grid.TextMatrix(1, 6)
        Case 7
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If cmbDK2.Visible = True Then Exit Sub
            
            cmbDK2.Width = grid1.ColWidth(grid1.Col) - 40
            cmbDK2 = grid1.TextMatrix(grid1.Row, grid1.Col)
            cmbDK2.Left = grid1.Left + grid1.CellLeft
            cmbDK2.Top = grid1.Top + grid1.CellTop - 30
            cmbDK2.Visible = True
            cmbDK2.SetFocus
        Case 8
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            txtamount.Width = grid1.ColWidth(grid1.Col) - 40
            txtamount = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtamount.Left = grid1.Left + grid1.CellLeft
            txtamount.Top = grid1.Top + grid1.CellTop + 20
            txtamount.Visible = True
            txtamount.SetFocus
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
    grid.TextMatrix(grid.Row, 9) = ""
    grid.TextMatrix(grid.Row, 10) = ""
    grid.TextMatrix(grid.Row, 11) = ""
    grid.TextMatrix(grid.Row, 12) = ""
    grid.TextMatrix(grid.Row, 13) = ""
    grid.TextMatrix(grid.Row, 14) = ""
    grid.TextMatrix(grid.Row, 15) = ""
    grid.TextMatrix(grid.Row, 16) = ""
    grid.TextMatrix(grid.Row, 17) = ""
    grid.TextMatrix(grid.Row, 18) = ""
    grid.TextMatrix(grid.Row, 19) = ""
    grid.TextMatrix(grid.Row, 20) = ""
    grid.TextMatrix(grid.Row, 21) = ""
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
            grid.TextMatrix(grid.Row, 14) = ""
            grid.TextMatrix(grid.Row, 15) = ""
            grid.TextMatrix(grid.Row, 16) = ""
            grid.TextMatrix(grid.Row, 17) = ""
            grid.TextMatrix(grid.Row, 18) = ""
            grid.TextMatrix(grid.Row, 19) = ""
            grid.TextMatrix(grid.Row, 20) = ""
            grid.TextMatrix(grid.Row, 21) = ""
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
        grid.TextMatrix(grid.Row, 14) = grid.TextMatrix(grid.Row + 1, 14)
        grid.TextMatrix(grid.Row, 15) = grid.TextMatrix(grid.Row + 1, 15)
        grid.TextMatrix(grid.Row, 16) = grid.TextMatrix(grid.Row + 1, 16)
        grid.TextMatrix(grid.Row, 17) = grid.TextMatrix(grid.Row + 1, 17)
        grid.TextMatrix(grid.Row, 18) = grid.TextMatrix(grid.Row + 1, 18)
        grid.TextMatrix(grid.Row, 19) = grid.TextMatrix(grid.Row + 1, 19)
        grid.TextMatrix(grid.Row, 20) = grid.TextMatrix(grid.Row + 1, 20)
        grid.TextMatrix(grid.Row, 21) = grid.TextMatrix(grid.Row + 1, 21)
        
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 21
    Set grid.CellPicture = blank
End Sub

Private Sub grid1_GotFocus()
    posrow = grid1.Row
    Select Case grid1.Col
        Case 5
            If hasil = "" Then Exit Sub
            grid1.TextMatrix(grid1.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub grid2_Click()
    posrow = grid2.Row
    Select Case grid2.Col
        Case 11:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If cmbDK.Visible = True Then Exit Sub
            
            cmbflag.Width = grid2.ColWidth(grid2.Col) - 40
            cmbflag = grid2.TextMatrix(grid2.Row, grid2.Col)
            cmbflag.Left = grid2.Left + grid2.CellLeft
            cmbflag.Top = grid2.Top + grid2.CellTop - 30
            cmbflag.Visible = True
            cmbflag.SetFocus
    End Select
    
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Caption
        Case "Jurnal Penjualan ( JJ )":
                date1.Visible = False
                txtstring.Visible = False
                cmbDK.Visible = False
                txtnilai.Visible = False
        Case "Bank Masuk ( BM )":
                date2.Visible = False
                txtnotrx.Visible = False
                txtdesc.Visible = False
                lblinfo.Visible = False
                cmbDK2.Visible = False
                txtamount.Visible = False
        Case "Jurnal Penyusutan ( JS )":
                cmbflag.Visible = False
    End Select
End Sub

Private Sub txtamount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtamount.Visible = False
    
    If KeyCode = 13 Then
        If grid1.TextMatrix(posrow, 1) = "" Then Exit Sub
        grid1.TextMatrix(posrow, 8) = Format(txtamount, "#,##0.00")
        grid1.TextMatrix(posrow, 9) = Format(txtamount, "#,##0.00")
        
        grid1.SetFocus
        grid1.Row = posrow
        txtamount.Visible = False
    End If
End Sub

Private Sub txtamount_LostFocus()
    txtamount.Visible = False
End Sub

Private Sub txtdesc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtdesc.Visible = False
    
    If KeyCode = 13 Then
        Select Case grid1.Col
            Case 6: grid1.TextMatrix(posrow, 6) = txtdesc
        End Select
        grid1.SetFocus
        grid1.Row = posrow
        txtdesc.Visible = False
    End If
End Sub

Private Sub txtdesc_LostFocus()
    txtdesc.Visible = False
End Sub

Private Sub txtjumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Update gl_transaksi set amounttrx='" & txtjumlah & "',nilaitrx='" & txtjumlah & "'"
        SQL = SQL + " Where notrx= '" & txtkdaktiva & "' and kdtrx='JS'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        hapusgrid
        hapusgrid2
        Clearform
        
    End If
End Sub

Private Sub txtkdaktiva_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtkdaktiva = "" Then Exit Sub
        hapusgrid
        hapusgrid2
        showFA
    End If
End Sub

Private Sub txtnilai_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtnilai.Visible = False
    
    If KeyCode = 13 Then
        grid.TextMatrix(posrow, 8) = Format(txtnilai, "#,##0.00")
        grid.TextMatrix(posrow, 9) = Format(txtnilai, "#,##0.00")
        
        grid.SetFocus
        grid.Row = posrow
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnotrx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, 3) = txtnotrx
        txtnotrx.Visible = False
    End If
End Sub

Private Sub txtnotrx_KeyUp(KeyCode As Integer, Shift As Integer)
    cari_in
End Sub

Private Sub txtnotrx_LostFocus()
    txtnotrx.Visible = False
End Sub

Private Sub txtstring_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtstring.Visible = False
    
    If KeyCode = 13 Then
        Select Case grid.Col
            Case 6: grid.TextMatrix(posrow, 6) = txtstring
        End Select
        grid.SetFocus
        grid.Row = posrow
        txtstring.Visible = False
    End If
End Sub

Private Sub cari_in()
    If grid1.TextMatrix(1, 2) = "BM" Then
        If Len(txtnotrx) = 8 Then
            If Not (Left(txtnotrx, 2) >= "08" And Left(txtnotrx, 2) < "99") Then
                MsgBox "Format digit pertama dan kedua salah, format yang dipakai adalah format tahun, YY", vbInformation, "Information"
                txtnotrx = ""
                txtnotrx.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotrx, 3, 2) >= "01" And Mid(txtnotrx, 3, 2) <= "12") Then
                MsgBox "Format digit ketiga dan keempat salah, format yang dipakai adalah format bulan, MM", vbInformation, "Information"
                txtnotrx = ""
                txtnotrx.SetFocus
                Exit Sub
            End If
            If Not Mid(txtnotrx, 5, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotrx = ""
                txtnotrx.SetFocus
                Exit Sub
            End If
            If Not Right(txtnotrx, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotrx = ""
                txtnotrx.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotrx, 6, 2) >= "01" And Mid(txtnotrx, 6, 2) <= "09") And Not Mid(txtnotrx, 6, 2) <= "99" Then
                MsgBox "Format digit keenam dan ketujuh salah, tekan F2 untuk melihat list.", vbInformation, "Information"
                txtnotrx = ""
                txtnotrx.SetFocus
                Exit Sub
            End If
            
            OBJ.Open dsn
            SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '01' and kdtrx = 'BM' and notrx like '" & txtnotrx & "%' and flagprint='I' order by notrx desc"
            Set RST = OBJ.Execute(SQL)
          
            If Not RST.EOF Then
                If Len(RST!notrx + 1) = 5 Then
                    txtnotrx = txtnotrx & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 4 Then
                    txtnotrx = txtnotrx & "0" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 3 Then
                    txtnotrx = txtnotrx & "00" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 2 Then
                    txtnotrx = txtnotrx & "000" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 1 Then
                    txtnotrx = txtnotrx & "0000" & RST!notrx + 1
                End If
            Else
                txtnotrx = txtnotrx & "00001"
            End If
            OBJ.Close

        End If
    Else
        OBJ.Open dsn
        SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '01' and kdtrx = 'BM' and notrx like '" & txtnotrx & "%' and flagprint='I' order by notrx desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If Len(RST!notrx + 1) = 5 Then
                txtnotrx = txtnotrx & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 4 Then
                txtnotrx = txtnotrx & "0" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 3 Then
                txtnotrx = txtnotrx & "00" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 2 Then
                txtnotrx = txtnotrx & "000" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 1 Then
                txtnotrx = txtnotrx & "0000" & RST!notrx + 1
            End If
        Else
            txtnotran = txtnotrx & "/00001"
        End If
        OBJ.Close
    End If
End Sub

Private Sub txtstring_LostFocus()
    txtstring.Visible = False
End Sub
