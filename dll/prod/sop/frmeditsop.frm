VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmeditsop 
   Caption         =   "EDIT FORMULA SOP PROSES"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   14685
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picmanual 
      Height          =   735
      Left            =   4920
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   49
      Top             =   6240
      Visible         =   0   'False
      Width           =   2655
      Begin TDBNumber6Ctl.TDBNumber txthppmanual 
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmeditsop.frx":0000
         Caption         =   "frmeditsop.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmeditsop.frx":008C
         Keys            =   "frmeditsop.frx":00AA
         Spin            =   "frmeditsop.frx":00EC
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
      Begin TDBNumber6Ctl.TDBNumber txtqty 
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calculator      =   "frmeditsop.frx":0114
         Caption         =   "frmeditsop.frx":0134
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmeditsop.frx":01A0
         Keys            =   "frmeditsop.frx":01BE
         Spin            =   "frmeditsop.frx":0200
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
      Begin TDBNumber6Ctl.TDBNumber txtrow 
         Height          =   255
         Left            =   1440
         TabIndex        =   54
         Top             =   1080
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calculator      =   "frmeditsop.frx":0228
         Caption         =   "frmeditsop.frx":0248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmeditsop.frx":02B4
         Keys            =   "frmeditsop.frx":02D2
         Spin            =   "frmeditsop.frx":0314
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
      Begin VB.Label lblitem 
         Caption         =   "."
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
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "Row"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Harga /Kg"
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
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Qty"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   840
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.ProgressBar Pb 
      Height          =   165
      Left            =   4920
      TabIndex        =   47
      Top             =   6300
      Visible         =   0   'False
      Width           =   9690
      _Version        =   851970
      _ExtentX        =   17092
      _ExtentY        =   291
      _StockProps     =   93
   End
   Begin TDBNumber6Ctl.TDBNumber txtnourut 
      Height          =   255
      Left            =   7560
      TabIndex        =   46
      Top             =   75
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmeditsop.frx":033C
      Caption         =   "frmeditsop.frx":035C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmeditsop.frx":03C8
      Keys            =   "frmeditsop.frx":03E6
      Spin            =   "frmeditsop.frx":0428
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
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   9255
      TabIndex        =   39
      Top             =   15
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmeditsop.frx":0450
      Caption         =   "frmeditsop.frx":0470
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmeditsop.frx":04DC
      Keys            =   "frmeditsop.frx":04FA
      Spin            =   "frmeditsop.frx":053C
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
   Begin VB.TextBox txtnolot1 
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
      Height          =   195
      Left            =   5055
      TabIndex        =   38
      Top             =   105
      Visible         =   0   'False
      Width           =   2010
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
      TabIndex        =   31
      Top             =   6630
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
      Left            =   7065
      Picture         =   "frmeditsop.frx":0564
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
      Top             =   6630
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
      Left            =   6810
      Picture         =   "frmeditsop.frx":0846
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   6630
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6840
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   4830
      Begin XtremeSuiteControls.CheckBox chkedit 
         Height          =   375
         Left            =   3120
         TabIndex        =   57
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851970
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Edit Lot"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtnolot_old 
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
         Left            =   3015
         TabIndex        =   43
         Top             =   1245
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.ComboBox cmbqc 
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
         Left            =   1755
         TabIndex        =   24
         Top             =   6210
         Width           =   2145
      End
      Begin VB.TextBox txttesvisual 
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
         Left            =   1770
         TabIndex        =   23
         Top             =   5115
         Width           =   2130
      End
      Begin VB.TextBox txtviskositas 
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
         Left            =   1770
         TabIndex        =   22
         Top             =   5475
         Width           =   2130
      End
      Begin VB.TextBox txtsolid 
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
         Left            =   1770
         TabIndex        =   21
         Top             =   5835
         Width           =   2130
      End
      Begin VB.TextBox txtwaktukemasan 
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
         Left            =   1785
         TabIndex        =   17
         Top             =   4575
         Width           =   900
      End
      Begin VB.TextBox txtwaktutambahan 
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
         Left            =   1785
         TabIndex        =   16
         Top             =   4215
         Width           =   900
      End
      Begin VB.TextBox txtwaktupelarutan 
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
         Left            =   1785
         TabIndex        =   15
         Top             =   3855
         Width           =   900
      End
      Begin VB.TextBox txtnobpb 
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
         Left            =   3015
         TabIndex        =   14
         Top             =   1620
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1230
         Left            =   210
         TabIndex        =   13
         Top             =   2445
         Width           =   4215
         Begin TDBNumber6Ctl.TDBNumber txttotalproduksi 
            Height          =   315
            Left            =   2475
            TabIndex        =   33
            Top             =   330
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   556
            Calculator      =   "frmeditsop.frx":0B94
            Caption         =   "frmeditsop.frx":0BB4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmeditsop.frx":0C20
            Keys            =   "frmeditsop.frx":0C3E
            Spin            =   "frmeditsop.frx":0C88
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999999
            MinValue        =   -9999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber txttotalhasilproduksi 
            Height          =   315
            Left            =   2475
            TabIndex        =   34
            Top             =   735
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   556
            Calculator      =   "frmeditsop.frx":0CB0
            Caption         =   "frmeditsop.frx":0CD0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmeditsop.frx":0D3C
            Keys            =   "frmeditsop.frx":0D5A
            Spin            =   "frmeditsop.frx":0DA4
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber txttotalHPP 
            Height          =   315
            Left            =   450
            TabIndex        =   44
            Top             =   495
            Visible         =   0   'False
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   556
            Calculator      =   "frmeditsop.frx":0DCC
            Caption         =   "frmeditsop.frx":0DEC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmeditsop.frx":0E58
            Keys            =   "frmeditsop.frx":0E76
            Spin            =   "frmeditsop.frx":0EC0
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.Label hpp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   255
            Left            =   60
            TabIndex        =   40
            Top             =   960
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL PRODUKSI :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   195
            TabIndex        =   37
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL HASIL PRODUKSI :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   780
            Width           =   2370
         End
         Begin VB.Label tg1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   255
            Left            =   180
            TabIndex        =   35
            Top             =   135
            Visible         =   0   'False
            Width           =   1605
         End
      End
      Begin VB.TextBox txtnoreaktor 
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
         Left            =   1155
         TabIndex        =   9
         Top             =   1620
         Width           =   1455
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
         Left            =   1155
         TabIndex        =   8
         Top             =   450
         Width           =   975
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
         Left            =   2175
         TabIndex        =   7
         Top             =   450
         Width           =   2550
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
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   825
         Width           =   2325
      End
      Begin XtremeSuiteControls.DateTimePicker datebahan 
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Top             =   1230
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.PushButton cmdnolot 
         Height          =   315
         Left            =   105
         TabIndex        =   6
         Top             =   825
         Width           =   990
         _Version        =   851970
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "NO LOT "
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
      Begin XtremeSuiteControls.DateTimePicker datedone 
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   2025
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
      Begin XtremeSuiteControls.PushButton btneditnolot 
         Height          =   345
         Left            =   3555
         TabIndex        =   42
         Top             =   810
         Width           =   1185
         _Version        =   851970
         _ExtentX        =   2090
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Update No Lot"
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lulus / Tidak :"
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
         Left            =   315
         TabIndex        =   28
         Top             =   6270
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tes Visual :"
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
         Left            =   315
         TabIndex        =   27
         Top             =   5145
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Viskositas mPA.s :"
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
         Left            =   315
         TabIndex        =   26
         Top             =   5520
         Width           =   1380
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid (%) :"
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
         Left            =   315
         TabIndex        =   25
         Top             =   5880
         Width           =   1380
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Pengemasan :"
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
         Left            =   105
         TabIndex        =   20
         Top             =   4605
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Tambahan :"
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
         Left            =   225
         TabIndex        =   19
         Top             =   4245
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Pelarutan :"
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
         Left            =   330
         TabIndex        =   18
         Top             =   3900
         Width           =   1380
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NO KOCEKAN "
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
         Left            =   -195
         TabIndex        =   12
         Top             =   1665
         Width           =   1290
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TGL SELESAI  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2070
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TANGGAL"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1275
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PRODUK "
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
         Left            =   315
         TabIndex        =   3
         Top             =   510
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   6090
      Left            =   4920
      TabIndex        =   0
      Top             =   210
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   10742
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
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
      Height          =   450
      Left            =   13590
      TabIndex        =   32
      Top             =   6495
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   794
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
   Begin XtremeSuiteControls.PushButton btnUpdate 
      Height          =   450
      Left            =   12525
      TabIndex        =   41
      Top             =   6495
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "UPDATE"
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
   Begin XtremeSuiteControls.PushButton btnClear 
      Height          =   450
      Left            =   11445
      TabIndex        =   45
      Top             =   6495
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   794
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
   Begin Chameleon.chameleonButton cmdunlock 
      Height          =   375
      Left            =   4920
      TabIndex        =   48
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Unlock SOP"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmeditsop.frx":0EE8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmeditsop"
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
Private kdlot As String
Private akses As Boolean

Private poscol As Integer
Private posrow As Integer

Private Sub btnclear_Click()
    txtproduk = ""
    txtkodeproduk = ""
    txtnolot = ""
    txtnolot_old = ""
    datebahan = Date
    datedone = Date
    txtnobpb = ""
    txtnoreaktor = ""
    txttotalhasilproduksi = 0
    txttotalHPP = 0
    txttotalproduksi = 0
    tg1 = ""
    hpp = ""
    txtwaktupelarutan = ""
    txtwaktukemasan = ""
    txtwaktutambahan = ""
    txttesvisual = ""
    txtviskositas = ""
    txtsolid = ""
    cmbqc = ""
    statuslot = False
    hapusgrid1
End Sub

Private Sub btnClose_Click()
    statuslot = False
    Unload Me
End Sub

Private Sub btneditnolot_Click()
    If txtkodeproduk = "" Or txtnolot = "" Then
        MsgBox "Data tidak lengkap", vbCritical, AppName
        Exit Sub
    End If
    If chkedit.Visible = False Then
        OBJ.Open dsn
        SQL = "Select * From list_produksi_hasil Where nolot='" & txtnolot_old & "'"
        Set RST = OBJ.Execute(SQL)
        
        If Not RST.EOF Then
            MsgBox "Palet sudah di scan, nomor lot tidak bisa diubah", vbCritical, AppName
            OBJ.Close
            Exit Sub
        End If
        OBJ.Close
    Else
        If MsgBox("Nolot lebih dari 1 palet tidak bisa diupdate", vbOKCancel + vbExclamation, "WARNING..") = vbCancel Then Exit Sub
    End If
    
    If MsgBox("Anda Yakin ingin mengganti nolot ini ?" & vbCrLf & _
    "NO LOT : " & txtnolot_old & vbCrLf & _
    "NO LOT BARU : " & txtnolot, vbQuestion + vbYesNo, "Konfirmasi !") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    'list_produksi_master
    SQL = "Update list_produksi_master set nolot='" & txtnolot & "'"
    SQL = SQL + " Where nolot='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    
    'list_produksi_child
    SQL = "Update list_produksi_child set nolot='" & txtnolot & "',ref='" & txtnolot & "/1'"
    SQL = SQL + " Where nolot='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    
    'list_historisop
    SQL = "Update list_historisop set nolot='" & txtnolot & "'"
    SQL = SQL + " Where nolot='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    
    'am_usehdr
    SQL = "Update am_usehdr set nobpb='" & txtnolot & "/1'"
    SQL = SQL + " Where nobpb='" & txtnolot_old & "/1'"
    Set RST = OBJ.Execute(SQL)
    
    'am_uselin
    SQL = "Update am_uselin set nobpb='" & txtnolot & "/1'"
    SQL = SQL + " Where nobpb='" & txtnolot_old & "/1'"
    Set RST = OBJ.Execute(SQL)
    
    'list_masterkeylot
    SQL = "UPDATE list_masterkeylot set noso='" & txtnolot & "'"
    SQL = SQL + " Where noso='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    
    'list_historicetaksop
    SQL = "UPDATE list_historicetaksop set nolot='" & txtnolot & "'"
    SQL = SQL + " Where nolot='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    
    'Cek Lot Base
    SQL = "Select * from am_sopbase Where nolot='" & txtnolot_old & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
    'Update lot base di am_sopbase
        SQL = "Update am_sopbase set nolot='" & txtnolot & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    
    If chkedit.Value = xtpChecked Then
        'list_produksi_hasil
        SQL = "UPDATE list_produksi_hasil set nolot='" & txtnolot & "',noref= '01" & txtnolot & "'"
        SQL = SQL + " Where nolot='" & txtnolot_old & "'"
        Set RST = OBJ.Execute(SQL)
        
        'list_produksi_kemasan
        SQL = "UPDATE list_produksi_kemasan set nolot='" & txtnolot & "',noref='01" & txtnolot & "'"
        SQL = SQL + " Where nolot='" & txtnolot_old & "'"
        Set RST = OBJ.Execute(SQL)
        
    End If
    
    OBJ.Close
    MsgBox "No Lot is successfully updated", vbInformation, AppName
    btnclear_Click
End Sub

Private Sub btnupdate_Click()
On Error GoTo Err_handler:
    If txtnolot = "" Then Exit Sub
    If grid1.TextMatrix(1, 1) = "" Then Exit Sub
    
    If MsgBox("Anda yakin ingin merubah data formula SOP " & vbCrLf & _
    "NOLOT : " & txtnolot & vbCrLf & "Pastikan terlebih dahulu data yang anda ubah sudah benar" _
    , vbQuestion + vbOKCancel, "KONFIRMASI !") = vbCancel Then Exit Sub
    
    If MsgBox("Apakah nomor urut SOP sudah benar ?", vbQuestion + vbYesNo, "KONFIRMASI !") = vbNo Then Exit Sub
    
    Pb.Max = (grid1.Row - 2) * 2
    Pb.Value = 0
    Pb.Visible = True
    
    OBJ.Open dsn
    'update list_produksi_master
    SQL = "Update list_produksi_master "
    SQL = SQL + "Set tanggal='" & Format(datebahan, "yyyy/MM/dd") & "',"
    SQL = SQL + "total_produksi='" & txttotalproduksi & "',"
    SQL = SQL + "total_hasil_produksi='" & txttotalhasilproduksi & "',"
    SQL = SQL + "tglakhir='" & Format(datedone, "yyyy/MM/dd") & "',"
    SQL = SQL + "noreaktor='" & txtnoreaktor & "',"
    SQL = SQL + "userupdate='" & nmuser & "',"
    SQL = SQL + "hpp='" & txttotalHPP & "' "
    SQL = SQL + "Where nolot = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    
    'hapus data bahan baku
    SQL = "delete from list_produksi_child where nolot='" & txtnolot & "' and proses_ke='1'"
    OBJ.Execute SQL
    
    'Update list_produksi_child
    SQL = "Select * From list_produksi_child Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
            RST!kode_produk = txtkodeproduk
            RST!nolot = txtnolot
            RST!kode_bahan = grid1.TextMatrix(grid1.Row, 1)
            RST!Lot_bahan = grid1.TextMatrix(grid1.Row, 3)
            RST!qty_bahan = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
            RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
            RST!flag_tambahan = "0"
            RST!hpp = Format(grid1.TextMatrix(grid1.Row, 7), "general number")
            RST!tanggal = Format(datedone, "yyyy/MM/dd")
            RST!REF = txtnolot & "/1"
            RST!Line = Format(grid1.TextMatrix(grid1.Row, 8), "general number")
            RST!proses_ke = "1"
            RST.Update
        Pb.Value = Pb.Value + 1
        grid1.Row = grid1.Row + 1
    Loop
    
    'Update LIST_HISTORISOP
    SQL = "Update LIST_HISTORISOP "
    SQL = SQL + "set TANGGAL='" & Format(datebahan, "yyyy/MM/dd") & "' "
    SQL = SQL + "Where nolot = '" & txtnolot & "' and PROSES_KE = '1'"
    Set RST = OBJ.Execute(SQL)
    
    'Update am_usehdr
    SQL = "Update am_usehdr "
    SQL = SQL + "set tglbpb='" & Format(datebahan, "yyyy/MM/dd") & "' "
    SQL = SQL + "Where nobpb='" & txtnolot & "/1'"
    Set RST = OBJ.Execute(SQL)
    
    'Hapus data bahan baku
    SQL = "Delete From am_uselin Where nobpb='" & txtnolot & "/1'"
    OBJ.Execute (SQL)
    
    'Update am_uselin
    SQL = "Select * From am_uselin where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!nobpb = txtnolot & "/1"
        RST!kodebarang = grid1.TextMatrix(grid1.Row, 1)
        RST!qty = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
        RST!kodesatuan = grid1.TextMatrix(grid1.Row, 5)
        RST!lineitem = Format(grid1.TextMatrix(grid1.Row, 8), "general number")
        RST.Update
        Pb.Value = Pb.Value + 1
        grid1.Row = grid1.Row + 1
    Loop
    
    'Update list_historicetaksop
    SQL = "Update list_historicetaksop set TANGGAL='" & Format(datebahan, "yyyy/MM/dd") & "' "
    SQL = SQL + "Where NOLOT = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    
    'Lock edit sop
    SQL = "Update list_masterkeyLot set otoritas='0' Where noso='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    Pb.Visible = False
    MsgBox "No LOT : " & txtnolot & " is Sucssesfuly updated.", vbInformation, AppName
    btnclear_Click
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdnolot_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    'carisql1 = carisql1 + " where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnolot_GotFocus()
    If hasil = "" Then Exit Sub
    If nmuser = "martsanto" Or nmuser = "Creator" Or nmuser = "angelgunawan" Or nmuser = "Angle" _
    Or nmuser = "kimlie" Or nmuser = "putri" Or nmuser = "HADY" Or nmuser = "bina" Or nmuser = "ENAH" Then GoTo jump:
        
    
'Periksa Otoritas edit sop
    OBJ.Open dsn
    SQL = "Select * From list_masterkeyLot"
    SQL = SQL + " Where noso ='" & hasil2 & "' and otoritas = '1'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        OBJ.Close
        MsgBox "Maaf Anda tidak memiliki otoritas, Silahkan hubungi Administrator Anda", vbCritical, AppName
        hasil = ""
        hasil1 = ""
        hasil2 = ""
        carisql1 = ""
        Exit Sub
    End If
    OBJ.Close
jump:
    
    txtkodeproduk = hasil
    txtproduk = hasil1
    txtnolot = hasil2
    txtnolot_old = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
    OpenHeader
    posrow = 0
    poscol = 0
End Sub

Private Sub OpenHeader()
    On Error GoTo Err_handler:
    Dim totalhasilproduksi As Double
    OBJ.Open dsn
    SQL = "select a.*,b.nama_produk from list_produksi_master a"
    SQL = SQL + " inner join list_produk_master b on a.kode_produk = b.kode_produk"
    SQL = SQL + " where nolot='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    txtkodeproduk = RST!kode_produk
    txtproduk = RST!nama_produk
    txtnolot = RST!nolot
    txtnoreaktor = Format(RST!noreaktor, "##,###,##0")
    txtnobpb = getnobpb(Format(Date, "yymm"))
    datebahan = RST!tanggal
    datedone = RST!tglakhir
    tg1 = RST!total_produksi
    txttotalhasilproduksi = RST!total_hasil_produksi
    txtwaktupelarutan = RST!waktu_larut
    txtwaktutambahan = RST!waktu_tambahan
    txtwaktukemasan = RST!waktu_kemasan
    txttesvisual = RST!qc_test_visual
    txtviskositas = RST!qc_viskositas
    txtsolid = RST!qc_solid
    If RST!flag_status = "0" Then
        cmbqc.text = "Lulus"
    Else
        cmbqc.text = "Tidak"
    End If
    hpp = RST!hpp
    txttotalHPP = RST!hpp
    
    SQL = "select distinct a.*,isnull(b.nama_bahan,d.namabarang) as namabahan,c.namasatuan from list_produksi_child a "
    SQL = SQL + "left join list_produk_child b on a.kode_produk = b.kode_produk and  "
    SQL = SQL + "a.kode_bahan=b.kode_bahan "
    SQL = SQL + "left join am_apunit c on a.kode_satuan=c.kodesatuan "
    SQL = SQL + "left join am_apitemmst d on a.kode_bahan = d.kodebarang and a.kode_satuan = d.kodeSatuan "
    SQL = SQL + "where a.nolot='" & txtnolot & "' and a.flag_tambahan ='0' order by a.line "
    Set RST = OBJ.Execute(SQL)
    hapusgrid1
    grid1.Row = 1

    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 1) = RST!kode_bahan
        grid1.TextMatrix(grid1.Row, 2) = RST!namabahan
        grid1.TextMatrix(grid1.Row, 3) = RST!Lot_bahan
        grid1.TextMatrix(grid1.Row, 4) = Format(RST!qty_bahan, "##,###,###,##0.000")
        grid1.TextMatrix(grid1.Row, 5) = RST!KODE_SATUAN
        grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
        grid1.TextMatrix(grid1.Row, 7) = Format(RST!hpp, "##,###,###,##0.00")
        grid1.TextMatrix(grid1.Row, 8) = RST!Line
        grid1.Col = 0
        Set grid1.CellPicture = uncheck
        setAlternatingGrid1 grid1.Row
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    
    
    txttotalproduksi.Value = GetTotalProduksi(txtnolot)

    'CARI TOTAL HASIL PRODUKSI
    SQL = "select SUM(qty_bahan)'total_hasil' from list_produksi_hasil where nolot='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        totalhasilproduksi = 0
    Else
        totalhasilproduksi = RST!total_hasil
    End If
    
    txttotalhasilproduksi.Value = totalhasilproduksi
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub totalg1()
On Error Resume Next
'TOTAL GRID1
    grid1.Row = 1
    tg1 = 0
    Do While True
        DoEvents
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            tg1 = CDbl(Format(grid1.TextMatrix(grid1.Row, 4), "general number") + CDbl(tg1))
                grid1.Row = grid1.Row + 1
    Loop
        tg1 = Format(tg1, "##,###,##0.0000")
End Sub

Private Sub totalhpp()
On Error Resume Next
'TOTAL GRID1
    grid1.Row = 1
    hpp = 0
    Do While True
        DoEvents
        If grid1.TextMatrix(grid1.Row, 7) = "" Then Exit Do
            hpp = CDbl(Format(grid1.TextMatrix(grid1.Row, 7), "general number") + CDbl(hpp))
                grid1.Row = grid1.Row + 1
    Loop
        hpp = Format(hpp, "##,###,##0.00")
        txttotalHPP = hpp
End Sub

Private Function setAlternatingGrid1(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If (i Mod 2) = 0 Then
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HFFFFFF
        Next
    End If
End Function

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
    setGrid1
End Sub
Private Sub initGrid1()
    With grid1
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "BAHAN"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "SATUAN"
        .TextMatrix(0, 7) = "HPP"
        .TextMatrix(0, 8) = "URUT"
    End With
End Sub
Private Sub setGrid1()
    With grid1
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1200
        .ColWidth(5) = 0
        .ColWidth(6) = 750
        If akses = True Then
            .ColWidth(7) = 2000
        Else
            .ColWidth(7) = 0
        End If
        .ColWidth(8) = 750
    End With
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

Private Sub cmdunlock_Click()
    frmunlocksop.Show vbModal
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
    If nmuser = "martsanto" Or nmuser = "Creator" Or nmuser = "angelgunawan" Or nmuser = "Angle" _
    Or nmuser = "kimlie" Or nmuser = "putri" Or nmuser = "HADY" Or nmuser = "bina" Then
        cmdunlock.Visible = True
    End If
    OBJ.Close
    
    initGrid1
    setGrid1
    cmbqc.AddItem "Lulus"
    cmbqc.AddItem "Tidak"
    cmbqc.text = "Lulus"
    If nmuser = "Creator" Then chkedit.Visible = True
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtnolot = "" Then Exit Sub
    
    poscol = grid1.Col
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 0:
                If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                If grid1.CellPicture = uncheck Then
                Set grid1.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid1.CellPicture = uncheck
                    hapusrow1
                    Exit Sub
                End If
                Set grid1.CellPicture = uncheck
                End If
        Case 1:
                If MsgBox("Ganti Bahan baku : " & grid1.TextMatrix(grid1.Row, 2) & vbCrLf & "Klik OK untuk melanjutkan.", vbQuestion + vbOKCancel, "Question") = vbCancel Then Exit Sub
                
                carisql1 = "select kodebarang, namabarang from am_apitemmst"
                namatabel = "Bahan Baku"
                frmsearch.Show vbModal
                If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
        Case 3:
                If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
                If txtnolot = "" Then
                    MsgBox "Nomor Lot harus diisi", vbCritical, AppName
                    Exit Sub
                End If
                If grid1.TextMatrix(grid1.Row, 3) <> "" Then
                    'If nmuser = "Creator" Then GoTo editmode:
                    'Exit Sub
                    'edit mode
'editmode:
                    kdlot = grid1.TextMatrix(grid1.Row, 3)
                End If
                lotbahan = grid1.TextMatrix(grid1.Row, 1)
                lotbahan1 = grid1.TextMatrix(grid1.Row, 2)
                lotbahan2 = grid1.TextMatrix(grid1.Row, 4)
                lotbahan3 = grid1.TextMatrix(grid1.Row, 3)
                If grid1.TextMatrix(grid1.Row, 3) <> "" Then grid1.TextMatrix(grid1.Row, 3) = ""
                statuslot = True
                frmaddlot.Show vbModal

        Case 4:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
                txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
                txtnilai.Left = grid1.Left + grid1.CellLeft
                txtnilai.Top = grid1.Top + grid1.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
        Case 6, 7:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 7) <> "0.00" Then Exit Sub
            'Input harga manual
                txthppmanual = "0"
                txtqty = "0"
                txtrow = "0"
                Picmanual.Visible = True
                lblitem = grid1.TextMatrix(grid1.Row, 2)
                txtqty = grid1.TextMatrix(grid1.Row, 4)
                txtrow = grid1.Row
                txthppmanual.SetFocus
        Case 8:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                txtnourut.Width = grid1.ColWidth(grid1.Col) - 40
                txtnourut = grid1.TextMatrix(grid1.Row, grid1.Col)
                txtnourut.Left = grid1.Left + grid1.CellLeft
                txtnourut.Top = grid1.Top + grid1.CellTop + 20
                txtnourut.Visible = True
                txtnourut.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid1.Col
        Case 1:
            With grid1
                .TextMatrix(.Row, 1) = hasil
                .TextMatrix(.Row, 2) = hasil1
                .TextMatrix(.Row, 4) = "0.000"
                .TextMatrix(.Row, 5) = "002"
                .TextMatrix(.Row, 6) = "Kg"
                .TextMatrix(.Row, 7) = "0.00"
                .Col = 0
                Set .CellPicture = uncheck
                'setAlternatingGrid1 grid1.Row
                '.Rows = .Rows + 1
                hasil = ""
                hasil1 = ""
                carisql1 = ""
                namatabel = ""
            End With
            
        Case 3:
            statuslot = False
            If hasil = "" Then
                'MsgBox kdlot
                grid1.TextMatrix(grid1.Row, 3) = kdlot
            Else
                grid1.TextMatrix(grid1.Row, 3) = hasil
                grid1.TextMatrix(grid1.Row, 7) = hasil1
            End If
            kdlot = ""
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub txthppmanual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.TextMatrix(txtrow, 7) = Format(CDbl(txthppmanual * txtqty), "##,###,###,##0.00")
        Picmanual.Visible = False
        
        Call totalg1
        Call totalhpp
        txttotalproduksi = tg1
        grid1.SetFocus
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim konvToKg As Double
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, 4) = txtnilai.text
        'cek konversi satuan bahan ke kilogram
        OBJ.Open dsn
            SQL = "Select * from am_apunit_konversi Where kdbrg ='" & grid1.TextMatrix(grid1.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            If Not RST.EOF Then
                konvToKg = txtnilai.text / RST!nilai
            Else
                konvToKg = txtnilai.Value
            End If
        OBJ.Close
        'If stokbahan = "0" Then
            'grid1.TextMatrix(grid1.Row, 7) = "0.00"
            'MsgBox "Stok " & grid1.TextMatrix(grid1.Row, 1) & " = 0.00", vbExclamation, AppName
        'Else
            grid1.TextMatrix(grid1.Row, 7) = Format(getHPP(grid1.TextMatrix(grid1.Row, 1), stokbahan, konvToKg), "##,####,###,##0.00")
        'End If
        'txttotalproduksi = txttotalproduksi + txtnilai
        Call totalg1
        Call totalhpp
        txttotalproduksi = tg1
        grid1.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnourut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, 8) = txtnourut.text
        grid1.SetFocus
    End If
End Sub

Private Sub txtnourut_LostFocus()
    txtnourut.Visible = False
End Sub
