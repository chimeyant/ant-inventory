VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmoutran 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmoutran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdCN 
      Height          =   285
      Left            =   3180
      TabIndex        =   74
      Top             =   2760
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No.Credit Note"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
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
      Left            =   3960
      TabIndex        =   73
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
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
      Left            =   6840
      TabIndex        =   72
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
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
      Left            =   5880
      TabIndex        =   71
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "View"
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
      TabIndex        =   70
      Top             =   7800
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   4200
      TabIndex        =   67
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Print PPn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   66
      Top             =   4560
      Width           =   1335
   End
   Begin Chameleon.chameleonButton cmdSaveP 
      Height          =   495
      Left            =   4920
      TabIndex        =   59
      Top             =   8520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Save + Print"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmoutran.frx":29A6
      Caption         =   "frmoutran.frx":29C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":2A32
      Keys            =   "frmoutran.frx":2A50
      Spin            =   "frmoutran.frx":2A92
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
   Begin XtremeSuiteControls.RadioButton optnonpajak 
      Height          =   255
      Left            =   2880
      TabIndex        =   53
      Top             =   4320
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Non Pajak"
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
   Begin XtremeSuiteControls.RadioButton optpajak 
      Height          =   270
      Left            =   1920
      TabIndex        =   52
      Top             =   4320
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   "Pajak"
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
   Begin VB.Frame Frame1 
      Height          =   2380
      Left            =   5640
      TabIndex        =   48
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ListBox List1 
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
         Height          =   2370
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   1815
      End
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "frmoutran.frx":2ABA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":2B26
      Key             =   "frmoutran.frx":2B44
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
      MaxLength       =   60
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
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      Picture         =   "frmoutran.frx":2B80
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      Picture         =   "frmoutran.frx":2ECE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1590
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2805
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
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
      _Band(0).Cols   =   5
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   143589379
      CurrentDate     =   37694
   End
   Begin TDBText6Ctl.TDBText txtkodecomp 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":31B0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":321C
      Key             =   "frmoutran.frx":323A
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   4
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
   Begin TDBText6Ctl.TDBText txtkodetran 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":3276
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":32E2
      Key             =   "frmoutran.frx":3300
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   2
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
   Begin TDBText6Ctl.TDBText txtnotran 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":333C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":33A8
      Key             =   "frmoutran.frx":33C6
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   15
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
   Begin TDBText6Ctl.TDBText txtkodecur 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   3480
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":3402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":346E
      Key             =   "frmoutran.frx":348C
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   4
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
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Top             =   5640
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":34C8
      Caption         =   "frmoutran.frx":34E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":3554
      Keys            =   "frmoutran.frx":3572
      Spin            =   "frmoutran.frx":35B4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
   Begin Chameleon.chameleonButton cmdclose1 
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   8730
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Close"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":35DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear1 
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   8760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Clear"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":38F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtkredit 
      Height          =   285
      Left            =   5280
      TabIndex        =   32
      Top             =   480
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":3C10
      Caption         =   "frmoutran.frx":3C30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":3C9C
      Keys            =   "frmoutran.frx":3CBA
      Spin            =   "frmoutran.frx":3CFC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
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
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtdebet 
      Height          =   285
      Left            =   5280
      TabIndex        =   33
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":3D24
      Caption         =   "frmoutran.frx":3D44
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":3DB0
      Keys            =   "frmoutran.frx":3DCE
      Spin            =   "frmoutran.frx":3E10
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
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
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   38
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Currency"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":3E38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   360
      TabIndex        =   39
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Company"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":4152
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtcash 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":446C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":44D8
      Key             =   "frmoutran.frx":44F6
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   15
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
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   360
      TabIndex        =   40
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Acc. Cash/Bank"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":4532
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtncash 
      Height          =   285
      Left            =   5880
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":484C
      Caption         =   "frmoutran.frx":486C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":48D8
      Keys            =   "frmoutran.frx":48F6
      Spin            =   "frmoutran.frx":4938
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText txtketcash 
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Top             =   4920
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":4960
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":49CC
      Key             =   "frmoutran.frx":49EA
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   60
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
   Begin TDBText6Ctl.TDBText txtcekbg 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   5280
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":4A26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":4A92
      Key             =   "frmoutran.frx":4AB0
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   20
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
   Begin TDBText6Ctl.TDBText txtnovoucher 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":4AEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":4B58
      Key             =   "frmoutran.frx":4B76
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   15
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
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   285
      Left            =   360
      TabIndex        =   51
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No. Voucher"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":4BB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtnobp 
      Height          =   285
      Left            =   1800
      TabIndex        =   54
      Top             =   4920
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":4ECC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":4F38
      Key             =   "frmoutran.frx":4F56
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   7
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
   Begin TDBText6Ctl.TDBText txtkpd 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":4F92
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":4FFE
      Key             =   "frmoutran.frx":501C
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   100
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1800
      TabIndex        =   57
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   143589379
      CurrentDate     =   37694
   End
   Begin Chameleon.chameleonButton cmdadd1 
      Height          =   495
      Left            =   5040
      TabIndex        =   58
      Top             =   8760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "View"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmoutran.frx":5058
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtpembulatan 
      Height          =   285
      Left            =   5280
      TabIndex        =   60
      Top             =   820
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":5372
      Caption         =   "frmoutran.frx":5392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":53FE
      Keys            =   "frmoutran.frx":541C
      Spin            =   "frmoutran.frx":545E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
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
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   285
      Left            =   5880
      TabIndex        =   62
      Top             =   4200
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":5486
      Caption         =   "frmoutran.frx":54A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":5512
      Keys            =   "frmoutran.frx":5530
      Spin            =   "frmoutran.frx":5572
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs2 
      Height          =   285
      Left            =   5880
      TabIndex        =   65
      Top             =   3480
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran.frx":559A
      Caption         =   "frmoutran.frx":55BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":5626
      Keys            =   "frmoutran.frx":5644
      Spin            =   "frmoutran.frx":5686
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
   Begin TDBText6Ctl.TDBText txtnobukti 
      Height          =   285
      Left            =   3480
      TabIndex        =   68
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":56AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":571A
      Key             =   "frmoutran.frx":5738
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   15
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
   Begin TDBText6Ctl.TDBText txtkdsupp 
      Height          =   285
      Left            =   6600
      TabIndex        =   69
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Caption         =   "frmoutran.frx":5774
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran.frx":57E0
      Key             =   "frmoutran.frx":57FE
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   15
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
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   1800
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      X1              =   240
      X2              =   7680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Nilai Kurs Bayar"
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
      Left            =   4440
      TabIndex        =   64
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total IDR"
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
      Left            =   4320
      TabIndex        =   63
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "PPN"
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
      Left            =   4680
      TabIndex        =   61
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      Caption         =   "Tanggal JT"
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
      Left            =   360
      TabIndex        =   56
      Top             =   5670
      Width           =   1335
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      Caption         =   "Di Bayar Kepada"
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
      Left            =   360
      TabIndex        =   55
      Top             =   3165
      Width           =   1335
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      Caption         =   "zz=<kode bank>"
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
      Height          =   180
      Left            =   6120
      TabIndex        =   50
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      Caption         =   "(manual) BK=Bank Keluar (YYMM/zz/XXXXX)"
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
      Left            =   3720
      TabIndex        =   47
      Top             =   2310
      Width           =   3255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Kode/No Transaksi"
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
      Left            =   360
      TabIndex        =   46
      Top             =   2310
      Width           =   1455
   End
   Begin MSForms.ComboBox cmbdaerah 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   855
      VariousPropertyBits=   746608667
      MaxLength       =   1
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1508;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   3
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Kode Daerah"
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
      Left            =   360
      TabIndex        =   45
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      Caption         =   "No. Cheque/BG"
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
      Left            =   360
      TabIndex        =   44
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Desc. Cash/Bank"
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
      Left            =   360
      TabIndex        =   43
      Top             =   4950
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Nilai Cash/Bank"
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
      Left            =   4440
      TabIndex        =   42
      Top             =   3870
      Width           =   1335
   End
   Begin VB.Label lblcash 
      Appearance      =   0  'Flat
      Caption         =   "Cash/Bank : "
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
      Left            =   240
      TabIndex        =   41
      Top             =   7665
      Width           =   4215
   End
   Begin VB.Label lblcomkredit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lines"
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
      Left            =   6840
      TabIndex        =   37
      Top             =   510
      Width           =   855
   End
   Begin VB.Label lblcomdebet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lines"
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
      Left            =   6840
      TabIndex        =   36
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Kredit"
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
      Left            =   4320
      TabIndex        =   35
      Top             =   510
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Debet"
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
      Left            =   4320
      TabIndex        =   34
      Top             =   150
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Bank Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblnamacc 
      Appearance      =   0  'Flat
      Caption         =   "Nama Account :"
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
      Left            =   240
      TabIndex        =   23
      Top             =   7905
      Width           =   4455
   End
   Begin VB.Label lblnamacur 
      Appearance      =   0  'Flat
      Caption         =   "Currency :"
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
      Left            =   240
      TabIndex        =   22
      Top             =   8145
      Width           =   4455
   End
   Begin VB.Label lblnamacomp 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   21
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Nilai Kurs Pajak"
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
      Left            =   4680
      TabIndex        =   19
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      Caption         =   "Tanggal Transaksi"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000011&
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmoutran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Private poscol As Integer
Private posrow1 As Integer
Dim posrow, str2, str3, str4, str6, str7, str8, str9 As String

Private Sub carisisa()
    str4 = 0
    Select Case grid.Col
        Case 3
            str4 = txtkredit - txtdebet + Val(Format(grid.TextMatrix(grid.Row, grid.Col), "general number"))
            If str4 <= 0 Then
                txtnilai = Val(Format(grid.TextMatrix(grid.Row, grid.Col), "general number"))
            Else
                txtnilai = str4
            End If
        Case 4
            str4 = txtdebet - txtkredit + Val(Format(grid.TextMatrix(grid.Row, grid.Col), "general number"))
            If str4 <= 0 Then
                txtnilai = Val(Format(grid.TextMatrix(grid.Row, grid.Col), "general number"))
            Else
                txtnilai = str4
            End If
    End Select
End Sub

Private Sub ViewReport()
    Dim nilai_rupiah As Double
    Dim nilai_pnn As Double
    Dim nilai_hutang As Double
    Dim isppn As Double
    
    SQL = "select sum(jumlah) as jml from am_voucherin where novoucher='" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    
    nilai_rupiah = SpyRound(RST!jml * txtnilaikurs)
    If txtppn.Value > 0 Then
        nilai_pnn = SpyRound((RST!jml * txtnilaikurs2) * (txtppn.Value / 100))
    End If
    
    nilai_hutang = nilai_rupiah + nilai_pnn
    'ini yg baru (dicek lagi kalau tidak ada yg dirubah selain yg dibawah 1 baris ini hapus aja)
    isppn = SpyRound((RST!jml * txtnilaikurs2) + nilai_pnn)
    
    SQL = "Select a.*,b.tgl,b.nilai,b.ppn,(a.jumlah * nilai) as jml From am_voucherin a inner join am_voucherhdr b "
    SQL = SQL + "On a.novoucher=b.novoucher Where a.novoucher='" + txtnovoucher + "'"
    
    With rptNONBB
        .Field17 = txtkpd
        .Field18 = txtcekbg
        .Field22 = Format(date2, "dd/MM/yyyy")
        .Field19 = txtnobp
        .Field20 = txtnovoucher
        .Field21 = Format(date1, "dd/MM/yyyy")
        .Field26 = txtketcash
        .Field27 = Format(nilai_rupiah, "###,###,##0.00")
        .Field31 = txtcash
        .lblkurs = txtkodecur
        .lblnilaikurs = txtnilaikurs.text
        .lblppn = Format(nilai_pnn, "###,###,##0.00")
        .lbljumlah = Format(nilai_rupiah, "###,###,##0.00")
        If txtkodecur = "IDR" Then
            .lblhutang = Format(nilai_hutang, "###,###,##0.00")
        Else
            'ini yg baru
            .lblhutang = Format(isppn, "###,###,##0.00")
        End If
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub

Private Sub cmbdaerah_LostFocus()
    If Not (cmbdaerah >= 1 And cmbdaerah <= 4) Then
        cmbdaerah = ""
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        cmbdaerah.SetFocus
    Else
        cari_out
    End If
End Sub

Private Sub cmdadd_Click()
On Error GoTo err_handler:
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Silahkan menunggu beberapa saat sedang ada proses posting data " & vbCrLf & _
            "Computer name : " & RST!comp_id & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from gl_accrl"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str7 = RST!rl_ptd
        str8 = RST!rl_ytd
    Else
        str7 = ""
        str8 = ""
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from gl_masterac where typeac = 'IS'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str9 = RST!noac
    Else
        str9 = ""
    End If
    OBJ.Close
    
    If txtdebet <> txtkredit Then
        If MsgBox("Transaction Is Unbalance, continue to Add ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    End If
    
    If txtkodecomp = "" Or txtkodetran = "" Or txtnotran = "" Or txtkodecur = "" Or txtcash = "" Or cmbdaerah = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    If txtnobp = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    If txtnilaikurs = 0 Then
        MsgBox "Nilai Kurs = 0.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    If Len(Trim(txtkodetran)) = 0 Or Len(Trim(txtnotran)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkodetran <> "BK" Then
        OBJ.Open dsn
        SQL = "select * from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx = '" & txtnotran & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Can't Add, " & txtkodetran & txtnotran & " , Transaction Already Exist." & vbCrLf & "click OK to add with next number.", vbOKCancel + vbQuestion, "Information") = vbOK Then
    
                SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx like '" & Format(date1, "YYMM") & "/" & cmbdaerah & "/%' and flagprint='O' order by notrx desc"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    If Len(RST!notrx + 1) = 5 Then
                        txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/" & RST!notrx + 1
                    ElseIf Len(RST!notrx + 1) = 4 Then
                        txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0" & RST!notrx + 1
                    ElseIf Len(RST!notrx + 1) = 3 Then
                        txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00" & RST!notrx + 1
                    ElseIf Len(RST!notrx + 1) = 2 Then
                        txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/000" & RST!notrx + 1
                    ElseIf Len(RST!notrx + 1) = 1 Then
                        txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0000" & RST!notrx + 1
                    End If
                Else
                    txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00001"
                End If
                
            Else
                OBJ.Close
                cmdclear_Click
                Exit Sub
            End If
        End If
        OBJ.Close
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" And grid.TextMatrix(grid.Row, 4) = "0.00" Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        If str7 <> "" And x_original(grid.TextMatrix(grid.Row, 1)) = str7 Then
            MsgBox "Account PTD not allowed On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If str8 <> "" And x_original(grid.TextMatrix(grid.Row, 1)) = str8 Then
            MsgBox "Account YTD not allowed On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If str9 <> "" And x_original(grid.TextMatrix(grid.Row, 1)) = str9 Then
            MsgBox "Account Income Summary not allowed On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop

If Check1 = Checked And Check2 = Checked Then GoTo JustPPN:
If Check2 = Checked And Check1 = Unchecked Then GoTo JustBahanBaku:

    grid.Row = 1
    OBJ.Open dsn
    SQL = "insert into gl_transaksi "
    SQL = SQL + "(kdcomp, "
    SQL = SQL + "tgltrx, "
    SQL = SQL + "kdtrx, "
    SQL = SQL + "notrx, "
    SQL = SQL + "kurs, "
    SQL = SQL + "noactrx, "
    SQL = SQL + "desctrx, "
    SQL = SQL + "dbkrtrx, "
    SQL = SQL + "amounttrx, "
    SQL = SQL + "nilaitrx, "
    SQL = SQL + "currtrx, "
    SQL = SQL + "flag, "
    SQL = SQL + "flagprint, "
    SQL = SQL + "flagadjustment, "
    SQL = SQL + "cekbg, "
    SQL = SQL + "identry, "
    SQL = SQL + "idupdate, "
    SQL = SQL + "dateentry, "
    SQL = SQL + "dateupdate, "
    SQL = SQL + "lineitem)"
    
    SQL = SQL + " values"
    SQL = SQL + "('" & txtkodecomp & "',"
    SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
    SQL = SQL + "'" & txtkodetran & "',"
    SQL = SQL + "'" & txtnotran & "',"
    SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
    SQL = SQL + "'" & x_original(txtcash) & "',"
    SQL = SQL + "'" & txtnobp + " " + txtketcash & "',"
    SQL = SQL + "'K',"
    If txtppn.Value = "11" Then
        SQL = SQL + "Ceiling(convert(money,'" & (txtncash * txtnilaikurs2) + (txtpembulatan * (txtppn.Value / 100)) & "')),"
        SQL = SQL + "Ceiling(convert(money,'" & txtncash + (txtncash * (txtppn.Value / 100)) & "')),"
    Else
        SQL = SQL + "Floor(convert(money,'" & txtncash * txtnilaikurs & "')),"
        SQL = SQL + "convert(money,'" & txtncash & "'),"
    End If
    SQL = SQL + "'" & txtkodecur & "',"
    SQL = SQL + "'" & lblbal & "',"
    SQL = SQL + "'O',"
    SQL = SQL + "'" & cmbdaerah & "',"
    SQL = SQL + "'" & txtcekbg & "',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "'0',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "convert(numeric,'1'))"
    Set RST = OBJ.Execute(SQL)

    Do While True
        If grid.TextMatrix(grid.Row, 2) = "" Then Exit Do

        SQL = "insert into gl_transaksi "
        SQL = SQL + "(kdcomp, "
        SQL = SQL + "tgltrx, "
        SQL = SQL + "kdtrx, "
        SQL = SQL + "notrx, "
        SQL = SQL + "kurs, "
        SQL = SQL + "noactrx, "
        SQL = SQL + "desctrx, "
        SQL = SQL + "dbkrtrx, "
        SQL = SQL + "amounttrx, "
        SQL = SQL + "nilaitrx, "
        SQL = SQL + "currtrx, "
        SQL = SQL + "flag, "
        SQL = SQL + "flagprint, "
        SQL = SQL + "flagadjustment, "
        SQL = SQL + "cekbg, "
        SQL = SQL + "identry, "
        SQL = SQL + "idupdate, "
        SQL = SQL + "dateentry, "
        SQL = SQL + "dateupdate, "
        SQL = SQL + "lineitem)"
        
        SQL = SQL + " values"
        SQL = SQL + "('" & txtkodecomp & "',"
        SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
        SQL = SQL + "'" & txtkodetran & "',"
        SQL = SQL + "'" & txtnotran & "',"
        SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
        SQL = SQL + "'" & x_original(grid.TextMatrix(grid.Row, 1)) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'D',"
        SQL = SQL + "Floor(convert(money,'" & (Format(grid.TextMatrix(grid.Row, 3), "general number") * txtnilaikurs) & "')),"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & txtkodecur & "',"
        SQL = SQL + "'" & lblbal & "',"
        SQL = SQL + "'O',"
        SQL = SQL + "'" & cmbdaerah & "',"
        SQL = SQL + "'" & txtcekbg & "',"
        SQL = SQL + "'" & kuser & "',"
        SQL = SQL + "'0',"
        SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL = SQL + "convert(numeric,'" & grid.Row + 1 & "'))"
           
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
    Loop
'SIMPAN PPN NON BAHAN BAKU
        If txtppn.Value = "11" Then
            SQL = "insert into gl_transaksi "
            SQL = SQL + "(kdcomp, "
            SQL = SQL + "tgltrx, "
            SQL = SQL + "kdtrx, "
            SQL = SQL + "notrx, "
            SQL = SQL + "kurs, "
            SQL = SQL + "noactrx, "
            SQL = SQL + "desctrx, "
            SQL = SQL + "dbkrtrx, "
            SQL = SQL + "amounttrx, "
            SQL = SQL + "nilaitrx, "
            SQL = SQL + "currtrx, "
            SQL = SQL + "flag, "
            SQL = SQL + "flagprint, "
            SQL = SQL + "flagadjustment, "
            SQL = SQL + "cekbg, "
            SQL = SQL + "identry, "
            SQL = SQL + "idupdate, "
            SQL = SQL + "dateentry, "
            SQL = SQL + "dateupdate, "
            SQL = SQL + "lineitem)"
            
            SQL = SQL + " values"
            SQL = SQL + "('" & txtkodecomp & "',"
            SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
            SQL = SQL + "'" & txtkodetran & "',"
            SQL = SQL + "'" & txtnotran & "',"
            SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
            SQL = SQL + "'14103000',"
            SQL = SQL + "'" & "PPN" + " " + txtketcash & "',"
            SQL = SQL + "'D',"
            'SQL = SQL + "floor(convert(money,'" & txtncash * txtnilaikurs2 * 0.1 & "')),"
            'SQL = SQL + "floor(convert(money,'" & txtncash * 0.1 & "')),"
            SQL = SQL + "floor(convert(money,'" & txtncash * txtnilaikurs2 * (txtppn.Value / 100) & "')),"
            SQL = SQL + "floor(convert(money,'" & txtncash * (txtppn.Value / 100) & "')),"
            SQL = SQL + "'" & txtkodecur & "',"
            SQL = SQL + "'" & lblbal & "',"
            SQL = SQL + "'O',"
            SQL = SQL + "'" & cmbdaerah & "',"
            SQL = SQL + "'" & txtcekbg & "',"
            SQL = SQL + "'" & kuser & "',"
            SQL = SQL + "'0',"
            SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL = SQL + "convert(numeric,'" & grid.Row + 1 & "'))"
            Set RST = OBJ.Execute(SQL)
        End If
        
    SQL = "Select * From no_bank_payment Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
       
    With RST
        .AddNew
        !notrx = txtnotran
        !no_payment = txtnobp
        !no_voucher = txtnovoucher
        !kpd = txtkpd
        !tgljt = date2
        !ppn = txtppn
        If optpajak.Value = True Then
            !is_pajak = "1"
        ElseIf optnonpajak.Value = True Then
            !is_pajak = "0"
        End If
        !ref = "P"
        !flag = "0"
        .Update
    End With
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"

    ViewReport
    OBJ.Close
    cmdclear_Click
    Exit Sub
    
JustPPN:
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    SimpanPPN
    OBJ.Close
    cmdclear_Click
    Exit Sub
    
JustBahanBaku:
    If MsgBox("Lanjutkan ke Pelunasan Piutang ?", vbQuestion + vbYesNo, "Konfirmasi Pelunasan") = vbYes Then
        hasil = txtnobp
        hasil1 = txtkodecur
        hasil2 = txtkpd
        hasil3 = Format(date1, "yyyy/MM/dd")
        hasil4 = Format(txtnilaikurs, "###,###,##0.00")
        hasil5 = txtkdsupp
        hasil6 = txtnovoucher
        Set frmpayap = New frmpayap
        frmpayap.Show
    Else
        If MsgBox("Click YES untuk Print total + PPn, Click NO untuk print tanpa PPn", vbQuestion + vbYesNo, "Print Confirm") = vbYes Then
            PrintBahanBakuPPn
        Else
        '---------------
            PrintBahanBaku
        End If
        cmdclear_Click
    End If
    Exit Sub
err_handler:
    OBJ.Close
    MsgBox Err.Description
End Sub
Private Sub PrintBahanBakuPPn()
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_ppn As Long
    Dim nilai_potongan As Long
    Dim nilai_hutang As Long
    Dim total As Double
'====================================
'SIMPAN NO PAYMENT
    OBJ.Open dsn
    SQL = "Select * From no_bank_payment Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
       
    With RST
        .AddNew
        !notrx = txtnotran
        !no_payment = txtnobp
        !no_voucher = txtnovoucher
        !kpd = txtkpd
        !tgljt = date2
        !ppn = txtppn
        If optpajak.Value = True Then
            !is_pajak = "1"
        ElseIf optnonpajak.Value = True Then
            !is_pajak = "0"
        End If
        !ref = "P"
        !flag = "1"
        .Update
    End With
'=====================================
    
    SQL = "SELECT a.NoApply,a.nilaikurs,a.Amount,a.Selisih,a.potongan,(a.PPN * a.nilaikurs) AS nilaippn,a.kodecur, a.TransType, a.Amount - a.Potongan + a.PPN AS jumlah "
    SQL = SQL + "From am_apopnfil a inner join am_beliapp b on b.NoBeli = a.NoBeli "
    SQL = SQL + "Where b.ref1 = '" + txtnovoucher + "'"
    
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        kode_kurs = RST!kodecur
        nilai_kurs = RST!nilaikurs
        nilai_jumlah = RST!amount
        nilai_ppn = RST!nilaippn
        nilai_potongan = RST!potongan
        nilai_hutang = RST!jumlah + RST!nilaippn
        RST.MoveNext
    Loop
    'OBJ.Close
    SQL = "Select SUM(Qty * Price) as Jml  From am_beliapp Where Ref1 = '" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        total = Format(RST!jml, "###,##0.00")
    End If
    OBJ.Close
    
    SQL = "Select  a.*, b.namabarang ,d.namasatuan ,(SUM(a.qty) * a.price) as jumlah,c.noapply"
    SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang"
    SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    SQL = SQL + " Where a.ref1 = '" + txtnovoucher + "'"
    SQL = SQL + " GROUP BY a.NoBeli,a.TglBeli,a.NoPO,a.Ref1,a.Ref2,a.Kodesupp,a.Kodecur,a.Nilaikurs,"
    SQL = SQL + " a.KodeBarang,a.Qty,a.Price,a.KodeSatuan,a.ppn,a.keterangan,a.keterangan2,a.keterangan3,"
    SQL = SQL + " a.keterangan4,a.LineItem,a.flag1,a.flag2,b.NamaBarang,d.NamaSatuan,c.NoApply"
    
    'With rptBBPPn
    '    .Field17 = txtkpd
    '    .Field18 = txtcekbg
    '    .Field22 = date2
    '    .Field19 = txtnobp
    '    .Field20 = txtnovoucher
    '    .Field21 = date1
    '    .Field31 = txtcash
    '    .Field26 = txtketcash
    '    .Field27 = total
    '    .lbljumlah = Format(total, "###,###,##0.00")
    '    .lblppn = Format(nilai_ppn, "###,###,##0.00")
    '    .lblpotongan = Format(nilai_potongan, "###,###,##0.00")
    '    .lblhutang = Format((total + nilai_ppn - nilai_potongan), "###,###,##0.00")
    '    .lblkurs = txtkodecur
    '    .lblnilaikurs = Format(txtnilaikurs, "###,###,##0.00")
    '    .DataControl1.Source = SQL
    '    .DataControl1.ConnectionString = dsn
    '    .Show vbModal
    'End With
End Sub

Private Sub PrintBahanBaku()
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_ppn As Long
    Dim nilai_potongan As Long
    Dim nilai_hutang As Long
    Dim total As Double
'====================================
'SIMPAN NO PAYMENT
    OBJ.Open dsn
    SQL = "Select * From no_bank_payment Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
       
    With RST
        .AddNew
        !notrx = txtnotran
        !no_payment = txtnobp
        !no_voucher = txtnovoucher
        !kpd = txtkpd
        !tgljt = date2
        !ppn = txtppn
        If optpajak.Value = True Then
            !is_pajak = "1"
        ElseIf optnonpajak.Value = True Then
            !is_pajak = "0"
        End If
        !ref = "P"
        !flag = "1"
        .Update
    End With
'=====================================
    
    SQL = "SELECT a.NoApply,a.nilaikurs,a.Amount,a.Selisih,a.potongan,(a.PPN * a.nilaikurs) AS nilaippn,a.kodecur, a.TransType, a.Amount - a.Potongan + a.PPN AS jumlah "
    SQL = SQL + "From am_apopnfil a inner join am_beliapp b on b.NoBeli = a.NoBeli "
    SQL = SQL + "Where b.ref1 = '" + txtnovoucher + "'"
    
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        kode_kurs = RST!kodecur
        nilai_kurs = RST!nilaikurs
        nilai_jumlah = RST!amount
        nilai_ppn = RST!nilaippn
        nilai_potongan = RST!potongan
        nilai_hutang = RST!jumlah
        RST.MoveNext
    Loop
    'OBJ.Close
    SQL = "Select SUM(Qty * Price) as Jml  From am_beliapp Where Ref1 = '" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        total = Format(RST!jml, "###,##0.00")
    End If
    OBJ.Close
    
    SQL = "Select  a.*, b.namabarang ,d.namasatuan ,(SUM(a.qty) * a.price) as jumlah,c.noapply"
    SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang"
    SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    SQL = SQL + " Where a.ref1 = '" + txtnovoucher + "'"
    SQL = SQL + " GROUP BY a.NoBeli,a.TglBeli,a.NoPO,a.Ref1,a.Ref2,a.Kodesupp,a.Kodecur,a.Nilaikurs,"
    SQL = SQL + " a.KodeBarang,a.Qty,a.Price,a.KodeSatuan,a.ppn,a.keterangan,a.keterangan2,a.keterangan3,"
    SQL = SQL + " a.keterangan4,a.LineItem,a.flag1,a.flag2,b.NamaBarang,d.NamaSatuan,c.NoApply"
    
   ' With rptBB
   '     .Field17 = txtkpd
   '     .Field18 = txtcekbg
   '     .Field22 = date2
   '     .Field19 = txtnobp
   '     .Field20 = txtnovoucher
   '     .Field21 = date1
   '     .Field31 = txtcash
   '     .Field26 = txtketcash
   '     .Field32 = total
   '     .Field23 = total
   '     .Field27 = total
   '     .Field28 = total
   '     .lblkurs = txtkodecur
   '     .lblnilaikurs = Format(txtnilaikurs, "###,###,##0.00")
   '     .DataControl1.Source = SQL
   '     .DataControl1.ConnectionString = dsn
   '     .Show vbModal
   ' End With
End Sub
Private Sub SimpanBahanBaku()
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_ppn As Long
    Dim nilai_potongan As Long
    Dim nilai_hutang As Long
    
    OBJ.Open dsn
    SQL = "SELECT NoApply,nilaikurs,Amount,Selisih,potongan,(PPN * nilaikurs) AS nilaippn,kodecur, TransType, Amount - Potongan + PPN AS jumlah"
    SQL = SQL + " From am_apopnfil"
    SQL = SQL + " Where NoBeli='" + txtnobukti + "'"
    
    OBJ1.Open dsn
    Set RST = OBJ1.Execute(SQL)
    Do While Not RST.EOF
        kode_kurs = RST!kodecur
        nilai_kurs = RST!nilaikurs
        nilai_jumlah = RST!amount
        nilai_ppn = RST!nilaippn
        nilai_potongan = RST!potongan
        nilai_hutang = RST!jumlah
        RST.MoveNext
    Loop
    OBJ1.Close
    
    SQL = "Select  a.* ,b.namabarang,a.qty , d.namasatuan ,(a.qty * a.price) as jumlah,c.noapply,e.noac,e.nmac "
    SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang "
    SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    SQL = SQL + " inner join gl_masterac e on a.kodesupp=e.jenisac10"
    SQL = SQL + " Where a.ref1 = '" + txtnovoucher + "'"
    
    'With rptBB
    '    .Field17 = txtkpd
    '    .Field18 = txtcekbg
    '    .Field22 = date2
    '    .Field19 = txtnobp
    '    .Field20 = txtnovoucher
    '    .Field21 = date1
    '    .Field31 = txtcash
    '    .lblkurs = kode_kurs
        '.lblnilaikurs = Format(nilai_kurs, "###,###,##0.00")
    '    .lblnilaikurs = Format(txtnilaikurs, "###,###,##0.00")
    '    .DataControl1.Source = SQL
    '    .DataControl1.ConnectionString = dsn
    '    .Show vbModal
    'End With
    'ViewHarga
End Sub
Private Sub SimpanPPN()
    Dim nilai_rupiah As Double
    Dim total As Double

    OBJ.Open dsn
    SQL = "Select * From no_bank_payment Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
       
    With RST
        .AddNew
        !notrx = txtnotran
        !no_payment = txtnobp
        !no_voucher = txtnovoucher
        !kpd = txtkpd
        !tgljt = date2
        !ppn = txtppn
        If optpajak.Value = True Then
            !is_pajak = "1"
        ElseIf optnonpajak.Value = True Then
            !is_pajak = "0"
        End If
        !ref = "P"
        !flag = "1"
        .Update
    End With
    
    SQL = "Select SUM(floor(Qty * Price * nilaikurs * (txtppn.Value / 100))) as Jml  From am_beliapp Where Ref1 = '" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        total = Format(RST!jml, "###,##0.00")
    End If
    'OBJ.Close
    
    SQL = "Select  a.* ,b.namabarang,a.qty , d.namasatuan ,ceiling(a.qty * a.price * a.nilaikurs * 0.11) as jumlah,c.noapply "
    SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang "
    SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    'SQL = SQL + " inner join gl_masterac e on a.kodesupp=e.jenisac10"
    SQL = SQL + " Where a.ref1 = '" + txtnovoucher + "'"
    
    'With rpt_ppn
    '    .Field17 = txtkpd
    '    .Field18 = txtcekbg
    '    .Field22 = date2
    '    .Field19 = txtnobp
    '    .Field20 = txtnovoucher
    '    .Field21 = date1
    '    .Field26 = txtketcash
    '    .Field34 = txtcash
    '    .Field27 = Format(total, "###,##0.00")
    '    .Field41 = Format(total, "###,##0.00")
    '    .Field42 = Format(total, "###,##0.00")
    '    .fperkiraan = "14103000"
    '    .lblkurs = txtkodecur
    '    .lblnilaikurs = Format(txtnilaikurs2, "###,###,##0.00")
    '    .DataControl1.Source = SQL
    '    .DataControl1.ConnectionString = dsn
    '    .Show vbModal
    'End With
End Sub
Private Sub SimpanPPNnonBB()
    Dim nilai_rupiah As Double
    Dim nilai_pnn As Double
    Dim nilai_hutang As Double
    
    OBJ.Open dsn
    SQL = "select sum(jumlah) as jml from am_voucherin where novoucher='" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    
    nilai_rupiah = RST!jml
    If txtppn.Value > 0 Then
        nilai_pnn = nilai_rupiah * (txtppn.Value / 100)
    End If
    OBJ.Close
    nilai_hutang = nilai_rupiah + nilai_pnn

    SQL = "Select a.*,b.tgl,b.nilai,b.ppn,(a.jumlah * 0.11) as jml From am_voucherin a inner join am_voucherhdr b "
    SQL = SQL + "On a.novoucher=b.novoucher Where a.novoucher='" + txtnovoucher + "'"
    
    With rptNONBB
        .Field17 = txtkpd
        .Field18 = txtcekbg
        .Field22 = Format(date2, "dd/MM/yyyy")
        .Field19 = txtnobp
        .Field20 = txtnovoucher
        .Field21 = Format(date1, "dd/MM/yyyy")
        .Field26 = txtketcash
        .Field27 = Format(nilai_rupiah * (txtppn.Value / 100), "###,###,##0.00")
        .Field33 = Format(nilai_rupiah * (txtppn.Value / 100), "###,###,##0.00")
        .Field31 = "14103000" 'txtcash
        .lblkurs = txtkodecur
        .lblnilaikurs = txtnilaikurs.text
        .lbljumlah.Visible = False
        .lblppn.Visible = False
        .lblhutang.Visible = False
        '.lblppn = Format(nilai_pnn, "###,###,##0.00")
        '.lbljumlah = Format(nilai_rupiah, "###,###,##0.00")
        '.lblhutang = Format(nilai_hutang, "###,###,##0.00")
        .Label26.Visible = False
        .Label29.Visible = False
        .Label30.Visible = False
        .Label31.Visible = False
        .Field32.Visible = True
        .Field33.Visible = True
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub

Private Sub cmdclear_Click()
    hapusemua
    txtkodecomp = ""
    lblnamacomp = ""
    cmbdaerah = ""
    txtcash = ""
    lblcash = "Cash/Bank : "
    date1.Value = Date
    date2.Value = Date
    txtkodetran = ""
    txtnotran = ""
    txtkpd = ""
    txtnobp = ""
    txtnovoucher = ""
    optpajak.Value = False
    optnonpajak.Value = False
    txtkodecomp.SetFocus
    Check1.Visible = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCN_Click()
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub cmddelete_Click()
    frmDelcbo.Show vbModal
End Sub

Private Sub cmdSaveP_Click()
OBJ.Open dsn
If txtkodecur.text = "IDR" Then
    ViewReport
Else
    If Check1 = Checked Then
        SimpanPPN
    ElseIf Check2 = Checked Then
        SimpanBahanBaku
    End If
End If
OBJ.Close
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecomp = hasil
    lblnamacomp = hasil1
    txtkodecomp_LostFocus
    hasil = ""
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecur = hasil
    carikurs
    hasil = ""
End Sub

Private Sub cmdsearch2_Click()
    'namatabel = "Cash/Bank"
    'setup1 = txtkodecomp
    'setup2 = txtkodecomp
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtcash = hasil
    carinoac
    hasil = ""
End Sub

Private Sub date1_LostFocus()
    cari_out
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then Frame1.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then Frame1.Visible = False
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Account"
    grid.TextMatrix(0, 2) = "Keterangan"
    grid.TextMatrix(0, 3) = "Debet"
    grid.TextMatrix(0, 4) = "IDR"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 4350
    grid.ColWidth(3) = 1650
    grid.ColWidth(4) = 0
    
    grid.RowHeightMin = 300
    date1.Value = Date
    date2.Value = Date
    cmbdaerah.Clear
    cmbdaerah.ColumnCount = 2
    cmbdaerah.ListWidth = "4 cm"
    cmbdaerah.ColumnWidths = "1 cm; 3 cm"
    cmbdaerah.AddItem "1"
    cmbdaerah.AddItem "2"
    cmbdaerah.AddItem "3"
    cmbdaerah.AddItem "4"
    cmbdaerah.List(0, 1) = "Pabrik"
    cmbdaerah.List(1, 1) = "Jakarta"
    cmbdaerah.List(2, 1) = "Surabaya"
    cmbdaerah.List(3, 1) = "Semarang"
    List1.Clear
    
    List1.AddItem "01-BCA SPARTA IDR"
    List1.AddItem "02-BCA SPARTA USD"
    List1.AddItem "03-BCA WFN IDR"
    List1.AddItem "04-BCA WFN USD"
    List1.AddItem "05-BCA SURABAYA IDR"
    List1.AddItem "06-BNI SPARTA IDR"
    List1.AddItem "07-BNI SPARTA USD"
    List1.AddItem "08-BDI SPARTA IDR"
    List1.AddItem "09-CIMB SPARTA IDR"
    List1.AddItem "10-CIMB SPARTA USD"
    List1.AddItem "11-BCA Gunarso Dede"
    List1.AddItem "12-BCA GB or MB"
    List1.AddItem "13-BCA FLEET"
    List1.AddItem "14-OCBC NISP IDR"
    List1.AddItem "15-OCBC NISP MULTI CURR"
    List1.AddItem "16-BCA 10"
    
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "select * from gl_masterac where noac = '" & x_original(grid.TextMatrix(grid.Row, 1)) & "'"
        Set RST = OBJ.Execute(SQL)
        lblnamacc = "Nama Account : " & RST!nmac
        OBJ.Close
    End If
    If txtkodecomp = "" Or txtkodetran = "" Or txtnotran = "" Or txtkodecur = "" Then Exit Sub
    posrow = grid.Row
    Select Case grid.Col
        Case 0
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
        Case 1
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 2
            If grid.TextMatrix(grid.Row, 1) = "" Or txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Or txtnilai.Visible = True Then Exit Sub
    
            'If grid.TextMatrix(grid.Row, 4) = "0.00" Then
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
                carisisa
            'End If
            
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Or txtnilai.Visible = True Then Exit Sub
            
            'If grid.TextMatrix(grid.Row, 3) = "0.00" Then
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
                carisisa
            'End If
    End Select
End Sub

Private Sub grid_EnterCell()
    Select Case grid.Col
    Case 1
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 2
        If grid.TextMatrix(grid.Row, 1) = "" Or txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 3
        'If grid.TextMatrix(grid.Row, 1) = "" Or grid.TextMatrix(grid.Row, 4) <> "0.00" Or txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        carisisa
    Case 4
        'If grid.TextMatrix(grid.Row, 1) = "" Or grid.TextMatrix(grid.Row, 3) <> "0.00" Or txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        'carisisa
        
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    grid.Row = posrow
    grid.Col = 1
    grid.CellAlignment = 1
    str6 = grid.TextMatrix(grid.Row, 1)
    grid.TextMatrix(grid.Row, 1) = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(grid.TextMatrix(grid.Row, 1)) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!flag = 1 Then
        grid.TextMatrix(grid.Row, 1) = str6
        
        OBJ.Close
        Exit Sub
    End If
    lblnamacc = "Nama Account : " & RST!nmac
    OBJ.Close
    
    If grid.Row <> 1 Then grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row - 1, 2)
    If grid.Row = 1 Then grid.TextMatrix(grid.Row, 2) = txtketcash
    
    grid.Col = 0
    Set grid.CellPicture = uncheck.Picture
    grid.SetFocus
    grid.Col = 2

    txtket.Width = grid.ColWidth(grid.Col) - 40
    txtket = grid.TextMatrix(grid.Row, grid.Col)
    txtket.Left = grid.Left + grid.CellLeft
    txtket.Top = grid.Top + grid.CellTop + 20
    txtket.Visible = True
    txtket.SetFocus
    
    If grid.TextMatrix(grid.Row, 3) = "" And grid.TextMatrix(grid.Row, 4) = "" Then
        grid.TextMatrix(grid.Row, 3) = "0.00"
        grid.TextMatrix(grid.Row, 4) = "0.00"
    End If
    
    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
End Sub

Private Sub optnonpajak_Click()
    txtnobp = GetnobpnonPajak
    txtketcash.SetFocus
End Sub

Private Sub optpajak_Click()
    txtnobp = GetnobpPajak
    txtketcash.SetFocus
End Sub

Private Sub txtcash_GotFocus()
    If hasil = "" Then Exit Sub
    txtcash = hasil
    carinoac
    hasil = ""
End Sub

Private Sub txtcekbg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 38 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtketcash_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 38 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtcekbg.SetFocus
End Sub

Private Sub txtncash_Change()
On Error Resume Next
    txtkredit = txtncash
    lblcomkredit = "1 Lines"
    txtpembulatan = txtncash * txtnilaikurs
    txtpembulatan = SpyRound(txtpembulatan)
End Sub

Private Sub txtdebet_Change()
    If txtdebet = txtkredit Then
        lblstatus = "Status : Balance"
        lblstatus.BackColor = &H80000011
        lblbal = "B"
        cmdadd.Enabled = True
    Else
        lblstatus = "Status : UnBalance"
        lblstatus.BackColor = vbRed
        lblbal = "U"
        cmdadd.Enabled = False
    End If
End Sub

Private Sub txtkodecomp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecomp_LostFocus
End Sub

Private Sub txtkodecomp_LostFocus()
    If txtkodecomp = "" Then Exit Sub
    If txtkodecomp.SelLength <> 0 Then Exit Sub
    hapusemua
    txtkodetran = ""
    txtnotran = ""
    cmbdaerah = ""
    date1 = Date
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacomp = RST!nmcompscr
        format_coa = RST!formatac
        date1.SetFocus
    Else
        MsgBox "Company " & txtkodecomp & " Not Found.", vbInformation, "Information"
        txtkodecomp = ""
        txtkodecomp.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkodecur_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecur_LostFocus
    If txtkodecur <> "" Then txtkodecur_LostFocus
End Sub

Private Sub txtkodecur_LostFocus()
    carikurs
End Sub

Private Sub carikurs()
    If txtkodecur = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_kurs where kdkurs = '" & txtkodecur & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacur = "Currency : " & RST!nmkurs
        If RST!base = 1 Then
            lblbase = "1"
        Else
            lblbase = "0"
        End If
        Select Case Month(date1)
        Case 1
            txtnilaikurs = RST!kurs1
        Case 2
            txtnilaikurs = RST!kurs2
        Case 3
            txtnilaikurs = RST!kurs3
        Case 4
            txtnilaikurs = RST!kurs4
        Case 5
            txtnilaikurs = RST!kurs5
        Case 6
            txtnilaikurs = RST!kurs6
        Case 7
            txtnilaikurs = RST!kurs7
        Case 8
            txtnilaikurs = RST!kurs8
        Case 9
            txtnilaikurs = RST!kurs9
        Case 10
            txtnilaikurs = RST!kurs10
        Case 11
            txtnilaikurs = RST!kurs11
        Case 12
            txtnilaikurs = RST!kurs12
        End Select
        grid.SetFocus
    Else
        MsgBox "Currency " & txtkodecur & " Not Found.", vbInformation, "Information"
        txtkodecur = ""
        txtkodecur.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtncash.SetFocus
End Sub

Private Sub txtcash_LostFocus()
    If txtcash = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.noac = '" & x_original(txtcash) & "' and a.kdcomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = 1 Then
            txtcash = ""
            txtcash.SetFocus
            
            OBJ.Close
            Exit Sub
        End If
        txtcash = original(txtcash)
        lblcash = "Cash/Bank : " & RST!nmac
        OBJ.Close
    Else
        OBJ.Close
        txtcash = ""
        lblcash = "Cash/Bank : "
        txtcash.SetFocus
        
        carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
        namatabel = "Company Account"

        frmsearch.Show vbModal
    End If
    
    'OBJ1.Open dsn
    'SQL1 = "select * from gl_cash where noac = '" & x_original(txtcash) & "'"
    'Set RST1 = OBJ1.Execute(SQL1)
    'If Not RST1.EOF Then
    '    OBJ1.Close
    '    GoTo jump_0000
    'Else
    '    SQL1 = "select * from gl_bank where noac = '" & x_original(txtcash) & "'"
    '    Set RST1 = OBJ1.Execute(SQL1)
    '    If Not RST1.EOF Then
    '        OBJ1.Close
    '        GoTo jump_0000
    '    Else
    '        MsgBox "Cash/Bank " & txtcash & " Not Found.", vbInformation, "Information"
    '        txtcash = ""
    '        txtcash.SetFocus
    '    End If
    'End If
    'OBJ1.Close
    'Exit Sub
'jump_0000:
    
    'carinoac
    '=====================================================================================
End Sub

Private Sub carinoac()
    If txtcash = "" Then Exit Sub
    OBJ1.Open dsn
    SQL1 = "select * from gl_masterac where noac = '" & x_original(txtcash) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        txtcash = original(RST1!noac)
        lblcash = "Cash/Bank : " & RST1!nmac
        txtncash.SetFocus
    Else
        MsgBox "Cash/Bank " & txtcash & " Not Found.", vbInformation, "Information"
        txtcash = ""
        txtcash.SetFocus
    End If
    OBJ1.Close
End Sub

Private Sub txtkodetran_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnotran.SetFocus
End Sub

Private Sub txtkodetran_LostFocus()
    txtnotran = ""
    hapusemua
    cari_out
End Sub

Private Sub hapusemua()
    txtkodecur = ""
    lblnamacur = "Currency :"
    'txtcash = ""
    'lblcash = "Cash/Bank : "
    txtnilaikurs = 0
    txtnilaikurs2 = 0
    txtpembulatan = 0
    txtdebet = 0
    txtkredit = 0
    txtncash = 0
    txtppn = 0
    txtketcash = ""
    txtcekbg = ""
    lblstatus = "Status :"
    lblcomdebet = "Lines"
    lblcomkredit = "Lines"
    lblnamacc = "Nama Account :"
    lblbal = ""
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 2) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 4350
    grid.ColWidth(3) = 1650
    grid.ColWidth(4) = 0
End Sub

Private Sub Genapkan()
    grid.Row = 1
    txtpembulatan = 0
    str2 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do
        If grid.TextMatrix(grid.Row, 4) <> "0.00" Then str2 = str2 + 1
        txtpembulatan = txtpembulatan + Val(Format(grid.TextMatrix(grid.Row, 4), "general number"))
        If grid.TextMatrix(grid.Row + 1, 2) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    
End Sub

Private Sub debet()
    grid.Row = 1
    txtdebet = 0
    str2 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) <> "0.00" Then str2 = str2 + 1
        txtdebet = txtdebet + Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
        If grid.TextMatrix(grid.Row + 1, 2) = "" Then Exit Do
        'If grid.TextMatrix(grid.Row + 1, 1) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    lblcomdebet = str2 & " Lines"
End Sub

Private Sub kredit()
    grid.Row = 1
    txtkredit = 0
    str3 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do
        If grid.TextMatrix(grid.Row, 4) <> "0.00" Then str3 = str3 + 1
        txtkredit = txtkredit + Val(Format(grid.TextMatrix(grid.Row, 4), "general number"))
        If grid.TextMatrix(grid.Row + 1, 2) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    lblcomkredit = str3 & " Lines"
End Sub

Private Sub txtkredit_Change()
    If txtdebet = txtkredit Then
        lblstatus = "Status : Balance"
        lblstatus.BackColor = &H80000011
        lblbal = "B"
        cmdadd.Enabled = True
    Else
        lblstatus = "Status : UnBalance"
        lblstatus.BackColor = vbRed
        lblbal = "U"
        cmdadd.Enabled = False
    End If
End Sub

Private Sub txtncash_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        txtketcash.SetFocus
        
    txtkredit = txtncash
    lblcomkredit = "1 Lines"
    txtpembulatan = txtncash * txtnilaikurs
    txtpembulatan = SpyRound(txtpembulatan)
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        Select Case grid.Col
        Case 3
            grid.TextMatrix(posrow, 4) = Format(txtnilaikurs, "###,###,##0.00") * Format(grid.TextMatrix(posrow, 3), "###,###,##0.00")
            grid.TextMatrix(posrow, 4) = Format(grid.TextMatrix(posrow, 4), "###,###,##0.00")
            debet
        Case 4
            'kredit
            grid.TextMatrix(posrow, 4) = SpyRound(grid.TextMatrix(posrow, 4))
            grid.TextMatrix(posrow, 4) = Format(grid.TextMatrix(posrow, 4), "###,###,##0.00")
           ' Genapkan
        End Select
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 38 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 1
                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                str6 = grid.TextMatrix(grid.Row, 1)
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False
        
                OBJ.Open dsn
                'sql = "select * from gl_masterac where noac = '" & x_original(grid.TextMatrix(grid.Row, 1)) & "'"
                SQL = "select b.noac, b.nmac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.noac = '" & x_original(grid.TextMatrix(grid.Row, 1)) & "' and a.kdcomp = '" & txtkodecomp & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    If RST!flag = 1 Then
                        grid.TextMatrix(grid.Row, 1) = str6
                        
                        OBJ.Close
                        Exit Sub
                    End If
                    
                    grid.TextMatrix(grid.Row, 1) = original(RST!noac)
                    lblnamacc = "Nama Account : " & RST!nmac
                    OBJ.Close
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    
                    If grid.Row <> 1 Then grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row - 1, 2)
                    If grid.Row = 1 Then grid.TextMatrix(grid.Row, 2) = txtketcash
    
                    If grid.TextMatrix(grid.Row, 3) = "" And grid.TextMatrix(grid.Row, 4) = "" Then
                        grid.TextMatrix(grid.Row, 3) = "0.00"
                        grid.TextMatrix(grid.Row, 4) = "0.00"
                    End If
    
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    
                    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
                    namatabel = "Company Account"
    
                    frmsearch.Show vbModal
                End If
                grid.Col = 1
            Case 2
                grid.TextMatrix(posrow, 2) = txtket
                txtket = ""
                grid.SetFocus
                grid.Row = posrow
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtnilaikurs_Change()
On Error Resume Next
    txtpembulatan = txtncash * txtnilaikurs
    txtpembulatan = SpyRound(txtpembulatan)
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtnotran_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtkodetran <> "BK" Then KeyCode = 0
End Sub

Private Sub txtnotran_KeyPress(KeyAscii As Integer)
    If txtkodetran <> "BK" Then KeyAscii = 0 Else KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtkodetran <> "BK" Then txtkodecur.SetFocus
End Sub

Private Sub txtnotran_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtkodetran = "BK" Then
        hapusemua
        cari_out
    End If
End Sub

Private Sub txtnotran_LostFocus()
    Frame1.Visible = False
    If txtkodecomp = "" Or txtkodetran = "" Or txtnotran = "" Or cmbdaerah = "" Then Exit Sub
    hapusemua
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx = '" & txtnotran & "' order by lineitem asc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then MsgBox "Transaction " & txtkodetran & txtnotran & " Already Exsist.", vbInformation, "Information"
    OBJ.Close
    
    If txtkodetran = "BK" Then
        Select Case Mid(txtnotran, 6, 2)
            Case "01"
                txtcash = "11102001"
            Case "02"
                txtcash = "11102001"
            Case "03"
                txtcash = "11102001"
            Case "04"
                txtcash = "22001004"
            Case "05"
                txtcash = "11102001"
            Case "06"
                txtcash = "11102001"
            Case "07"
                txtcash = "11102001"
            Case "08"
                txtcash = "11102001"
            Case "09"
                txtcash = "11102001"
            Case "10"
                txtcash = "11102001"
            Case "11"
                txtcash = "11102001"
            Case "12"
                txtcash = "11102020"
            Case "13"
                txtcash = "11102021"
            Case "14"
                txtcash = "22001016"
            Case "15"
                txtcash = "11102022"
            Case "16"
                txtcash = "11102026"
        End Select
        
        OBJ1.Open dsn
        SQL1 = "select * from gl_masterac where noac = '" & x_original(txtcash) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtcash = original(RST1!noac)
            lblcash = "Cash/Bank : " & RST1!nmac
        Else
            MsgBox "Cash/Bank " & txtcash & " Not Found.", vbInformation, "Information"
            txtcash = ""
            txtcash.SetFocus
        End If
        OBJ1.Close
    End If
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    
    grid.Col = 0
    Set grid.CellPicture = blank
    debet
    If grid.Rows = 2 Then lblstatus = "Status :"
End Sub

Private Sub cari_out()
    If txtkodecomp = "" Or cmbdaerah = "" Or txtkodetran = "" Then Exit Sub
    If txtkodetran = "BK" Then
        If Len(txtnotran) = 8 Then
            If Not (Left(txtnotran, 2) >= "08" And Left(txtnotran, 2) < "99") Then
                MsgBox "Format digit pertama dan kedua salah, format yang dipakai adalah format tahun, YY", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotran, 3, 2) >= "01" And Mid(txtnotran, 3, 2) <= "13") Then
                MsgBox "Format digit ketiga dan keempat salah, format yang dipakai adalah format bulan, MM", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not Mid(txtnotran, 5, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not Right(txtnotran, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotran, 6, 2) >= "01" And Mid(txtnotran, 6, 2) <= "09") And Not Mid(txtnotran, 6, 2) <= "17" Then
                MsgBox "Format digit keenam dan ketujuh salah, tekan F2 untuk melihat list.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
                    
            OBJ.Open dsn
            SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx like '" & txtnotran & "%' and flagprint='O' order by notrx desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If Len(RST!notrx + 1) = 5 Then
                    txtnotran = txtnotran & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 4 Then
                    txtnotran = txtnotran & "0" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 3 Then
                    txtnotran = txtnotran & "00" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 2 Then
                    txtnotran = txtnotran & "000" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 1 Then
                    txtnotran = txtnotran & "0000" & RST!notrx + 1
                End If
            Else
                txtnotran = txtnotran & "00001"
            End If
            OBJ.Close
            txtnovoucher.SetFocus
            'txtkodecur.SetFocus
        End If
    Else
        OBJ.Open dsn
        SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx like '" & Format(date1, "YYMM") & "/" & cmbdaerah & "/%' and flagprint='O' order by notrx desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If Len(RST!notrx + 1) = 5 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 4 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 3 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 2 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/000" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 1 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0000" & RST!notrx + 1
            End If
        Else
            txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00001"
        End If
        OBJ.Close
    End If
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function GetnobpnonPajak() As String
On Error GoTo err_handler:
    Dim i As Long
    Dim strformat As String
    Dim int_kode As Integer
    Dim temp_kode As String
    OBJ.Open dsn
    SQL = "Select max(no_payment) as nobp From no_bank_payment "
    SQL = SQL + "Where no_payment like '" + strformat + "%' and is_pajak ='0'"
    Set RST = OBJ.Execute(SQL)
    If RST!nobp = "" Then
        temp_kode = "K00001"
    End If
    If RST!nobp <> "" Then
        int_kode = Right(RST!nobp, 5)
        int_kode = int_kode + 1
        
        If (Len(Trim(Str(int_kode))) = 1) Then
            temp_kode = int_kode
            temp_kode = "K0000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 2) Then
            temp_kode = int_kode
            temp_kode = "K000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 3) Then
            temp_kode = int_kode
            temp_kode = "K00" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 4) Then
            temp_kode = int_kode
            temp_kode = "K0" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 5) Then
            temp_kode = int_kode
            temp_kode = "K" + Trim(Str(int_kode))
        End If
    End If
    GetnobpnonPajak = temp_kode
    OBJ.Close
    Exit Function
err_handler:
    GetnobpnonPajak = "K00001"
    OBJ.Close

End Function

Function GetnobpPajak()
On Error GoTo err_handler:
    Dim i As Long
    Dim int_kode As Long
    Dim temp_kode As String
    Dim tempyear As String
    tempyear = Format(Date, "yy")
    OBJ.Open dsn
    SQL = "Select max(no_payment) as nobp From no_bank_payment Where is_pajak ='1' and no_payment like '" & tempyear & "%'"
    SQL = SQL + " and tgljt>'2021-12-31'"
    Set RST = OBJ.Execute(SQL)
    If RST!nobp = "" Or IsNull(RST!nobp) Then
        'temp_kode = "00001"
        temp_kode = "0001"
    End If
    If RST!nobp <> "" Then
        int_kode = RST!nobp ' + 1
        int_kode = int_kode + 1
        
        'If (Len(Trim(Str(int_kode))) = 1) Then
            'temp_kode = "0000" + Trim(Str(int_kode))
        'End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 1) Then
            temp_kode = "000" + Trim(Str(Right(int_kode, 1)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 2) Then
            temp_kode = "00" + Trim(Str(Right(int_kode, 2)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 3) Then
            temp_kode = "0" + Trim(Str(Right(int_kode, 3)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 4) Then
            temp_kode = Trim(Str(Right(int_kode, 4)))
        End If
    End If
    GetnobpPajak = Format(Date, "yy") & temp_kode
    OBJ.Close
    Exit Function
err_handler:
    GetnobpPajak = "00001"
    OBJ.Close
End Function

Private Sub txtnovoucher_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler:
    If KeyAscii = 13 Then
        hapusemua
        If txtnovoucher.text = "" Then Exit Sub
        If txtkodecomp = "" Or txtkodetran = "" Or txtnotran = "" Or cmbdaerah = "" Then
            MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
            Exit Sub
        End If
        
            OBJ.Open dsn
            SQL = "Select ref1 From am_beliapp Where ref1= '" & txtnovoucher & "'"
            Set RST = OBJ.Execute(SQL)
            
            If RST.EOF Then
'MsgBox "NON BB (Note)"
                Check2.Visible = False
                Check1.Visible = False
                NonBahanBaku
            Else
                SQL = "Select flag From no_bank_payment Where no_voucher= '" & txtnovoucher & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    Check1.Visible = True
                Else
'MsgBox "BB"
                    If RST!flag = "1" Then
                        Check1.Visible = False
                    ElseIf RST!flag = "0" Then
                        Check1.Visible = True
                    End If
                End If
                Check2.Value = Checked
                BahanBaku
            End If
            debet
            OBJ.Close
            carikurs
            optpajak_Click 'non pajak tdk pakai
            txtcash.SetFocus
    End If
err_handler:
'    OBJ.Close
End Sub

Private Sub BahanBaku()
poscol = grid.Col
posrow1 = grid.Row

    SQL = "Select a.keterangan,a.perkiraan,a.jumlah,b.kepada,b.npwp,b.alamat,b.kdkurs,b.nilai,b.ppn,c.nobeli,c.nilaikurs,c.ref1,c.kodesupp "
    ',d.noac,d.nmac
    '===Tidak ada No Perkiraan===' di form Vina(penerimaan app)
    SQL = SQL + "From am_voucherin a inner join am_voucherhdr b "
    SQL = SQL + "On a.novoucher=b.novoucher inner join am_beliapp c "
    SQL = SQL + "On c.ref1=a.novoucher Where a.novoucher = '" & txtnovoucher & "' and a.keterangan = c.KodeBarang"
    Set RST = OBJ.Execute(SQL)

    If Not RST.EOF Then
        txtkpd.text = RST!kepada
        If Not IsNull(RST!kdkurs) Then txtkodecur.text = RST!kdkurs

            txtppn.text = Format(RST!ppn, "###,###,##0.00")
            txtnilaikurs2.text = Format(RST!nilaikurs, "###,###,##0.00")
            txtnobukti.text = RST!nobeli
            txtkdsupp.text = RST!kodesupp
            txtketcash.text = RST!keterangan
'            lblnamacc = "Nama Account : " & RST!nmac
    End If

        grid.Row = 1
        Do While Not RST.EOF
            With grid
                .Col = 0
                Set .CellPicture = uncheck.Picture
'                .TextMatrix(.Row, 1) = RST!noac
                .TextMatrix(.Row, 2) = RST!keterangan
                .TextMatrix(.Row, 3) = Format(RST!jumlah, "###,###,##0.00")
                .TextMatrix(.Row, 4) = Format(txtnilaikurs2, "###,###,##0.00") * Format(RST!jumlah, "###,###,##0.00")
                .TextMatrix(.Row, 4) = Format(.TextMatrix(.Row, 4), "###,###,##0.00")
                .Rows = .Rows + 1
                .Row = .Row + 1
            End With
            RST.MoveNext
            DoEvents
        Loop
        
        'carikurs
        'OBJ.Open dsn
        SQL = "Select SUM(jumlah) as total From am_voucherin where novoucher = '" & txtnovoucher & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtncash.text = Format(RST!total, "###,###,##0.00")
        End If
End Sub

Private Sub NonBahanBaku()
    SQL = "Select a.keterangan,a.perkiraan,a.jumlah,b.kepada,b.npwp,b.alamat,b.kdkurs,"
    SQL = SQL + "b.nilai,b.ppn "
    SQL = SQL + "From am_voucherin a inner join am_voucherhdr b "
    SQL = SQL + "On a.novoucher=b.novoucher where a.novoucher = '" & txtnovoucher & "'"
    Set RST = OBJ.Execute(SQL)
                
    If Not RST.EOF Then
        txtkpd.text = RST!kepada
        If Not IsNull(RST!kdkurs) Then txtkodecur.text = RST!kdkurs
            txtnilaikurs2.text = Format(RST!nilai, "###,###,##0.00")
            txtppn.text = Format(RST!ppn, "###,###,##0.00")
        End If
          
        grid.Row = 1
        Do While Not RST.EOF
            With grid
                grid.Col = 0
                Set grid.CellPicture = uncheck.Picture
                grid.TextMatrix(grid.Row, 1) = RST!perkiraan
                'MsgBox Len(RST!keterangan)
                If Len(RST!keterangan) > 60 Then
                    MsgBox "Data tidak dapat ditampilkan" + Chr(13) + "Keterangan dengan Account, " _
                    + RST!perkiraan + " terlalu panjang", vbExclamation, "WARNING !"
                    cmdclear_Click
                    Exit Sub
                End If
                grid.TextMatrix(grid.Row, 2) = RST!keterangan
                grid.TextMatrix(grid.Row, 3) = Format(RST!jumlah, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 4) = Format(txtnilai, "###,###,##0.00") * Format(RST!jumlah, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 4) = Format(grid.TextMatrix(grid.Row, 4), "###,###,##0.00")
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
            End With
            RST.MoveNext
            DoEvents
        Loop
    
        'carikurs
        'OBJ.Open dsn
        SQL = "Select SUM(jumlah) as total From am_voucherin where novoucher = '" & txtnovoucher & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtncash.text = Format(RST!total, "###,###,##0.00")
        End If
        If txtkodecur <> "IDR" Then
            Check1.Visible = True
        End If
                
End Sub

Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function
Private Function SpyRoundUp(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.1) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRoundUp = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRoundUp = Val(arVal(0)) Else: SpyRoundUp = Val(arVal(0)) + 1
End Function
