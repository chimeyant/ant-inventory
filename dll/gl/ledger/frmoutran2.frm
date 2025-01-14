VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmoutran2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "test"
      Height          =   495
      Left            =   4440
      TabIndex        =   50
      Top             =   6720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   2380
      Left            =   5040
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
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
         TabIndex        =   48
         Top             =   0
         Width           =   2535
      End
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "frmoutran2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":006C
      Key             =   "frmoutran2.frx":008A
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
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmoutran2.frx":00CE
      Caption         =   "frmoutran2.frx":00EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":015A
      Keys            =   "frmoutran2.frx":0178
      Spin            =   "frmoutran2.frx":01C2
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   7200
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
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   7200
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
      TabIndex        =   23
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
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
      Top             =   1440
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
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":01EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":0256
      Key             =   "frmoutran2.frx":0274
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
      Top             =   2160
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":02B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":0324
      Key             =   "frmoutran2.frx":0342
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
      Top             =   2160
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":0386
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":03F2
      Key             =   "frmoutran2.frx":0410
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
      TabIndex        =   5
      Top             =   2520
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":0454
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":04C0
      Key             =   "frmoutran2.frx":04DE
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
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran2.frx":0522
      Caption         =   "frmoutran2.frx":0542
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":05AE
      Keys            =   "frmoutran2.frx":05CC
      Spin            =   "frmoutran2.frx":0616
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Save"
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
      MICON           =   "frmoutran2.frx":063E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   5760
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
      MICON           =   "frmoutran2.frx":0958
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   5760
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
      MICON           =   "frmoutran2.frx":0C72
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
      Left            =   4920
      TabIndex        =   31
      Top             =   480
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran2.frx":0F8C
      Caption         =   "frmoutran2.frx":0FAC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1018
      Keys            =   "frmoutran2.frx":1036
      Spin            =   "frmoutran2.frx":1080
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
      ValueVT         =   83165189
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtdebet 
      Height          =   285
      Left            =   4920
      TabIndex        =   32
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran2.frx":10A8
      Caption         =   "frmoutran2.frx":10C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1134
      Keys            =   "frmoutran2.frx":1152
      Spin            =   "frmoutran2.frx":119C
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
      ValueVT         =   83165189
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   37
      Top             =   2520
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
      MICON           =   "frmoutran2.frx":11C4
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
      TabIndex        =   38
      Top             =   1080
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
      MICON           =   "frmoutran2.frx":14DE
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
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":17F8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1864
      Key             =   "frmoutran2.frx":1882
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
      TabIndex        =   39
      Top             =   2880
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
      MICON           =   "frmoutran2.frx":18C6
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
      Left            =   4920
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmoutran2.frx":1BE0
      Caption         =   "frmoutran2.frx":1C00
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1C6C
      Keys            =   "frmoutran2.frx":1C8A
      Spin            =   "frmoutran2.frx":1CD4
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText txtketcash 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   3240
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":1CFC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1D68
      Key             =   "frmoutran2.frx":1D86
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
      TabIndex        =   10
      Top             =   3600
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   503
      Caption         =   "frmoutran2.frx":1DCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoutran2.frx":1E36
      Key             =   "frmoutran2.frx":1E54
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
      TabIndex        =   49
      Top             =   2010
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
      TabIndex        =   46
      Top             =   2190
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
      TabIndex        =   45
      Top             =   2190
      Width           =   1455
   End
   Begin MSForms.ComboBox cmbdaerah 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
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
      TabIndex        =   44
      Top             =   1830
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
      TabIndex        =   43
      Top             =   3630
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
      TabIndex        =   42
      Top             =   3270
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
      Left            =   3480
      TabIndex        =   41
      Top             =   2910
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
      TabIndex        =   40
      Top             =   5730
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
      Left            =   6480
      TabIndex        =   36
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
      Left            =   6480
      TabIndex        =   35
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
      Left            =   3960
      TabIndex        =   34
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
      Left            =   3960
      TabIndex        =   33
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
      TabIndex        =   29
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
      TabIndex        =   28
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   7200
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
      TabIndex        =   22
      Top             =   5970
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
      TabIndex        =   21
      Top             =   6210
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
      TabIndex        =   20
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Nilai Kurs"
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
      Left            =   3840
      TabIndex        =   18
      Top             =   2550
      Width           =   975
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
      TabIndex        =   17
      Top             =   1470
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
      Left            =   4920
      TabIndex        =   19
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmoutran2"
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
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtnilaikurs = 0 Then
        MsgBox "Nilai Kurs = 0.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    If Len(Trim(txtkodetran)) = 0 Or Len(Trim(txtnotran)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
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
            MsgBox "Data Entry Not Complite, On Row " & grid.Row, vbExclamation, "Warning"
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
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "'" & txtkodetran & "',"
    SQL = SQL + "'" & txtnotran & "',"
    SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
    SQL = SQL + "'" & x_original(txtcash) & "',"
    SQL = SQL + "'" & txtketcash & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "convert(money,'" & (txtncash * txtnilaikurs) & "'),"
    SQL = SQL + "convert(money,'" & txtncash & "'),"
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
    
    grid.Row = 1
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
        SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
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
    'grid.Row = 1
    
    'OBJ.Open dsn
    'SQL = "insert into journal_umum "
    'SQL = SQL + "(kdcomp, "
    'SQL = SQL + "tgltrx, "
    'SQL = SQL + "kdtrx, "
    'SQL = SQL + "notrx, "
    'SQL = SQL + "kurs, "
    'SQL = SQL + "noactrx, "
    'SQL = SQL + "desctrx, "
    'SQL = SQL + "dbkrtrx, "
    'SQL = SQL + "amounttrx, "
    'SQL = SQL + "nilaitrx_debet, "
    'SQL = SQL + "nilaitrx_kredit,"
    'SQL = SQL + "debet,"
    'SQL = SQL + "kredit,"
    'SQL = SQL + "currtrx, "
    'SQL = SQL + "flag, "
    'SQL = SQL + "flagprint, "
    'SQL = SQL + "flagadjustment, "
    'SQL = SQL + "cekbg, "
    'SQL = SQL + "identry, "
    'SQL = SQL + "idupdate, "
    'SQL = SQL + "dateentry, "
    'SQL = SQL + "dateupdate, "
    'SQL = SQL + "lineitem)"
    
    'SQL = SQL + " values"
    'SQL = SQL + "('" & txtkodecomp & "',"
    'SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    'SQL = SQL + "'" & txtkodetran & "',"
    'SQL = SQL + "'" & txtnotran & "',"
    'SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
    'SQL = SQL + "'" & x_original(txtcash) & "',"
    'SQL = SQL + "'" & txtketcash & "',"
    'SQL = SQL + "'K',"
    'SQL = SQL + "convert(money,'" & txtncash * txtnilaikurs & "'),"
    
    'transaksi asli
    'SQL = SQL + "'0',"
    'SQL = SQL + "convert(money,'" & txtncash & "'),"
    
    'transaksi dalam rupiah
    'SQL = SQL + "'0',"
    'SQL = SQL + "convert(money,'" & txtncash * txtnilaikurs & "'),"
    
    'SQL = SQL + "'" & txtkodecur & "',"
    'SQL = SQL + "'" & lblbal & "',"
    'SQL = SQL + "'O',"
    'SQL = SQL + "'" & cmbdaerah & "',"
    'SQL = SQL + "'" & txtcekbg & "',"
    'SQL = SQL + "'" & kuser & "',"
    'SQL = SQL + "'',"
    'SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    'SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    'SQL = SQL + "convert(numeric,'1'))"
    'Set RST = OBJ.Execute(SQL)
    
    'Do While True
        'If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        'SQL = "insert into journal_umum "
        'SQL = SQL + "(kdcomp, "
        'SQL = SQL + "tgltrx, "
        'SQL = SQL + "kdtrx, "
        'SQL = SQL + "notrx, "
        'SQL = SQL + "kurs, "
        'SQL = SQL + "noactrx, "
        'SQL = SQL + "desctrx, "
        'SQL = SQL + "dbkrtrx, "
        'SQL = SQL + "amounttrx, "
        'SQL = SQL + "nilaitrx_debet, "
        'SQL = SQL + "nilaitrx_kredit,"
        'SQL = SQL + "debet,"
        'SQL = SQL + "kredit,"
        'SQL = SQL + "currtrx, "
        'SQL = SQL + "flag, "
        'SQL = SQL + "flagprint, "
        'SQL = SQL + "flagadjustment, "
        'SQL = SQL + "cekbg, "
        'SQL = SQL + "identry, "
        'SQL = SQL + "idupdate, "
        'SQL = SQL + "dateentry, "
        'SQL = SQL + "dateupdate, "
        'SQL = SQL + "lineitem)"
        
        'SQL = SQL + " values"
        'SQL = SQL + "('" & txtkodecomp & "',"
        'SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
        'SQL = SQL + "'" & txtkodetran & "',"
        'SQL = SQL + "'" & txtnotran & "',"
        'SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
        'SQL = SQL + "'" & x_original(grid.TextMatrix(grid.Row, 1)) & "',"
        'SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        'SQL = SQL + "'D',"
        'SQL = SQL + "convert(money,'" & (Format(grid.TextMatrix(grid.Row, 3), "general number") * txtnilaikurs) & "'),"
        
        'nilai transaksi asli
        'SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        'SQL = SQL + "'0',"
        
        'nilai transaksi rupiah
        'SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        'SQL = SQL + "'0',"
        
        'SQL = SQL + "'" & txtkodecur & "',"
        'SQL = SQL + "'" & lblbal & "',"
        'SQL = SQL + "'O',"
        'SQL = SQL + "'" & cmbdaerah & "',"
        'SQL = SQL + "'" & txtcekbg & "',"
        'SQL = SQL + "'" & kuser & "',"
        'SQL = SQL + "'',"
        'SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
        'SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
        'SQL = SQL + "convert(numeric,'" & grid.Row + 1 & "'))"
            
        'Set RST = OBJ.Execute(SQL)
        'grid.Row = grid.Row + 1
    'Loop
    
    OBJ.Close
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click

End Sub

Private Sub cmdclear_Click()
    hapusemua
    txtkodecomp = ""
    lblnamacomp = ""
    cmbdaerah = ""
    date1.Value = Date
    txtkodetran = ""
    txtnotran = ""
    txtcash = ""
    lblcash = "Cash/Bank : "
    txtkodecomp.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
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

Private Sub Command1_Click()
    
'PRINT PREVIEW
    With rpt_trans_out
        .lbljurnal = "BANK PAYMENT"
        .lblcompany = "PT. SPARTA PRIMA"
        '.lbltanggal = "Dari : " & Format(date1, "dd-MM-yyyy") & " s.d : " & Format(date2, "dd-MM-yyyy")
        '.DataControl1.Source = SQL
        '.DataControl1.ConnectionString = dsn
        .Show
    End With
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
    grid.TextMatrix(0, 4) = "Kredit"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2350
    grid.ColWidth(3) = 1650
    grid.ColWidth(4) = 0
    
    grid.RowHeightMin = 300
    date1.Value = Date
    
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
            
            If grid.TextMatrix(grid.Row, 4) = "0.00" Then
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
                carisisa
            End If
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Or txtnilai.Visible = True Then Exit Sub
            
            If grid.TextMatrix(grid.Row, 3) = "0.00" Then
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
                carisisa
            End If
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
        If grid.TextMatrix(grid.Row, 1) = "" Or grid.TextMatrix(grid.Row, 4) <> "0.00" Or txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        carisisa
    Case 4
        If grid.TextMatrix(grid.Row, 1) = "" Or grid.TextMatrix(grid.Row, 3) <> "0.00" Or txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        carisisa
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
    txtkredit = txtncash
    lblcomkredit = "1 Lines"
End Sub

Private Sub txtdebet_Change()
    If txtdebet = txtkredit Then
        lblstatus = "Status : Balance"
        lblbal = "B"
    Else
        lblstatus = "Status : UnBalance"
        lblbal = "U"
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
    txtdebet = 0
    txtkredit = 0
    txtncash = 0
    txtketcash = ""
    txtcekbg = ""
    lblstatus = "Status :"
    lblcomdebet = "Lines"
    lblcomkredit = "Lines"
    lblnamacc = "Nama Account :"
    lblbal = ""
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
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
    grid.ColWidth(2) = 2350
    grid.ColWidth(3) = 1650
    grid.ColWidth(4) = 0
End Sub

Private Sub debet()
    grid.Row = 1
    txtdebet = 0
    str2 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) <> "0.00" Then str2 = str2 + 1
        txtdebet = txtdebet + Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then Exit Do
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
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    lblcomkredit = str3 & " Lines"
End Sub

Private Sub txtkredit_Change()
    If txtdebet = txtkredit Then
        lblstatus = "Status : Balance"
        lblbal = "B"
    Else
        lblstatus = "Status : UnBalance"
        lblbal = "U"
    End If
End Sub

Private Sub txtncash_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtketcash.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        Select Case grid.Col
        Case 3
            debet
        Case 4
            'kredit
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
    
    'OBJ.Open dsn
    'SQL = "select * from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx = '" & txtnotran & "' order by lineitem asc"
    'Set RST = OBJ.Execute(SQL)
    'If Not RST.EOF Then MsgBox "Transaction " & txtkodetran & txtnotran & " Already Exsist.", vbInformation, "Information"
    'OBJ.Close
    
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
            
            txtkodecur.SetFocus
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
        'OBJ.Open dsn
        'SQL = "select top 1 right(notrx,5)'notrx' from journal_umum where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkodetran & "' and notrx like '" & Format(date1, "YYMM") & "/" & cmbdaerah & "/%' and flagprint='O' order by notrx desc"
        'Set RST = OBJ.Execute(SQL)
        'If Not RST.EOF Then
            'If Len(RST!notrx + 1) = 5 Then
                'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/" & RST!notrx + 1
            'ElseIf Len(RST!notrx + 1) = 4 Then
                'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0" & RST!notrx + 1
            'ElseIf Len(RST!notrx + 1) = 3 Then
                'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00" & RST!notrx + 1
            'ElseIf Len(RST!notrx + 1) = 2 Then
                'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/000" & RST!notrx + 1
            'ElseIf Len(RST!notrx + 1) = 1 Then
                'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0000" & RST!notrx + 1
            'End If
        'Else
            'txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00001"
        'End If
        'OBJ.Close
    End If
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
