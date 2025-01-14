VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmwriteoff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Write Off"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
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
      Left            =   735
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
      Top             =   2760
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
      Left            =   375
      Picture         =   "frmwriteoff.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   2760
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
      Picture         =   "frmwriteoff.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtsup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pembayaran dengan base currency"
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   945
      Width           =   1335
   End
   Begin TDBText6Ctl.TDBText txtbukti 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmwriteoff.frx":0630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":069C
      Key             =   "frmwriteoff.frx":06BA
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
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Format          =   134742019
      CurrentDate     =   37421
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":06F6
      Caption         =   "frmwriteoff.frx":0716
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":0782
      Keys            =   "frmwriteoff.frx":07A0
      Spin            =   "frmwriteoff.frx":07E2
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   4095
      Left            =   0
      TabIndex        =   7
      Top             =   1635
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   8
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   225
      TabIndex        =   8
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Customer"
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
      MICON           =   "frmwriteoff.frx":080A
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
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   6060
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmwriteoff.frx":0B24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   6060
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmwriteoff.frx":0E3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   6375
      TabIndex        =   11
      Top             =   6060
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmwriteoff.frx":1158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Currency"
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
      MICON           =   "frmwriteoff.frx":1472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtsisa 
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":178C
      Caption         =   "frmwriteoff.frx":17AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":1818
      Keys            =   "frmwriteoff.frx":1836
      Spin            =   "frmwriteoff.frx":1878
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai2 
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":18A0
      Caption         =   "frmwriteoff.frx":18C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":192C
      Keys            =   "frmwriteoff.frx":194A
      Spin            =   "frmwriteoff.frx":198C
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
      ValueVT         =   1991573509
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai3 
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":19B4
      Caption         =   "frmwriteoff.frx":19D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":1A40
      Keys            =   "frmwriteoff.frx":1A5E
      Spin            =   "frmwriteoff.frx":1AA0
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
      ValueVT         =   1991573509
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai4 
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":1AC8
      Caption         =   "frmwriteoff.frx":1AE8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":1B54
      Keys            =   "frmwriteoff.frx":1B72
      Spin            =   "frmwriteoff.frx":1BB4
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai5 
      Height          =   255
      Left            =   7320
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmwriteoff.frx":1BDC
      Caption         =   "frmwriteoff.frx":1BFC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":1C68
      Keys            =   "frmwriteoff.frx":1C86
      Spin            =   "frmwriteoff.frx":1CC8
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
      ValueVT         =   1991573509
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   4410
      TabIndex        =   27
      Top             =   465
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmwriteoff.frx":1CF0
      Caption         =   "frmwriteoff.frx":1D10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmwriteoff.frx":1D7C
      Keys            =   "frmwriteoff.frx":1D9A
      Spin            =   "frmwriteoff.frx":1DDC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
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
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4095
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblsisa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Sisa : 0.00"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   6165
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   3225
      TabIndex        =   24
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "No. Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Bayar : 0.00"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5925
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblapply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Apply : 0.00"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6165
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lblbayar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Bayar Apply : 0.00"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   5925
      Width           =   3015
   End
   Begin VB.Label lblsup 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Top             =   840
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5850
      Width           =   6060
   End
End
Attribute VB_Name = "frmwriteoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim str2, posrow As String

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        If lblbase = "1" Then
            txtketerangan = "Piutang tak tertagih"
            Check1.Value = 0
        Else
            txtketerangan = "Piutang tak tertagih dengan base currency"
        End If
    Else
        txtketerangan = "Piutang tak tertagih"
    End If
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If

    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not add, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close

    str2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        If grid2.TextMatrix(grid2.Row, 3) <> "0.00" Then
            str2 = 1
            Exit Do
        End If
        grid2.Row = grid2.Row + 1
    Loop

    If str2 = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do

        If grid2.TextMatrix(grid2.Row, 3) <> "0.00" Then
            OBJ.Open dsn
            SQL = "select * from am_aropnfil where noapply = '" & grid2.TextMatrix(grid2.Row, 1) & "' and transtype <> 'PM'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                MsgBox "Data Entry Not Complete, please refresh customer.", vbExclamation, "Warning"
                OBJ.Close
                Exit Sub
            End If
            OBJ.Close
        End If

        grid2.Row = grid2.Row + 1
    Loop

    OBJ.Open dsn
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click

        Exit Sub
    End If
    
    SQL = "select * from am_aropnfil where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    OBJ.Close

    OBJ.Open dsn
    SQL = "INSERT INTO AM_CashHdr"
    SQL = SQL + " (Kodecust"
    SQL = SQL + ", NoBkt"
    SQL = SQL + ", TglBkt"
    SQL = SQL + ", kodebayar"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", Amount"
    SQL = SQL + ", noac"
    SQL = SQL + ", kodecol"
    SQL = SQL + ", Posted"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", IdEntry"
    SQL = SQL + ", DateEntry"
    SQL = SQL + ", IdUpdate"
    SQL = SQL + ", DateUpdate)"

    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", 'PM'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ",Convert(Money," & hitbayar & ")"
    SQL = SQL + ", '0'"
    SQL = SQL + ", 'CL-001'"
    SQL = SQL + ", '" & Check1.Value & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert(Money," & txtnilaikurs & ")"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", '1'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)

    SQL = "INSERT INTO AM_Cashsub"
    SQL = SQL + " (NoBkt"
    SQL = SQL + ", tglbkt"
    SQL = SQL + ", typeBayar"
    SQL = SQL + ", Kodecust"
    SQL = SQL + ", Nogiro"
    SQL = SQL + ", tgljt"
    SQL = SQL + ", tglcair"
    SQL = SQL + ", tgltolak"
    SQL = SQL + ", bank"
    SQL = SQL + ", acbank"
    SQL = SQL + ", jumlah)"
    
    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", 'TN'"
    SQL = SQL + ", '" & txtsup & "'"
    SQL = SQL + ", ' '"
    SQL = SQL + ",convert(datetime,'01/01/1900')"
    SQL = SQL + ",convert(datetime,'01/01/1900')"
    SQL = SQL + ",convert(datetime,'01/01/1900')"
    SQL = SQL + ", ' '"
    SQL = SQL + ", ' '"
    SQL = SQL + ",Convert(Money," & hitbayar & "))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid2.Row = 1
    OBJ.Open dsn
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do

        If grid2.TextMatrix(grid2.Row, 3) <> "0.00" Then

            SQL = "INSERT INTO AM_CashLin"
            SQL = SQL + " (NoBkt"
            SQL = SQL + ", tglbkt"
            SQL = SQL + ", KodeBayar"
            SQL = SQL + ", Kodecust"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", nilaikurs"
            SQL = SQL + ", jumlah"
            SQL = SQL + ", selisih"
            SQL = SQL + ", potongan)"

            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 1) & "'"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 6), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 3), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 5), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 4), "general number") & "'))"
            Set RST = OBJ.Execute(SQL)

            SQL = "INSERT INTO AM_Aropnfil"
            SQL = SQL + " (KodeCust"
            SQL = SQL + ", NoBkt"
            SQL = SQL + ", TglBkt"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", TransType"
            SQL = SQL + ", JatuhTempo"
            SQL = SQL + ", Keterangan"
            SQL = SQL + ", kodecur"
            SQL = SQL + ", nilaikurs"
            SQL = SQL + ", Amount"
            SQL = SQL + ", Potongan"
            SQL = SQL + ", selisih"
            SQL = SQL + ", PPN)"
        
            SQL = SQL + "VALUES"
            SQL = SQL + " ('" & txtsup & "'"
            SQL = SQL + ", '" & txtbukti & "'"
            SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 1) & "'"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
            SQL = SQL + ", '" & txtketerangan & "'"
            SQL = SQL + ", '" & txtkurs & "'"
            SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 3), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 4), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 5), "General number") & "')"
            SQL = SQL + ",Convert (Money, '0'))"
            Set RST = OBJ.Execute(SQL)
        End If
        grid2.Row = grid2.Row + 1
        DoEvents
    Loop
    OBJ.Close
    
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
    SQL = SQL + "reconsil, "
    SQL = SQL + "lineitem)"
                
    SQL = SQL + " values"
    SQL = SQL + "('01',"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "'PT'," 'PIUTANG TAK TERTAGIH
    SQL = SQL + "'" & txtbukti & "',"
    SQL = SQL + "convert(money,'1'),"
    SQL = SQL + "'69010000',"
    SQL = SQL + "'Penghapusan piutang tak tertagih (Tunai)',"
    SQL = SQL + "'D',"
    SQL = SQL + "Convert(Money," & hitbayar & "),"
    SQL = SQL + "Convert(Money," & hitbayar & "),"
    SQL = SQL + "'" & txtkurs & "',"
    SQL = SQL + "'P',"
    SQL = SQL + "'J',"
    SQL = SQL + "'0',"
    SQL = SQL + "'',"
    SQL = SQL + "'auto',"
    SQL = SQL + "'',"
    SQL = SQL + "Convert(dateTime, '" & tanggal1 & "'),"
    SQL = SQL + "Convert(dateTime, '" & tanggal1 & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "'1')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

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
    SQL = SQL + "reconsil, "
    SQL = SQL + "lineitem)"
        
    SQL = SQL + " values"
    SQL = SQL + "('01',"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "'PT',"
    SQL = SQL + "'" & txtbukti & "',"
    SQL = SQL + "convert(money,'1'),"
        OBJ2.Open dsn
        SQL2 = "Select noac From gl_masterac Where nmac like '%" & txtsup & "%'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            SQL = SQL + "'" & RST2!noac & "',"
        Else
            SQL = SQL + "'',"
        End If
        OBJ2.Close

        OBJ2.Open dsn
        SQL2 = "select kodecust,namacust from am_customer where kodecust='" & txtsup & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            SQL = SQL + "'Penghapusan piutang tak tertagih - " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
        Else
            SQL = SQL + "'Penghapusan piutang tak tertagih - Customer',"
        End If
        OBJ2.Close
        
    SQL = SQL + "'K',"
    SQL = SQL + "Convert(Money," & hitbayar & "),"
    SQL = SQL + "Convert(Money," & hitbayar & "),"
    SQL = SQL + "'" & txtkurs & "',"
    SQL = SQL + "'P',"
    SQL = SQL + "'J',"
    SQL = SQL + "'0',"
    SQL = SQL + "'',"
    SQL = SQL + "'auto',"
    SQL = SQL + "'',"
    SQL = SQL + "Convert(dateTime, '" & tanggal1 & "'),"
    SQL = SQL + "Convert(dateTime, '" & tanggal1 & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "'2')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
' simpan tracking_histori (belum)
'Print Bukti WO
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtbukti = ""
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    txtsup = ""
    lblsup = ""
    txtkurs = ""
    Check1.Value = 0
    txtketerangan = ""
    txtsisa = 0
    hapusgrid
    txtbukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If txtbukti = "" Then Exit Sub
    
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtsup = hasil
    lblsup = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    caripiutang
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs = hasil
    carikurs
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub
Private Sub caripiutang()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    hapusgrid
    grid2.Row = 1
    OBJ.Open dsn

    If lblbase = "1" Then
        SQL = "select a.NoApply, sum(((a.Amount + a.potongan + a.PPN + a.selisih)* a.nilaikurs)-isnull(b.nilaikurs,0)) as Total from AM_Aropnfil a left join am_cashlin b on a.nobkt=b.nobkt and a.noapply=b.noapply WHERE a.kodecust = '" & txtsup & "' and a.tglbkt <= '" & tanggal1 & "' group by a.Noapply order by a.noapply asc"
    Else
        SQL = "select NoApply, sum(Amount + potongan + PPN + selisih) as Total from AM_Aropnfil WHERE kodecust = '" & txtsup & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' group by Noapply order by noapply asc"
    End If
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do Until RST.EOF
            'If Round(RST!total, 0) = 0 Then
            If RST!total = 0 Then
                RST.MoveNext
                GoTo jump2
            End If

            grid2.TextMatrix(grid2.Row, 1) = RST!noapply
            grid2.TextMatrix(grid2.Row, 2) = Format(RST!total, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 3) = "0.00"
            grid2.TextMatrix(grid2.Row, 4) = "0.00"
            grid2.TextMatrix(grid2.Row, 5) = "0.00"
            grid2.TextMatrix(grid2.Row, 7) = Format(RST!total, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 6) = "0.00"
            grid2.Col = 0
            Set grid2.CellPicture = uncheck
            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump2:
        Loop
        OBJ.Close
    Else
        MsgBox "No Transaction For Write Off.", vbInformation, "Information"
        OBJ.Close
        cmdclear_Click
    End If
End Sub

Private Sub date1_Change()
    caripiutang
End Sub

Private Sub Form_Load()
    date1 = Date
'klik checklist nilai bayar otomatis terisi
    grid2.TextMatrix(0, 0) = "X"
    grid2.TextMatrix(0, 1) = "No Apply"
    grid2.TextMatrix(0, 2) = "Piutang"
    grid2.TextMatrix(0, 3) = "Nilai Bayar"
    grid2.TextMatrix(0, 4) = "Disc Bayar"
    grid2.TextMatrix(0, 5) = "Selisih"
    grid2.TextMatrix(0, 7) = "Sisa Piutang"
    grid2.TextMatrix(0, 6) = "Selisih Kurs"

    grid2.ColWidth(0) = 600
    grid2.ColWidth(1) = 1150
    grid2.ColWidth(2) = 1700
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 0
    grid2.ColWidth(5) = 0
    grid2.ColWidth(6) = 0
    grid2.ColWidth(7) = 1700

    grid2.RowHeightMin = 300
End Sub

Private Sub grid2_Click()
If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Then Exit Sub
    posrow = grid2.Row
    
    Select Case grid2.Col
    Case 0
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
        If grid2.CellPicture = uncheck Then
            Set grid2.CellPicture = check
            grid2.TextMatrix(grid2.Row, 3) = Format((Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 5), "general number"))), "##,###,###,###,##0.00")
            If grid2.TextMatrix(grid2.Row, 3) < 0 Then grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row, 3) * -1
            
            lbltotal = "Total Bayar : " & Format(hitbayar, "###,###,##0.00")
            lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
            lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
            
            txtsisa = hitbayar2 - hitbayar
            If lblbase = "0" Then hitselisihkurs
            grid2.TextMatrix(posrow, 7) = Format((Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") - Format(grid2.TextMatrix(posrow, 4), "general number") + Format(grid2.TextMatrix(posrow, 5), "general number")), "###,###,###,##0.00")
        Else
            Set grid2.CellPicture = uncheck
            grid2.TextMatrix(grid2.Row, 3) = "0.00"
            
            lbltotal = "Total Bayar : " & Format(hitbayar, "###,###,##0.00")
            lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
            lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
            
            txtsisa = hitbayar2 - hitbayar
            If lblbase = "0" Then hitselisihkurs
            grid2.TextMatrix(posrow, 7) = Format((Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") - Format(grid2.TextMatrix(posrow, 4), "general number") + Format(grid2.TextMatrix(posrow, 5), "general number")), "###,###,###,##0.00")
        End If
    Case 2
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 5), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 3
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 5), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 4
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
    
    If Len(txtbukti) = 0 And Not (KeyAscii = 87) Then
        KeyAscii = 0
    ElseIf Len(txtbukti) > 0 And Len(txtbukti) <= 5 And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then
        KeyAscii = 0
    ElseIf KeyAscii = 87 Then
        OBJ.Open dsn
        SQL = "select max(nobkt)'no' from am_cashhdr where kodebayar='PM' and nobkt like 'W%'"
        Set RST = OBJ.Execute(SQL)
        If Len(Mid(RST!no, 2, 5) + 1) = 1 Then
            txtbukti = txtbukti & "0000" & Mid(RST!no, 2, 5) + 1
        ElseIf Len(Mid(RST!no, 2, 5) + 1) = 2 Then
            txtbukti = txtbukti & "000" & Mid(RST!no, 2, 5) + 1
        ElseIf Len(Mid(RST!no, 2, 5) + 1) = 3 Then
            txtbukti = txtbukti & "00" & Mid(RST!no, 2, 5) + 1
        ElseIf Len(Mid(RST!no, 2, 5) + 1) = 4 Then
            txtbukti = txtbukti & "0" & Mid(RST!no, 2, 5) + 1
        ElseIf Len(Mid(RST!no, 2, 5) + 1) = 5 Then
            txtbukti = txtbukti & Mid(RST!no, 2, 5) + 1
        Else
            txtbukti = txtbukti & "00001"
        End If
        OBJ.Close
    End If
End Sub

Private Sub hapusgrid()
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

        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.ColWidth(0) = 600
    grid2.ColWidth(1) = 1150
    grid2.ColWidth(2) = 1700
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 0
    grid2.ColWidth(5) = 0
    grid2.ColWidth(6) = 0
    grid2.ColWidth(7) = 1700
    grid2.Col = 0
    Set grid2.CellPicture = blank
    lblapply = "Total Apply : 0.00"
    lblbayar = "Bayar Apply : 0.00"
    lbltotal = "Total Bayar : 0.00"
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub txtbukti_LostFocus()
    If txtbukti = "" Then Exit Sub

    hapusgrid

    OBJ.Open dsn
    SQL = "Select * From AM_CashHdr Where NoBkt = '" & txtbukti & "' And kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Pembayaran already exist.", vbInformation, "Information"
        
        If Date > date1.MaxDate Then
            date1 = date1.MaxDate
        ElseIf Date < date1.MinDate Then
            date1 = date1.MinDate
        Else
            date1.Value = Date
        End If
        txtbukti = ""
        txtsup = ""
        lblsup = ""
        Check1.Value = 0
        txtketerangan = "Piutang tak tertagih"
        txtkurs = ""
        date1 = Date
        txtbukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkurs_Change()
    txtsup = ""
    lblsup = ""
    Check1.Value = 0
    txtketerangan = "Piutang tak tertagih"
    hapusgrid
    txtbukti.SetFocus
End Sub
Private Sub txtkurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkurs_LostFocus
    KeyAscii = 0
End Sub
Private Sub txtkurs_LostFocus()
    carikurs
End Sub

Private Sub carikurs()
    If txtkurs = "" Then Exit Sub
    OBJ2.Open dsn
    SQL2 = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If Not RST2.EOF Then
        If RST2!base = 1 Then
            lblbase = "1"
        Else
            lblbase = "0"
        End If
        Select Case Month(date1)
        Case 1
            txtnilaikurs = RST2!kurs1
        Case 2
            txtnilaikurs = RST2!kurs2
        Case 3
            txtnilaikurs = RST2!kurs3
        Case 4
            txtnilaikurs = RST2!kurs4
        Case 5
            txtnilaikurs = RST2!kurs5
        Case 6
            txtnilaikurs = RST2!kurs6
        Case 7
            txtnilaikurs = RST2!kurs7
        Case 8
            txtnilaikurs = RST2!kurs8
        Case 9
            txtnilaikurs = RST2!kurs9
        Case 10
            txtnilaikurs = RST2!kurs10
        Case 11
            txtnilaikurs = RST2!kurs11
        Case 12
            txtnilaikurs = RST2!kurs12
        End Select
    Else
        MsgBox "Currency " & txtkurs & " Not Found.", vbInformation, "Information"
        txtkurs = ""
        txtkurs.SetFocus
    End If
    OBJ2.Close
End Sub

Function hitbayar()
    hitbayar = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 2) = "" Then Exit Do
        hitbayar = Val(hitbayar) + Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number"))

        grid2.Row = grid2.Row + 1
    Loop
End Function

Function hitbayar2()
    hitbayar2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 2) = "" Then Exit Do
        hitbayar2 = Val(hitbayar2) + Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number"))

        grid2.Row = grid2.Row + 1
    Loop
End Function
Private Sub hitselisihkurs()
    OBJ.Open dsn
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 1) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype<>'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai2 = RST!total
    Else
        txtnilai2 = 0
    End If
    
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 1) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype='PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai3 = RST!total * -1
    Else
        txtnilai3 = 0
    End If
    
    txtnilai4 = Val(Format(grid2.TextMatrix(posrow, 3), "general number")) + Val(Format(grid2.TextMatrix(posrow, 4), "general number")) - Val(Format(grid2.TextMatrix(posrow, 5), "general number"))
    txtnilai5 = txtnilai4 + txtnilai3
    grid2.TextMatrix(posrow, 6) = "0.00"
    If txtnilai2 = txtnilai5 Then
        SQL = "select isnull(sum((Amount + potongan + PPN + selisih)*nilaikurs),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype<>'PM'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtnilai2 = RST!total
        Else
            txtnilai2 = 0
        End If
        
        SQL = "select isnull(sum((Amount + potongan + PPN + selisih)*nilaikurs),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype='PM'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtnilai3 = RST!total * -1
        Else
            txtnilai3 = 0
        End If
        
        txtnilai4 = Val((Format(grid2.TextMatrix(posrow, 3), "general number")) + Val(Format(grid2.TextMatrix(posrow, 4), "general number")) - Val(Format(grid2.TextMatrix(posrow, 5), "general number"))) * txtnilaikurs
        txtnilai5 = txtnilai4 + txtnilai3
        
        If txtnilai2 <> txtnilai5 Then
            grid2.TextMatrix(posrow, 6) = Format(txtnilai2 - txtnilai3 - txtnilai4, "###,###,##0.00")
        End If
    End If
    OBJ.Close
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, grid2.Col) = Format(txtnilai, "###,###,##0.00")
        
        If grid2.Col = 3 Then
            If Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")) < 0 Then
                grid2.SetFocus
                grid2.TextMatrix(grid2.Row, 4) = "0.00"
                txtnilai = 0
                Exit Sub
            End If
        End If
        lbltotal = "Total Bayar : " & Format(hitbayar, "###,###,##0.00")
        lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
        
        txtsisa = hitbayar2 - hitbayar
        If lblbase = "0" Then hitselisihkurs
        grid2.TextMatrix(posrow, 7) = Format((Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") - Format(grid2.TextMatrix(posrow, 4), "general number") + Format(grid2.TextMatrix(posrow, 5), "general number")), "###,###,###,##0.00")
        
        grid2.SetFocus
        grid2.Row = posrow
    End If
    If KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
    End If
End Sub
Private Sub txtsisa_Change()
    lblsisa = " Sisa : " & Format(txtsisa, "###,###,##0.00")
End Sub
Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
    txtnilai = 0
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsup_LostFocus()
    If txtsup = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_customer where kodecust = '" & txtsup & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsup = RST!namacust
        OBJ.Close
        caripiutang
    Else
        MsgBox "Customer " & txtsup & " Not Found.", vbExclamation, "Warning"
        txtsup = ""
        lblsup = ""
        txtsup.SetFocus
        OBJ.Close
        Exit Sub
    End If
End Sub
