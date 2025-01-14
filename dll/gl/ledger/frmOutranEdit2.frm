VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmOutranEdit2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Updating Cash/Bank Out"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8430
   Icon            =   "frmOutranEdit2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   8430
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Picture         =   "frmOutranEdit2.frx":000C
      ScaleHeight     =   1575
      ScaleWidth      =   1200
      TabIndex        =   75
      Top             =   0
      Width           =   1200
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7920
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7920
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   240
      Picture         =   "frmOutranEdit2.frx":62BE
      ScaleHeight     =   1575
      ScaleWidth      =   1200
      TabIndex        =   74
      Top             =   0
      Width           =   1200
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
      TabIndex        =   17
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
      Picture         =   "frmOutranEdit2.frx":C570
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      Picture         =   "frmOutranEdit2.frx":C852
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2380
      Left            =   7920
      TabIndex        =   12
      Top             =   2640
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
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
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
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   7800
      Width           =   855
   End
   Begin Chameleon.chameleonButton cmdSaveP 
      Height          =   495
      Left            =   4920
      TabIndex        =   8
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
      MICON           =   "frmOutranEdit2.frx":CBA0
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
      TabIndex        =   9
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmOutranEdit2.frx":CEBA
      Caption         =   "frmOutranEdit2.frx":CEDA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":CF46
      Keys            =   "frmOutranEdit2.frx":CF64
      Spin            =   "frmOutranEdit2.frx":CFA6
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
      TabIndex        =   10
      Top             =   4560
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
      TabIndex        =   11
      Top             =   4560
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
      Caption         =   "frmOutranEdit2.frx":CFCE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D03A
      Key             =   "frmOutranEdit2.frx":D058
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   6240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2355
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
      TabIndex        =   19
      Top             =   2520
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
      Format          =   108134403
      CurrentDate     =   37694
   End
   Begin TDBText6Ctl.TDBText txtkodecomp 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":D094
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D100
      Key             =   "frmOutranEdit2.frx":D11E
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
      TabIndex        =   20
      Top             =   3000
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":D15A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D1C6
      Key             =   "frmOutranEdit2.frx":D1E4
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
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":D220
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D28C
      Key             =   "frmOutranEdit2.frx":D2AA
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
      TabIndex        =   22
      Top             =   3720
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":D2E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D352
      Key             =   "frmOutranEdit2.frx":D370
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
      TabIndex        =   23
      Top             =   5880
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":D3AC
      Caption         =   "frmOutranEdit2.frx":D3CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":D438
      Keys            =   "frmOutranEdit2.frx":D456
      Spin            =   "frmOutranEdit2.frx":D498
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
      TabIndex        =   24
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
      MICON           =   "frmOutranEdit2.frx":D4C0
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
      TabIndex        =   25
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
      MICON           =   "frmOutranEdit2.frx":D7DA
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
      TabIndex        =   26
      Top             =   480
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":DAF4
      Caption         =   "frmOutranEdit2.frx":DB14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":DB80
      Keys            =   "frmOutranEdit2.frx":DB9E
      Spin            =   "frmOutranEdit2.frx":DBE0
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
      TabIndex        =   27
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":DC08
      Caption         =   "frmOutranEdit2.frx":DC28
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":DC94
      Keys            =   "frmOutranEdit2.frx":DCB2
      Spin            =   "frmOutranEdit2.frx":DCF4
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
      TabIndex        =   28
      Top             =   3720
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
      MICON           =   "frmOutranEdit2.frx":DD1C
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
      TabIndex        =   29
      Top             =   1440
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
      MICON           =   "frmOutranEdit2.frx":E036
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
      TabIndex        =   30
      Top             =   4080
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":E350
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":E3BC
      Key             =   "frmOutranEdit2.frx":E3DA
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
      TabIndex        =   31
      Top             =   4080
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
      MICON           =   "frmOutranEdit2.frx":E416
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
      TabIndex        =   32
      Top             =   4080
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":E730
      Caption         =   "frmOutranEdit2.frx":E750
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":E7BC
      Keys            =   "frmOutranEdit2.frx":E7DA
      Spin            =   "frmOutranEdit2.frx":E81C
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
   Begin TDBText6Ctl.TDBText txtketcash 
      Height          =   285
      Left            =   2640
      TabIndex        =   33
      Top             =   5160
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":E844
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":E8B0
      Key             =   "frmOutranEdit2.frx":E8CE
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
      TabIndex        =   34
      Top             =   5520
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":E90A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":E976
      Key             =   "frmOutranEdit2.frx":E994
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":E9D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":EA3C
      Key             =   "frmOutranEdit2.frx":EA5A
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
      TabIndex        =   35
      Top             =   2160
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
      MICON           =   "frmOutranEdit2.frx":EA96
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
      TabIndex        =   36
      Top             =   5160
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":EDB0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":EE1C
      Key             =   "frmOutranEdit2.frx":EE3A
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
      MaxLength       =   6
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
      TabIndex        =   37
      Top             =   3360
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":EE76
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":EEE2
      Key             =   "frmOutranEdit2.frx":EF00
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
      TabIndex        =   38
      Top             =   5880
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
      Format          =   142213123
      CurrentDate     =   37694
   End
   Begin Chameleon.chameleonButton cmdadd1 
      Height          =   495
      Left            =   5040
      TabIndex        =   39
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
      MICON           =   "frmOutranEdit2.frx":EF3C
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
      TabIndex        =   40
      Top             =   820
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":F256
      Caption         =   "frmOutranEdit2.frx":F276
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":F2E2
      Keys            =   "frmOutranEdit2.frx":F300
      Spin            =   "frmOutranEdit2.frx":F342
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
      TabIndex        =   41
      Top             =   4440
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":F36A
      Caption         =   "frmOutranEdit2.frx":F38A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":F3F6
      Keys            =   "frmOutranEdit2.frx":F414
      Spin            =   "frmOutranEdit2.frx":F456
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
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs2 
      Height          =   285
      Left            =   5880
      TabIndex        =   42
      Top             =   3720
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmOutranEdit2.frx":F47E
      Caption         =   "frmOutranEdit2.frx":F49E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":F50A
      Keys            =   "frmOutranEdit2.frx":F528
      Spin            =   "frmOutranEdit2.frx":F56A
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
      TabIndex        =   43
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":F592
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":F5FE
      Key             =   "frmOutranEdit2.frx":F61C
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
      TabIndex        =   44
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Caption         =   "frmOutranEdit2.frx":F658
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOutranEdit2.frx":F6C4
      Key             =   "frmOutranEdit2.frx":F6E2
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
   Begin VB.Label posted 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POSTED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6120
      TabIndex        =   73
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   71
      Top             =   2520
      Width           =   2415
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
      TabIndex        =   70
      Top             =   2550
      Width           =   1455
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
      TabIndex        =   69
      Top             =   3750
      Width           =   1095
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
      TabIndex        =   68
      Top             =   1440
      Width           =   4935
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
      TabIndex        =   67
      Top             =   8145
      Width           =   4455
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
      TabIndex        =   66
      Top             =   7905
      Width           =   4455
   End
   Begin VB.Label lblbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   65
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   64
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   63
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Updating"
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
      TabIndex        =   62
      Top             =   0
      Width           =   2655
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
      TabIndex        =   61
      Top             =   150
      Width           =   1575
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
      TabIndex        =   60
      Top             =   510
      Width           =   1575
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
      TabIndex        =   59
      Top             =   150
      Width           =   855
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
      TabIndex        =   58
      Top             =   510
      Width           =   855
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
      TabIndex        =   57
      Top             =   7665
      Width           =   4215
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
      TabIndex        =   56
      Top             =   4110
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
      TabIndex        =   55
      Top             =   5190
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
      TabIndex        =   54
      Top             =   5550
      Width           =   1335
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
      TabIndex        =   53
      Top             =   1830
      Width           =   1335
   End
   Begin MSForms.ComboBox cmbdaerah 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
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
      TabIndex        =   52
      Top             =   3030
      Width           =   1455
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
      TabIndex        =   51
      Top             =   3030
      Width           =   3255
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
      Top             =   2850
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
      TabIndex        =   49
      Top             =   3405
      Width           =   1335
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
      TabIndex        =   48
      Top             =   5910
      Width           =   1335
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
      TabIndex        =   47
      Top             =   4440
      Width           =   1095
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
      TabIndex        =   46
      Top             =   840
      Width           =   1575
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
      TabIndex        =   45
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      X1              =   240
      X2              =   7680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   1800
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   1150
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmOutranEdit2"
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

Private test As Boolean
Private poscol As Integer
Private posrow1 As Integer
Dim posrow, str2, str3, str4, str6, str7, str8, str9 As String

Private Sub cmbdaerah_LostFocus()
    If Not (cmbdaerah >= 1 And cmbdaerah <= 4) Then
        cmbdaerah = ""
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        cmbdaerah.SetFocus
    Else
        hapusemua
        'txtnotran.Enabled = True
        'txtkodetran.Enabled = True
        'cmdsearch3.Enabled = True
        
        'txtkodetran = ""
        'txtnotran = ""
    End If
End Sub

Private Sub cmdclear_Click()
    hapusemua
    txtkodecomp = ""
    lblnamacomp = ""
    cmbdaerah = ""
    txtnovoucher = ""
    date1 = Date
    date2 = Date
    txtkodetran = ""
    txtnotran = ""
    txtkpd = ""
    txtketcash = ""
    txtnobp = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    test = True
    grid.TextMatrix(0, 1) = "Account"
    grid.TextMatrix(0, 2) = "Keterangan"
    grid.TextMatrix(0, 3) = "Debet"
    grid.TextMatrix(0, 4) = "IDR"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 4350
    grid.ColWidth(3) = 1650
    grid.ColWidth(4) = 2000
    
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
    List1.AddItem "09-LIPPO SPARTA IDR"
    List1.AddItem "11-BCA Gunarso Dede"
End Sub

Private Sub Timer1_Timer()
    If test = True Then
        If Picture1.Left < 7100 Then
            Picture1.Left = Picture1.Left + 200
        End If
        If Picture1.Left = 7240 And Picture2.Left < 5960 Then
            Picture2.Left = Picture2.Left + 200
        End If
    Else
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    If test = False Then
        If Picture1.Left > 0 Then
            Picture1.Left = Picture1.Left - 200
        Else
            Timer1.Enabled = True
        End If
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
    'txtkodetran = ""
    'txtnotran = ""
    cmbdaerah = ""
    'date1 = Date
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacomp = RST!nmcompscr
        format_coa = RST!formatac
        cmbdaerah.SetFocus
    Else
        MsgBox "Company " & txtkodecomp & " Not Found.", vbInformation, "Information"
        txtkodecomp = ""
        txtkodecomp.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtnovoucher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hapusemua
            OBJ.Open dsn
            SQL = "Select no_voucher From no_bank_payment Where no_voucher = '" & txtnovoucher & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                MsgBox "Data not Found..!", vbCritical, "Warning"
                Exit Sub
            End If

            SQL = "Select ref1 From am_beliapp Where ref1= '" & txtnovoucher & "'"
            Set RST = OBJ.Execute(SQL)
            
            If RST.EOF Then
                Check2.Visible = False
                Check1.Visible = False
                NonBahanBaku
            Else
                SQL = "Select flag From no_bank_payment Where no_voucher= '" & txtnovoucher & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    Check1.Visible = True
                Else
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
            'carikurs
            txtcash.SetFocus
            Timer1.Enabled = True
            'Picture1.Visible = True
    End If
    
End Sub

Private Sub NonBahanBaku()
    SQL = "Select b.flag From gl_transaksi b inner join no_bank_payment a on a.notrx = b.notrx "
    SQL = SQL + "Where b.dbkrtrx = 'K' and a.no_voucher = '" & txtnovoucher & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = "P" Then
            posted.Visible = True
        Else
            posted.Visible = False
        End If
    End If

    SQL = "Select a.keterangan,a.perkiraan,a.jumlah,b.kepada,b.npwp,b.alamat,b.kdkurs,"
    SQL = SQL + "b.nilai,b.ppn "
    SQL = SQL + "From am_voucherin a inner join am_voucherhdr b "
    SQL = SQL + "On a.novoucher=b.novoucher where a.novoucher = '" & txtnovoucher & "'"
    Set RST = OBJ.Execute(SQL)
                
    If Not RST.EOF Then
        
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
                If Len(RST!keterangan) > 60 Then
                    MsgBox "Data tidak dapat ditampilkan" + Chr(13) + "Keterangan dengan Account, " _
                    + RST!perkiraan + " terlalu panjang", vbExclamation, "WARNING !"
                   ' cmdclear_Click
                    Exit Sub
                End If
                grid.TextMatrix(grid.Row, 2) = RST!keterangan
                grid.TextMatrix(grid.Row, 3) = Format(RST!jumlah, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 4) = Format(txtnilaikurs2, "###,###,##0.00") * Format(RST!jumlah, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 4) = Format(grid.TextMatrix(grid.Row, 4), "###,###,##0.00")
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
            End With
            RST.MoveNext
            DoEvents
        Loop
        
'        SQL = "Select b.amounttrx From gl_transaksi b inner join no_bank_payment a on a.notrx = b.notrx "
'        SQL = SQL + "Where b.dbkrtrx = 'K' and a.no_voucher = '" & txtnovoucher & "'"
'        Set RST = OBJ.Execute(SQL)

'        If Not RST.EOF Then
'            txtncash.text = Format(RST!amounttrx, "###,###,##0.00")
'            txtdebet = Format(RST!amounttrx, "###,###,##0.00")
'            txtpembulatan = Format(RST!amounttrx, "###,###,##0.00")
'            MsgBox "NBB"
'        End If
        If txtkodecur <> "IDR" Then
            Check1.Visible = True
        End If
        SQL = "Select a.kdcomp,a.tgltrx,a.kdtrx,a.notrx,a.kurs,a.desctrx,a.noactrx,a.cekbg,b.no_payment,b.kpd,b.tgljt,b.is_pajak From gl_transaksi a "
        SQL = SQL + "Inner join no_bank_payment b on a.notrx=b.notrx "
        SQL = SQL + "Where a.dbkrtrx = 'K' and b.no_voucher = '" & txtnovoucher & "'"
        Set RST = OBJ.Execute(SQL)
            
            date1.Value = RST!tgltrx
            txtkodetran.text = RST!kdtrx
            txtnotran.text = RST!notrx
            txtnilaikurs2 = RST!kurs
            txtketcash.text = RST!desctrx
            txtcash.text = RST!noactrx
            txtcekbg.text = RST!cekbg
            txtkpd.text = RST!kpd
            date2.Value = RST!tgljt
            txtnobp.text = RST!no_payment
            

            'If RST!is_pajak = "1" Then
            '    optpajak.Value = True
            '    optnonpajak.Value = False
            'Else
            '    optpajak.Value = False
            '    optnonpajak.Value = True
            'End If

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

Private Sub txtnovoucher_LostFocus()
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
    carikurs
    
        OBJ.Open dsn
        SQL = "Select b.amounttrx From gl_transaksi b inner join no_bank_payment a on a.notrx = b.notrx "
        SQL = SQL + "Where b.dbkrtrx = 'K' and a.no_voucher = '" & txtnovoucher & "'"
        Set RST = OBJ.Execute(SQL)

        If Not RST.EOF Then
            txtncash.text = Format(RST!amounttrx, "###,###,##0.00")
            txtdebet = Format(RST!amounttrx, "###,###,##0.00")
            txtpembulatan = Format(RST!amounttrx, "###,###,##0.00")
        End If
        OBJ.Close
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

Private Sub hapusemua()
    txtkodecur = ""
    lblnamacur = "Currency :"
    txtcash = ""
    lblcash = "Cash/Bank : "
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
    
    optnonpajak.Value = False
    optpajak.Value = False
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
    grid.ColWidth(4) = 2000
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

Private Sub optnonpajak_Click()
If MsgBox("Apakah anda yakin, " + Chr(13) + _
    "Ingin mengganti dengan No.Payment Non Pajak yg baru ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    txtnobp = Getnobpnonpajak
    txtketcash.SetFocus
End Sub

Private Sub optpajak_Click()
If MsgBox("Apakah anda yakin, " + Chr(13) + _
    "Ingin mengganti dengan No.Payment Pajak yg baru ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    txtnobp = GetnobpPajak
    txtketcash.SetFocus
End Sub

Function Getnobpnonpajak() As String
On Error GoTo err_handler:
    Dim i As Long
    Dim strFormat As String
    Dim int_kode As Integer
    Dim Temp_kode As String
    OBJ.Open dsn
    SQL = "Select max(no_payment) as nobp From no_bank_payment "
    SQL = SQL + "Where no_payment like '" + strFormat + "%' and is_pajak ='0'"
    Set RST = OBJ.Execute(SQL)
    If RST!nobp = "" Then
        Temp_kode = "K00001"
    End If
    If RST!nobp <> "" Then
        int_kode = Right(RST!nobp, 5)
        int_kode = int_kode + 1
        
        If (Len(Trim(Str(int_kode))) = 1) Then
            Temp_kode = int_kode
            Temp_kode = "K0000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 2) Then
            Temp_kode = int_kode
            Temp_kode = "K000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 3) Then
            Temp_kode = int_kode
            Temp_kode = "K00" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 4) Then
            Temp_kode = int_kode
            Temp_kode = "K0" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 5) Then
            Temp_kode = int_kode
            Temp_kode = "K" + Trim(Str(int_kode))
        End If
    End If
    Getnobpnonpajak = Temp_kode
    OBJ.Close
    Exit Function
err_handler:
    Getnobpnonpajak = "K00001"
    OBJ.Close

End Function

Function GetnobpPajak()
On Error GoTo err_handler:
    Dim i As Long
    Dim int_kode As Integer
    Dim Temp_kode As String
    OBJ.Open dsn
    SQL = "Select max(no_payment) as nobp From no_bank_payment Where is_pajak ='1'"
    Set RST = OBJ.Execute(SQL)
    If RST!nobp = "" Then
        Temp_kode = "00001"
    End If
    If RST!nobp <> "" Then
        int_kode = RST!nobp ' + 1
        int_kode = int_kode + 1
        
        If (Len(Trim(Str(int_kode))) = 1) Then
            Temp_kode = "0000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 2) Then
            Temp_kode = "000" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 3) Then
            Temp_kode = "00" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 4) Then
            Temp_kode = "0" + Trim(Str(int_kode))
        End If
        If (Len(Trim(Str(int_kode))) = 5) Then
            Temp_kode = Trim(Str(int_kode))
        End If
    End If
    GetnobpPajak = Temp_kode
    OBJ.Close
    Exit Function
err_handler:
    GetnobpPajak = "00001"
    OBJ.Close
End Function

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

Private Sub txtncash_Change()
'On Error Resume Next
    txtkredit = txtncash
    lblcomkredit = "1 Lines"
    txtpembulatan = txtncash * txtnilaikurs
    txtpembulatan = SpyRound(txtpembulatan)
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

Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function

Private Function SpyRoundUp(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.1) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRoundUp = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRoundUp = Val(arVal(0)) Else: SpyRoundUp = Val(arVal(0)) + 1
End Function
