VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmpayap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Bayar Hutang"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   ClipControls    =   0   'False
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
   Icon            =   "frmpayap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   7800
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmpayap.frx":2372
      Caption         =   "frmpayap.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":23FE
      Keys            =   "frmpayap.frx":241C
      Spin            =   "frmpayap.frx":245E
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
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "frmpayap.frx":2486
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":24F2
      Key             =   "frmpayap.frx":2510
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
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   0
         TabIndex        =   7
         Top             =   -45
         Width           =   975
         VariousPropertyBits=   612386843
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1720;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1200
      Width           =   6255
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   1560
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
      Format          =   135004163
      CurrentDate     =   38767
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
      Left            =   6360
      Picture         =   "frmpayap.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   480
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
      Left            =   6120
      Picture         =   "frmpayap.frx":2902
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   255
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
      Left            =   6600
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin TDBText6Ctl.TDBText txtbukti 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmpayap.frx":2CB8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":2D24
      Key             =   "frmpayap.frx":2D42
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
   Begin VB.TextBox txtsup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
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
      Format          =   135004163
      CurrentDate     =   37421
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmpayap.frx":2D7E
      Caption         =   "frmpayap.frx":2D9E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":2E0A
      Keys            =   "frmpayap.frx":2E28
      Spin            =   "frmpayap.frx":2E6A
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Supplier"
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
      MICON           =   "frmpayap.frx":2E92
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
      Left            =   9840
      TabIndex        =   15
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayap.frx":31AC
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
      Left            =   8880
      TabIndex        =   14
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayap.frx":34C6
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
      Left            =   7920
      TabIndex        =   13
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayap.frx":37E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   3625
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
   Begin TDBNumber6Ctl.TDBNumber txtsisa 
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmpayap.frx":3AFA
      Caption         =   "frmpayap.frx":3B1A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":3B86
      Keys            =   "frmpayap.frx":3BA4
      Spin            =   "frmpayap.frx":3BE6
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmpayap.frx":3C0E
      Caption         =   "frmpayap.frx":3C2E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":3C9A
      Keys            =   "frmpayap.frx":3CB8
      Spin            =   "frmpayap.frx":3CFA
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   31
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
      MICON           =   "frmpayap.frx":3D22
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   255
      Left            =   10080
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmpayap.frx":403C
      Caption         =   "frmpayap.frx":405C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":40C8
      Keys            =   "frmpayap.frx":40E6
      Spin            =   "frmpayap.frx":4128
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
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   255
      Left            =   9360
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmpayap.frx":4150
      Caption         =   "frmpayap.frx":4170
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayap.frx":41DC
      Keys            =   "frmpayap.frx":41FA
      Spin            =   "frmpayap.frx":423C
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
   Begin VB.Label Label2 
      Caption         =   "Note : Untuk Pembayaran memakai GIRO, 1 GIRO adalah 1 NO BAYAR."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   6120
      Width           =   7575
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   510
      Width           =   855
   End
   Begin VB.Label lblbayar 
      Caption         =   "Bayar Apply : 0.00"
      Height          =   255
      Left            =   7920
      TabIndex        =   29
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lblapply 
      Caption         =   "Total Apply : 0.00"
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   120
      Width           =   2895
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
      Left            =   3120
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbltotal 
      Caption         =   "Total Bayar : 0.00"
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "No. Bayar"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Bayar"
      Height          =   255
      Left            =   3225
      TabIndex        =   17
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblsup 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label lblsisa 
      Caption         =   "Sisa Bayar : 0.00"
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frmpayap"
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

Dim cmd As New ADODB.Command
Dim vcmd(0) As Variant

Dim str2 As String
Dim posrow As String
Dim i As Integer

Private Sub cmbtype_Click()
    If cmbtype = "" Then Exit Sub
    
    Grid1.Row = 1
    Do While cmbtype = "Tunai"
        If Grid1.Row = Grid1.Rows - 1 Then Exit Do
        
        If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then
            cmbtype = ""
            Frame1.Visible = False
            Exit Sub
        End If
        
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Row = posrow
    
    Grid1.SetFocus
    Grid1.TextMatrix(Grid1.Row, 1) = cmbtype
    Grid1.TextMatrix(Grid1.Row, 6) = "0.00"
    Grid1.TextMatrix(Grid1.Row, 7) = "0.00"
    cmbtype = ""
    Frame1.Visible = False
    
    Grid1.Col = 0
    Set Grid1.CellPicture = uncheck.Picture
                        
    If Grid1.Row = (Grid1.Rows - 1) Then Grid1.Rows = Grid1.Rows + 1
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then cmbtype_LostFocus
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    Frame1.Visible = False
End Sub

Private Sub cmdadd_Click()
Dim Nkurs As String
Dim Rkurs As String
    Rkurs = txtnilaikurs
    If txtkurs <> "IDR" Then
        Nkurs = InputBox("Masukkan nilaikurs !", "Konfirmasi Nilai Kurs", 0)
            If Nkurs = "" Then
                txtnilaikurs = Rkurs
                Exit Sub
            Else
                txtnilaikurs = Nkurs
            End If
    End If

    txtbukti = Trim(txtbukti)
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If Asc(Right(txtbukti, 1)) >= 65 And Asc(Right(txtbukti, 1)) <= 90 Then
        If MsgBox("Apakah ini pembayaran cicil ?", vbYesNo + vbQuestion, "Question") = vbNo Then
            txtbukti.SetFocus
            Exit Sub
        End If
    End If
    
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtsisa <> 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid2.Rows = 2 Or Grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Grid1.Rows - 2 > 1 Then
        Grid1.Row = 1
        Do While True
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
            
            If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" Then
                MsgBox "Untuk Pembayaran memakai GIRO, 1 GIRO adalah 1 NO BAYAR.", vbExclamation, "Warning"
                Exit Sub
            End If
            
            Grid1.Row = Grid1.Row + 1
        Loop
    End If
    
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
                    
        If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" And Grid1.TextMatrix(Grid1.Row, 2) = "" Then
            MsgBox "Data Entry Not Complite. No Giro Harus Diisi..!", vbExclamation, "Warning"
            Exit Sub
        End If
                    
        If Grid1.TextMatrix(Grid1.Row, 1) <> "Tunai" And Grid1.TextMatrix(Grid1.Row, 3) = "" Then
            MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If Grid1.TextMatrix(Grid1.Row, 1) <> "Tunai" And Grid1.TextMatrix(Grid1.Row, 4) = "" Then
            MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If Grid1.TextMatrix(Grid1.Row, 6) = "0.00" And Grid1.TextMatrix(Grid1.Row, 7) = "0.00" Then
            MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If Grid1.TextMatrix(Grid1.Row, 6) = "0.00" And Grid1.TextMatrix(Grid1.Row, 7) <> "0.00" Then
            MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
            Exit Sub
        End If

        Grid1.Row = Grid1.Row + 1
    Loop
    
    str2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            str2 = 1
            Exit Do
        End If
        grid2.Row = grid2.Row + 1
    Loop
    
    If str2 = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
            
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            OBJ.Open dsn
            SQL = "select * from am_apopnfil where noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "' and transtype <> 'PM'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                MsgBox "Data Entry Not Complite, please refresh supplier.", vbExclamation, "Warning"
                OBJ.Close
                Exit Sub
            End If
            OBJ.Close
        End If
        
        grid2.Row = grid2.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_apcashhdr where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    
    SQL = "select * from am_apopnfil where nobeli = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "Select * From am_apcashhdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kodesupp = txtsup
        !nobkt = txtbukti
        !tglbkt = Format(date1, "yyyy/MM/dd")
        !kodebayar = "PM"
        !noapply = txtbukti
        !keterangan = txtketerangan
        !amount = hitbayar
        !posted = "0"
        !identry = UserOnline
        !dateentry = tanggal2
        !idupdate = ""
        !DateUpdate = tanggalsekarang
        !kodecur = txtkurs
        !nilaikurs = txtnilaikurs
        .Update
    End With
    OBJ.Close
    
    Grid1.Row = 1
    OBJ.Open dsn
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        
        SQL = "INSERT INTO AM_apCashsub"
        SQL = SQL + " (Nobukti"
        SQL = SQL + ", tglbukti"
        SQL = SQL + ", type"
        SQL = SQL + ", Kodesupp"
        SQL = SQL + ", Nogiro"
        SQL = SQL + ", tgljt"
        SQL = SQL + ", tanggalcair"
        SQL = SQL + ", tanggaltolak"
        SQL = SQL + ", bank"
        SQL = SQL + ", acbank"
        SQL = SQL + ", byadmin"
        SQL = SQL + ", jumlah)"
        
        SQL = SQL + " VALUES"
        SQL = SQL + " ('" & txtbukti & "'"
        SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then SQL = SQL + ", 'TN'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Cek" Then SQL = SQL + ", 'C'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" Then SQL = SQL + ", 'G'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" Then SQL = SQL + ", 'TF'"
        SQL = SQL + ", '" & txtsup & "'"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
        If Grid1.TextMatrix(Grid1.Row, 3) = "" Then SQL = SQL + ",convert(datetime,'01/01/1900')"
        If Grid1.TextMatrix(Grid1.Row, 3) <> "" Then SQL = SQL + ",convert(datetime,'" & tanggalgrid & "')"
        SQL = SQL + ",convert(datetime,'01/01/1900')"
        SQL = SQL + ",convert(datetime,'01/01/1900')"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
        SQL = SQL + ",convert(money,'" & Format(Grid1.TextMatrix(Grid1.Row, 7), "general number") & "')"
        SQL = SQL + ",convert(money,'" & Format(Grid1.TextMatrix(Grid1.Row, 6), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        Grid1.Row = Grid1.Row + 1
    Loop
    OBJ.Close
    
    grid2.Row = 1
    OBJ.Open dsn
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            
            SQL = "INSERT INTO AM_apCashLin"
            SQL = SQL + " (Nobkt"
            SQL = SQL + ", tglbkt"
            SQL = SQL + ", kodebayar"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", kodesupp"
            SQL = SQL + ", jumlah"
            SQL = SQL + ", selisih"
            SQL = SQL + ", selisihkurs"
            SQL = SQL + ", potongan)"
            
            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 4), "general number") & "')"
            SQL = SQL + ",convert(money,'" & cariselisih & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 3), "general number") & "'))"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "INSERT INTO AM_Apopnfil"
            SQL = SQL + " (Kodesupp"
            SQL = SQL + ", NoBeli"
            SQL = SQL + ", TglBeli"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", TransType"
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
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ", '" & txtketerangan & "'"
            SQL = SQL + ", '" & txtkurs & "'"
            SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 2), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 3), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 4), "General number") & "')"
            SQL = SQL + ",Convert (Money, 0))"
            Set RST = OBJ.Execute(SQL)
        End If
        grid2.Row = grid2.Row + 1
    Loop
    OBJ.Close
    
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
    txtnilaikurs = 0
    lblbase = ""
    txtketerangan = ""
    txtsisa = 0
    hapusgrid
    hapusgrid1
    txtbukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If txtbukti = "" Then Exit Sub
    
    carisql1 = "select namasupp, AlamatSupp1,kodesupp from am_supplier"
    namatabel = "Supplier"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtsup = hasil2
    lblsup = hasil
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
    caripiutang
End Sub

Private Sub date1_Change()
    caripiutang
End Sub

Private Sub date2_CloseUp()
    Grid1.TextMatrix(posrow, 3) = Format(date2, "dd/MM/yyyy")

    Grid1.SetFocus
    Grid1.Row = posrow
    date2.Visible = False
End Sub

Private Sub date2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date2.Visible = False
    
    If KeyCode = 13 Then
        Grid1.TextMatrix(posrow, 3) = Format(date2, "dd/MM/yyyy")
        
        Grid1.SetFocus
        Grid1.Row = posrow
        date2.Visible = False
    End If
End Sub

Private Sub date2_LostFocus()
    date2.Visible = False
End Sub

Private Sub Form_Activate()
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "The period is empty !!" & vbCrLf & _
        "Please define Period on proces, Starting period date and Ending period date.", vbCritical, "Critical"
        
        OBJ.Close
        Unload Me
        Exit Sub
    End If
    OBJ.Close
    
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='171' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    date1 = Date
    
    cmbtype.AddItem "Tunai"
    cmbtype.AddItem "Cek"
    cmbtype.AddItem "Giro"
    cmbtype.AddItem "Transfer"
    
    grid2.TextMatrix(0, 0) = "No Apply"
    grid2.TextMatrix(0, 1) = "Hutang"
    grid2.TextMatrix(0, 2) = "Nilai Bayar"
    grid2.TextMatrix(0, 3) = "Disc Bayar"
    grid2.TextMatrix(0, 4) = "Selisih"
    grid2.TextMatrix(0, 5) = "Sisa Hutang"
    
    grid2.ColWidth(0) = 1800
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300
        
    Grid1.TextMatrix(0, 1) = "Type Bayar"
    Grid1.TextMatrix(0, 2) = "No Cek/Giro"
    Grid1.TextMatrix(0, 3) = "Jatuh Tempo"
    Grid1.TextMatrix(0, 4) = "Bank"
    Grid1.TextMatrix(0, 5) = "A/c Bank"
    Grid1.TextMatrix(0, 6) = "Nilai"
    Grid1.TextMatrix(0, 7) = "Administrasi"
    
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1500
    
    Grid1.RowHeightMin = 300
    grid2.RowHeightMin = 300
    
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1.MinDate = RST!tanggal1
        date1.MaxDate = RST!tanggal2
    End If
    OBJ.Close
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 0
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.CellPicture = uncheck Then
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    hapusrow
                    Exit Sub
                End If
            End If
        Case 1
            If Grid1.TextMatrix(Grid1.Row, 1) <> "" Then Exit Sub
            
            If Frame1.Visible = True Then Exit Sub
            
            Frame1.Width = Grid1.ColWidth(Grid1.Col) - 20
            cmbtype.Width = Grid1.ColWidth(Grid1.Col) - 20
            cmbtype = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            Frame1.Left = Grid1.Left + Grid1.CellLeft - 10
            Frame1.Top = Grid1.Top + Grid1.CellTop - 20
            Frame1.Visible = True
            cmbtype.SetFocus
        Case 3
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
    
            If date2.Visible = True Then Exit Sub
            
            date2.Width = Grid1.ColWidth(Grid1.Col) - 20
            date2.Height = 290
            If Grid1.TextMatrix(Grid1.Row, Grid1.Col) <> "" Then date2 = Grid1.TextMatrix(Grid1.Row, 3)
            date2.Left = Grid1.Left + Grid1.CellLeft - 10
            date2.Top = Grid1.Top + Grid1.CellTop - 20
            date2.Visible = True
            date2 = Date
            date2.SetFocus
        Case 2, 4
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtket = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtket.Left = Grid1.Left + Grid1.CellLeft
            txtket.Top = Grid1.Top + Grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6, 7
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" And Grid1.Col = 7 Then Exit Sub
            
            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_EnterCell()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 3
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            If date2.Visible = True Then Exit Sub
            
            date2.Width = Grid1.ColWidth(Grid1.Col) - 20
            date2.Height = 290
            If Grid1.TextMatrix(Grid1.Row, 3) <> "" Then date2 = Grid1.TextMatrix(Grid1.Row, 3)
            date2.Left = Grid1.Left + Grid1.CellLeft - 10
            date2.Top = Grid1.Top + Grid1.CellTop - 20
            date2.Visible = True
            date2 = Date
            date2.SetFocus
        Case 2, 4
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtket = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtket.Left = Grid1.Left + Grid1.CellLeft
            txtket.Top = Grid1.Top + Grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6, 7
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" And Grid1.Col = 7 Then Exit Sub

            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If Grid1.Col = 4 Then
        If hasil = "" Then Exit Sub
    
        Grid1.Row = posrow
        Grid1.Col = 4
        Grid1.CellAlignment = 1
        Grid1.TextMatrix(Grid1.Row, 4) = hasil
        hasil = ""
        hasil1 = ""
        hasil2 = ""
        
        OBJ.Open dsn
        SQL = "select * from am_bank where kode = '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Bank Not Found", vbExclamation, "Warning"
            Grid1.TextMatrix(Grid1.Row, 4) = ""
            Grid1.TextMatrix(Grid1.Row, 5) = ""
        Else
            Grid1.TextMatrix(Grid1.Row, 5) = RST!acc
        End If
        OBJ.Close
    End If
End Sub

Private Sub Grid1_Scroll()
    Frame1.Visible = False
    txtket.Visible = False
    txtnilai1.Visible = False
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = grid2.Row

    Select Case grid2.Col
    Case 2, 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub

        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        If grid2.Col <> 4 Then
            txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
            If txtnilai < 0 Then txtnilai = txtnilai * -1
        End If
    End Select
End Sub

Private Sub grid2_EnterCell()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub

    Select Case grid2.Col
    Case 2, 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub

        posrow = grid2.Row

        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        If grid2.Col <> 4 Then
            txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
            If txtnilai < 0 Then txtnilai = txtnilai * -1
        End If
    End Select
End Sub

Private Sub grid2_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    If txtbukti = "" Then Exit Sub
    
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
    txtnilaikurs = 0
    lblbase = ""
    txtketerangan = ""
    txtsisa = 0
    hapusgrid
    hapusgrid1

    OBJ.Open dsn
    SQL = "Select * From AM_apCashHdr Where Nobkt = '" & txtbukti & "' And kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Payment already exist.", vbInformation, "Information"
        
        txtbukti = ""
        txtbukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtket.Visible = False
    ElseIf KeyAscii = 13 Then
        Select Case Grid1.Col
            Case 2
                For i = 1 To Grid1.Rows - 2
                    If Grid1.TextMatrix(i, 1) = "" Then Exit For
                    If Grid1.TextMatrix(i, 2) = Trim(txtket) Then
                        txtket = ""
                        txtket.Visible = False
                        MsgBox "No Cek/Giro Already exist.", vbExclamation, "Information"
                        
                        Exit Sub
                    End If
                Next i
                
                Grid1.Row = posrow
                Grid1.Col = 2
                
                OBJ2.Open dsn
                SQL2 = "select * from am_apcashsub where nogiro = '" & Trim(txtket) & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    OBJ2.Close
                    
                    txtket = ""
                    txtket.Visible = False
                    
                    MsgBox "No Cek/Giro Already exist.", vbExclamation, "Exclamation"
                        
                    Exit Sub
                Else
                    OBJ2.Close
                End If
                
                Grid1.SetFocus
                Grid1.TextMatrix(Grid1.Row, 2) = Trim(txtket)
                txtket = ""
                txtket.Visible = False
            Case 4
                Grid1.Row = posrow
                Grid1.Col = 4
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "select * from am_bank where kode = '" & txtket & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    OBJ.Close
                    txtket = ""
                    
                    carisql1 = "select kode, description from am_bank"
                    namatabel = "Bank"
   
                    frmsearch.Show vbModal
                Else
                    Grid1.SetFocus
                    Grid1.CellAlignment = 1
                    Grid1.TextMatrix(Grid1.Row, 4) = txtket
                    txtket = ""
                    Grid1.TextMatrix(Grid1.Row, 5) = RST!acc
                    
                    OBJ.Close
                End If
        End Select
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Grid1.SetFocus
End Sub

Private Sub txtkurs_Change()
    txtsup = ""
    lblsup = ""
    txtketerangan = ""
    hapusgrid
    hapusgrid1
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
End Sub

Private Sub txtkurs_LostFocus()
    carikurs
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, grid2.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0

        If grid2.Col = 3 Then
            If Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) < 0 Then
                grid2.SetFocus
                grid2.TextMatrix(grid2.Row, 3) = "0.00"
                txtnilai = 0
                Exit Sub
            End If
        End If

        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")

        txtsisa = hitbayar1 - hitbayar

        grid2.TextMatrix(posrow, 5) = Format((Format(grid2.TextMatrix(posrow, 1), "general number") - Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") + Format(grid2.TextMatrix(posrow, 4), "general number")), "###,###,###,##0.00")

        txtnilai.Visible = False
        grid2.SetFocus
        grid2.Row = posrow
    End If
    If KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
    txtnilai = 0
End Sub

Private Sub txtnilai1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Grid1.Col = 6 Then
        Grid1.TextMatrix(Grid1.Row, 6) = Format(txtnilai1, "###,###,##0.00")

        lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")

        txtsisa = hitbayar1 - hitbayar

        txtnilai1.Visible = False
        Grid1.SetFocus
        Grid1.Row = posrow
        
        txtnilai1 = 0
    ElseIf KeyAscii = 13 And Grid1.Col = 7 Then
        Grid1.TextMatrix(Grid1.Row, 7) = Format(txtnilai1, "###,###,##0.00")
        txtnilai1 = 0

        txtnilai1.Visible = False
        Grid1.SetFocus
        Grid1.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai1 = 0
        txtnilai1.Visible = False
    End If
End Sub

Private Sub txtnilai1_LostFocus()
    txtnilai1.Visible = False
    txtnilai1 = 0
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtsisa_Change()
    lblsisa = " Sisa Bayar : " & Format(txtsisa, "###,###,##0.00")
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalgrid()
    tanggalgrid = Month(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 3))
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

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

        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.ColWidth(0) = 1800
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300

    lblapply = "Total Apply : 0.00"
    lblbayar = "Bayar Apply : 0.00"
End Sub

Private Sub hapusgrid1()
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        Grid1.TextMatrix(Grid1.Row, 1) = ""
        Grid1.TextMatrix(Grid1.Row, 2) = ""
        Grid1.TextMatrix(Grid1.Row, 3) = ""
        Grid1.TextMatrix(Grid1.Row, 4) = ""
        Grid1.TextMatrix(Grid1.Row, 5) = ""
        Grid1.TextMatrix(Grid1.Row, 6) = ""
        Grid1.TextMatrix(Grid1.Row, 7) = ""
        
        Grid1.Col = 0
        Set Grid1.CellPicture = blank
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = 2
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1500

    lbltotal = "Total Bayar : 0.00"
End Sub

Private Sub hapusrow()
    Grid1.TextMatrix(Grid1.Row, 1) = ""
    Grid1.TextMatrix(Grid1.Row, 2) = ""
    Grid1.TextMatrix(Grid1.Row, 3) = ""
    Grid1.TextMatrix(Grid1.Row, 4) = ""
    Grid1.TextMatrix(Grid1.Row, 5) = ""
    Grid1.TextMatrix(Grid1.Row, 6) = ""
    Grid1.TextMatrix(Grid1.Row, 7) = ""
    Do While True
        If Grid1.TextMatrix(Grid1.Row + 1, 1) = "" Then
            Grid1.TextMatrix(Grid1.Row, 1) = ""
            Grid1.TextMatrix(Grid1.Row, 2) = ""
            Grid1.TextMatrix(Grid1.Row, 3) = ""
            Grid1.TextMatrix(Grid1.Row, 4) = ""
            Grid1.TextMatrix(Grid1.Row, 5) = ""
            Grid1.TextMatrix(Grid1.Row, 6) = ""
            Grid1.TextMatrix(Grid1.Row, 7) = ""
            Exit Do
        End If
        Grid1.TextMatrix(Grid1.Row, 1) = Grid1.TextMatrix(Grid1.Row + 1, 1)
        Grid1.TextMatrix(Grid1.Row, 2) = Grid1.TextMatrix(Grid1.Row + 1, 2)
        Grid1.TextMatrix(Grid1.Row, 3) = Grid1.TextMatrix(Grid1.Row + 1, 3)
        Grid1.TextMatrix(Grid1.Row, 4) = Grid1.TextMatrix(Grid1.Row + 1, 4)
        Grid1.TextMatrix(Grid1.Row, 5) = Grid1.TextMatrix(Grid1.Row + 1, 5)
        Grid1.TextMatrix(Grid1.Row, 6) = Grid1.TextMatrix(Grid1.Row + 1, 6)
        Grid1.TextMatrix(Grid1.Row, 7) = Grid1.TextMatrix(Grid1.Row + 1, 7)
        
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = Grid1.Rows - 1
    Grid1.Col = 0
    Set Grid1.CellPicture = blank

    lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")

    txtsisa = hitbayar1 - hitbayar
End Sub

Function hitbayar()
    hitbayar = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        hitbayar = Val(hitbayar) + Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number"))
        
        grid2.Row = grid2.Row + 1
    Loop
End Function

Function hitbayar1()
    hitbayar1 = 0
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        hitbayar1 = Val(hitbayar1) + Val(Format(Grid1.TextMatrix(Grid1.Row, 6), "general number"))

        Grid1.Row = Grid1.Row + 1
    Loop
End Function

Function hitbayar2()
    hitbayar2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        hitbayar2 = Val(hitbayar2) + Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number"))

        grid2.Row = grid2.Row + 1
    Loop
End Function

Private Sub caripiutang()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub

    hapusgrid

    grid2.Row = 1
    
    OBJ.Open dsn
    If lblbase = "1" Then
        SQL = "select NoApply, sum(case when kodecur <> 'IDR' and (transtype = 'CI' or transtype = 'I' or transtype = 'DN' or transtype = 'CN') then (ppn*nilaikurs) when kodecur = 'IDR' then ((Amount + potongan + PPN + selisih)*nilaikurs) end) as Total from AM_Apopnfil WHERE kodesupp = '" & txtsup & "' and tglbeli <= '" & tanggal1 & "' group by Noapply"
    Else
        SQL = "select NoApply, sum(Amount + potongan + selisih) as Total from AM_Apopnfil WHERE kodesupp = '" & txtsup & "' and kodecur = '" & txtkurs & "' and tglbeli <= '" & tanggal1 & "' group by Noapply"
    End If
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do Until RST.EOF
            If Round(RST!Total, 0) <= 0 Then
                RST.MoveNext
                GoTo jump2
            End If

            grid2.TextMatrix(grid2.Row, 0) = RST!noapply
            grid2.TextMatrix(grid2.Row, 1) = Format(RST!Total, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 2) = "0.00"
            grid2.TextMatrix(grid2.Row, 3) = "0.00"
            grid2.TextMatrix(grid2.Row, 4) = "0.00"
            grid2.TextMatrix(grid2.Row, 5) = Format(RST!Total, "###,###,###,##0.00")

            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump2:
        Loop
        OBJ.Close
    Else
        MsgBox "No Transaction For Payment.", vbInformation, "Information"
        OBJ.Close
        
        txtsup = ""
        lblsup = ""
        txtketerangan.SetFocus
    End If
    lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
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

Private Function cariselisih()
    If lblbase = "1" Then
        cariselisih = 0
    Else
        If Asc(Right(txtbukti, 1)) >= 65 And Asc(Right(txtbukti, 1)) <= 90 Then
            OBJ2.Open dsn
            SQL2 = "select nilaikurs from am_apopnfil where transtype = 'I' and noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then
                cariselisih = (RST2!nilaikurs * (Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))) - ((Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number"))) * txtnilaikurs)
            Else
                cariselisih = 0
            End If
            OBJ2.Close
        Else
            OBJ2.Open dsn
            SQL2 = "select sum(amount+potongan+selisih)'net',sum((amount+potongan+selisih)*nilaikurs)'total' from am_apopnfil where kodecur = '" & txtkurs & "' and noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then
                If RST2!net - Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")) = 0 Then cariselisih = RST2!Total - ((Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number"))) * txtnilaikurs) Else cariselisih = 0
            Else
                cariselisih = 0
            End If
            OBJ2.Close
        End If
    End If
End Function
