VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmpayar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Payment"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpayar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Pembayaran dengan base currency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   37
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   300
      Left            =   8040
      TabIndex        =   36
      Top             =   2055
      Visible         =   0   'False
      Width           =   975
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   0
         TabIndex        =   10
         Top             =   0
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
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1560
      Width           =   6255
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "frmpayar.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":23DE
      Key             =   "frmpayar.frx":23FC
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
      Format          =   61276163
      CurrentDate     =   38767
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":2438
      Caption         =   "frmpayar.frx":2458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":24C4
      Keys            =   "frmpayar.frx":24E2
      Spin            =   "frmpayar.frx":2524
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
      Left            =   9120
      Picture         =   "frmpayar.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   240
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
      Left            =   9120
      Picture         =   "frmpayar.frx":2902
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
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
      Left            =   9120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtkurs 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtkodecol 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
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
      Caption         =   "frmpayar.frx":2CB8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":2D24
      Key             =   "frmpayar.frx":2D42
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
   Begin VB.TextBox txtsup 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1575
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
      Format          =   61276163
      CurrentDate     =   37421
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":2D7E
      Caption         =   "frmpayar.frx":2D9E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":2E0A
      Keys            =   "frmpayar.frx":2E28
      Spin            =   "frmpayar.frx":2E6A
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   1935
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   20
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
      MICON           =   "frmpayar.frx":2E92
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
      TabIndex        =   16
      Top             =   6120
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
      MICON           =   "frmpayar.frx":31AC
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
      TabIndex        =   15
      Top             =   6120
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
      MICON           =   "frmpayar.frx":34C6
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
      TabIndex        =   14
      Top             =   6120
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
      MICON           =   "frmpayar.frx":37E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Collector"
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
      MICON           =   "frmpayar.frx":3AFA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Calculator      =   "frmpayar.frx":3E14
      Caption         =   "frmpayar.frx":3E34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":3EA0
      Keys            =   "frmpayar.frx":3EBE
      Spin            =   "frmpayar.frx":3F00
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
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   465
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
      MICON           =   "frmpayar.frx":3F28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin TDBNumber6Ctl.TDBNumber txtsisa 
      Height          =   255
      Left            =   7320
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":4242
      Caption         =   "frmpayar.frx":4262
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":42CE
      Keys            =   "frmpayar.frx":42EC
      Spin            =   "frmpayar.frx":432E
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
      ValueVT         =   1991573509
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai2 
      Height          =   255
      Left            =   6360
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":4356
      Caption         =   "frmpayar.frx":4376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":43E2
      Keys            =   "frmpayar.frx":4400
      Spin            =   "frmpayar.frx":4442
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
      ValueVT         =   99811333
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai3 
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":446A
      Caption         =   "frmpayar.frx":448A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":44F6
      Keys            =   "frmpayar.frx":4514
      Spin            =   "frmpayar.frx":4556
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
      ValueVT         =   99811333
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai4 
      Height          =   255
      Left            =   6360
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":457E
      Caption         =   "frmpayar.frx":459E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":460A
      Keys            =   "frmpayar.frx":4628
      Spin            =   "frmpayar.frx":466A
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
      ValueVT         =   99811333
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai5 
      Height          =   255
      Left            =   7320
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmpayar.frx":4692
      Caption         =   "frmpayar.frx":46B2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayar.frx":471E
      Keys            =   "frmpayar.frx":473C
      Spin            =   "frmpayar.frx":477E
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
      ValueVT         =   466878469
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblbayar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Bayar Apply : 0.00"
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
      Left            =   3120
      TabIndex        =   34
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label lblapply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Apply : 0.00"
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
      TabIndex        =   30
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7755
      TabIndex        =   25
      Top             =   825
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblnamacol 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      TabIndex        =   23
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Bayar : 0.00"
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
      TabIndex        =   21
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "No. Bukti"
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
      TabIndex        =   19
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Bukti"
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
      Left            =   3225
      TabIndex        =   18
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblsup 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblsisa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Sisa : 0.00"
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
      Left            =   3120
      TabIndex        =   31
      Top             =   6360
      Width           =   3015
   End
End
Attribute VB_Name = "frmpayar"
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
Dim i As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        If lblbase = "1" Then
            txtketerangan = ""
            Check1.Value = 0
        Else
            txtketerangan = "Pembayaran dengan base currency"
        End If
    Else
        txtketerangan = ""
    End If
End Sub

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
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If

    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    If txtsisa <> 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    If Grid1.Rows = 2 Then
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
    
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
            
        If Grid1.TextMatrix(Grid1.Row, 1) <> "Tunai" And Grid1.TextMatrix(Grid1.Row, 3) = "" Then
            MsgBox "Data Entry Not Complete, acc sparta is empty.", vbExclamation, "Warning"
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
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do

        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            OBJ.Open dsn
            SQL = "select * from am_aropnfil where noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "' and transtype <> 'PM'"
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
    SQL = SQL + ", '" & txtkodecol & "'"
    SQL = SQL + ", '" & Check1.Value & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert(Money," & txtnilaikurs & ")"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", '0'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

    Grid1.Row = 1
    OBJ.Open dsn
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do

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
        SQL = SQL + ",convert(money,'" & Format(Grid1.TextMatrix(Grid1.Row, 6), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)

        Grid1.Row = Grid1.Row + 1
        DoEvents
    Loop
    OBJ.Close

    grid2.Row = 1
    OBJ.Open dsn
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do

        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then

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
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 5), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 4), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 3), "general number") & "'))"
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
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
            SQL = SQL + ", '" & txtketerangan & "'"
            SQL = SQL + ", '" & txtkurs & "'"
            SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 2), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 3), "General number") * -1 & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 4), "General number") & "')"
            SQL = SQL + ",Convert (Money, '0'))"
            Set RST = OBJ.Execute(SQL)
        End If
        grid2.Row = grid2.Row + 1
        DoEvents
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
    txtkodecol = ""
    lblnamacol = ""
    txtkurs = ""
    Check1.Value = 0
    txtketerangan = ""
    txtnilaikurs = 0
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

Private Sub cmdsearch1_Click()
    carisql1 = "select kode, nama, idupdate from AM_collector"
    namatabel = "Collector"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecol = hasil
    lblnamacol = hasil1
    caricollector
    hasil = ""
    hasil1 = ""
    hasil2 = ""
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
    
    'validasi data user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='251' and b.kodeuser = '1" & kuser & "'"
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
    cmbtype.AddItem "Giro"
    cmbtype.AddItem "Transfer"

    grid2.TextMatrix(0, 0) = "No Apply"
    grid2.TextMatrix(0, 1) = "Piutang"
    grid2.TextMatrix(0, 2) = "Nilai Bayar"
    grid2.TextMatrix(0, 3) = "Disc Bayar"
    grid2.TextMatrix(0, 4) = "Selisih"
    grid2.TextMatrix(0, 6) = "Sisa Piutang"
    grid2.TextMatrix(0, 5) = "Selisih Kurs"

    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300
    grid2.ColWidth(6) = 1300

    Grid1.TextMatrix(0, 1) = "Type Bayar"
    Grid1.TextMatrix(0, 2) = "No Cek/Giro"
    Grid1.TextMatrix(0, 3) = "J/T - Trans"
    Grid1.TextMatrix(0, 4) = "Bank"
    Grid1.TextMatrix(0, 5) = "Acc Sparta"
    Grid1.TextMatrix(0, 6) = "Nilai"

    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500

    Grid1.RowHeightMin = 300
    grid2.RowHeightMin = 300
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtkodecol = "" Then Exit Sub
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
            If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" And Grid1.TextMatrix(Grid1.Row, 2) = "" Then
                MsgBox "Fill first No Check/Giro", vbExclamation, AppName
                date2.Visible = False
                Exit Sub
            End If

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
        Case 5
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            carisql1 = "select acc,description from am_bank"
            namatabel = "Acc Sparta"

            frmsearch.Show vbModal
        Case 6
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub

            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_EnterCell()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtkodecol = "" Then Exit Sub
    posrow = Grid1.Row

    Select Case Grid1.Col
        Case 3
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" And Grid1.TextMatrix(Grid1.Row, 2) = "" Then
                MsgBox "Fill first No Check/Giro", vbExclamation, AppName
                date2.Visible = False
                Exit Sub
            End If
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
            If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" And Grid1.TextMatrix(Grid1.Row, 2) = "" Then
                'MsgBox "Fill first No Check/Giro", vbExclamation, AppName
                'date2.Visible = False
                'Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub

            txtket.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtket = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtket.Left = Grid1.Left + Grid1.CellLeft
            txtket.Top = Grid1.Top + Grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub

            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case Grid1.Col
        Case 5
            Grid1.Row = posrow
            Grid1.Col = 5
            Grid1.CellAlignment = 1
            Grid1.TextMatrix(Grid1.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
    End Select
End Sub

Private Sub Grid1_Scroll()
    Frame1.Visible = False
    txtket.Visible = False
    txtnilai1.Visible = False
    date2.Visible = False
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Then Exit Sub
    posrow = grid2.Row
    
    Select Case grid2.Col
    Case 2
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.TextMatrix(Grid1.Row, 6) = "0.00" Then
            MsgBox "Transfer value must not be Null", vbExclamation, AppName
            txtnilai.Value = "0.00"
            Exit Sub
        End If
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 3
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid2_EnterCell()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Then Exit Sub
    
    Select Case grid2.Col
    Case 2
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
            
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 3
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid2_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
    
    If Len(txtbukti) = 0 And Not (KeyAscii = 76 Or KeyAscii = 80) Then
        KeyAscii = 0
    ElseIf Len(txtbukti) > 0 And Len(txtbukti) <= 5 And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then
        KeyAscii = 0
    ElseIf KeyAscii = 76 Then
        OBJ.Open dsn
        SQL = "select max(nobkt)'no' from am_cashhdr where kodebayar='PM' and nobkt like 'L%'"
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
        End If
        OBJ.Close
    ElseIf KeyAscii = 80 Then
        OBJ.Open dsn
        SQL = "select max(nobkt)'no' from am_cashhdr where kodebayar='PM' and nobkt like 'P%'"
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
        End If
        OBJ.Close
    End If
End Sub

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
        txtketerangan = ""
        txtkodecol = ""
        lblnamacol = ""
        txtkurs = ""
        txtnilaikurs = 0
        date1 = Date
        txtbukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 27 Then
        txtket_LostFocus
        Exit Sub
    End If
    If KeyAscii = 13 Then
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
                
                OBJ2.Open dsn
                SQL2 = "select * from am_cashsub where nogiro = '" & Trim(txtket) & "'"
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
                
                Grid1.SetFocus
                Grid1.Col = 4
                Grid1.CellAlignment = 1
                Grid1.TextMatrix(Grid1.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
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

Private Sub txtkodecol_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkodecol_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtkurs_Change()
    txtsup = ""
    lblsup = ""
    txtkodecol = ""
    lblnamacol = ""
    Check1.Value = 0
    txtketerangan = ""
    hapusgrid
    hapusgrid1
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

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, grid2.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        
        If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.TextMatrix(Grid1.Row, 6) = "0.00" Then
            MsgBox "Transfer value must not be Null", vbExclamation, AppName
            Exit Sub
        End If
        
        If grid2.Col = 3 Then
            If Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) < 0 Then
                grid2.SetFocus
                grid2.TextMatrix(grid2.Row, 3) = "0.00"
                txtnilai = 0
                Exit Sub
            End If
        End If
        
        lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
        
        txtsisa = hitbayar1 - hitbayar
        If lblbase = "0" Then hitselisihkurs
                
        grid2.TextMatrix(posrow, 6) = Format((Format(grid2.TextMatrix(posrow, 1), "general number") - Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") + Format(grid2.TextMatrix(posrow, 4), "general number")), "###,###,###,##0.00")

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
    If KeyAscii = 13 Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Format(txtnilai1, "###,###,##0.00")
        txtnilai1 = 0
        
        If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" And Grid1.TextMatrix(Grid1.Row, 2) = "" Then
            MsgBox "Fill first No Check/Giro", vbExclamation, AppName
            txtnilai1.Visible = False
            Exit Sub
        ElseIf Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.TextMatrix(Grid1.Row, 6) = "0.00" Then
            MsgBox "Transfer value must not be Null", vbExclamation, AppName
            Exit Sub
        End If
        
        lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")
        
        txtsisa = hitbayar1 - hitbayar
        
        txtnilai1.Visible = False
        Grid1.SetFocus
        Grid1.Row = posrow
    End If
    If KeyAscii = 27 Then
        txtnilai1 = 0
        txtnilai1.Visible = False
    End If
End Sub

Private Sub txtnilai1_LostFocus()
    txtnilai1.Visible = False
    txtnilai1 = 0
End Sub

Private Sub txtnilaikurs_Change()
    If grid2.Rows > 2 Then
        grid2.Row = 1
        Do While True
            grid2.TextMatrix(grid2.Row, 5) = "0.00"
            If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
                posrow = grid2.Row
                If lblbase = "0" Then hitselisihkurs
            End If
            
            grid2.Row = grid2.Row + 1
            If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        Loop
    End If
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtsisa_Change()
    lblsisa = " Sisa : " & Format(txtsisa, "###,###,##0.00")
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecol.SetFocus
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

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalgrid()
    tanggalgrid = Month(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 3))
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

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

        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300
    grid2.ColWidth(6) = 1300

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

    lbltotal = "Total Bayar : 0.00"
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

            grid2.TextMatrix(grid2.Row, 0) = RST!noapply
            grid2.TextMatrix(grid2.Row, 1) = Format(RST!total, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 2) = "0.00"
            grid2.TextMatrix(grid2.Row, 3) = "0.00"
            grid2.TextMatrix(grid2.Row, 4) = "0.00"
            grid2.TextMatrix(grid2.Row, 6) = Format(RST!total, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 5) = "0.00"
            
            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump2:
        Loop
        OBJ.Close
    Else
        MsgBox "No Transaction For Payment.", vbInformation, "Information"
        OBJ.Close
        cmdclear_Click
    End If
End Sub

Private Sub hapusrow()
    Grid1.TextMatrix(Grid1.Row, 1) = ""
    Grid1.TextMatrix(Grid1.Row, 2) = ""
    Grid1.TextMatrix(Grid1.Row, 3) = ""
    Grid1.TextMatrix(Grid1.Row, 4) = ""
    Grid1.TextMatrix(Grid1.Row, 5) = ""
    Grid1.TextMatrix(Grid1.Row, 6) = ""
    Do While True
        If Grid1.TextMatrix(Grid1.Row + 1, 1) = "" Then
            Grid1.TextMatrix(Grid1.Row, 1) = ""
            Grid1.TextMatrix(Grid1.Row, 2) = ""
            Grid1.TextMatrix(Grid1.Row, 3) = ""
            Grid1.TextMatrix(Grid1.Row, 4) = ""
            Grid1.TextMatrix(Grid1.Row, 5) = ""
            Grid1.TextMatrix(Grid1.Row, 6) = ""
            Exit Do
        End If
        Grid1.TextMatrix(Grid1.Row, 1) = Grid1.TextMatrix(Grid1.Row + 1, 1)
        Grid1.TextMatrix(Grid1.Row, 2) = Grid1.TextMatrix(Grid1.Row + 1, 2)
        Grid1.TextMatrix(Grid1.Row, 3) = Grid1.TextMatrix(Grid1.Row + 1, 3)
        Grid1.TextMatrix(Grid1.Row, 4) = Grid1.TextMatrix(Grid1.Row + 1, 4)
        Grid1.TextMatrix(Grid1.Row, 5) = Grid1.TextMatrix(Grid1.Row + 1, 5)
        Grid1.TextMatrix(Grid1.Row, 6) = Grid1.TextMatrix(Grid1.Row + 1, 6)
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = Grid1.Rows - 1
    Grid1.Col = 0
    Set Grid1.CellPicture = blank

    lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")

    txtsisa = hitbayar1 - hitbayar
End Sub

Private Sub hitselisihkurs()
    OBJ.Open dsn
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype<>'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai2 = RST!total
    Else
        txtnilai2 = 0
    End If
    
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype='PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai3 = RST!total * -1
    Else
        txtnilai3 = 0
    End If
    
    txtnilai4 = Val(Format(grid2.TextMatrix(posrow, 2), "general number")) + Val(Format(grid2.TextMatrix(posrow, 3), "general number")) - Val(Format(grid2.TextMatrix(posrow, 4), "general number"))
    txtnilai5 = txtnilai4 + txtnilai3
    grid2.TextMatrix(posrow, 5) = "0.00"
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
        
        txtnilai4 = Val((Format(grid2.TextMatrix(posrow, 2), "general number")) + Val(Format(grid2.TextMatrix(posrow, 3), "general number")) - Val(Format(grid2.TextMatrix(posrow, 4), "general number"))) * txtnilaikurs
        txtnilai5 = txtnilai4 + txtnilai3
        
        If txtnilai2 <> txtnilai5 Then
            grid2.TextMatrix(posrow, 5) = Format(txtnilai2 - txtnilai3 - txtnilai4, "###,###,##0.00")
        End If
    End If
    OBJ.Close
End Sub
Private Sub caricollector()
    If txtkodecol = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * from AM_collector WHERE Kode = '" & txtkodecol & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Collector " & RST!nama & " is not active !", vbExclamation, "Warning"
        txtkodecol = ""
        lblnamacol = ""
    End If
    OBJ.Close
End Sub
