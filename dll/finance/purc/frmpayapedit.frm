VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmpayapedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Bayar Hutang"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpayapedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   300
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   0
         TabIndex        =   40
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   0
   End
   Begin Chameleon.chameleonButton cmdswitch 
      Height          =   285
      Left            =   3030
      TabIndex        =   29
      ToolTipText     =   "Search by ..."
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   ""
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayapedit.frx":2372
      PICN            =   "frmpayapedit.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtketerangan 
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
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1200
      Width           =   6255
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
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   1320
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
      Left            =   0
      Picture         =   "frmpayapedit.frx":4E3E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   840
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
      Left            =   0
      Picture         =   "frmpayapedit.frx":51F4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   7680
      TabIndex        =   21
      Top             =   1200
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
      Format          =   135004161
      CurrentDate     =   38515
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
      Caption         =   "frmpayapedit.frx":55AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":5616
      Key             =   "frmpayapedit.frx":5634
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
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   5880
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
      Left            =   8280
      TabIndex        =   10
      Top             =   4230
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmpayapedit.frx":5670
      Caption         =   "frmpayapedit.frx":5690
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":56FC
      Keys            =   "frmpayapedit.frx":571A
      Spin            =   "frmpayapedit.frx":575C
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Bukti"
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
      MICON           =   "frmpayapedit.frx":5784
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
      Top             =   6510
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
      MICON           =   "frmpayapedit.frx":5A9E
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
      Top             =   6510
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
      MICON           =   "frmpayapedit.frx":5DB8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel 
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   6510
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
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
      MICON           =   "frmpayapedit.frx":60D2
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
      Left            =   6960
      TabIndex        =   12
      Top             =   6510
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Update"
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
      MICON           =   "frmpayapedit.frx":63EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "frmpayapedit.frx":6706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6772
      Key             =   "frmpayapedit.frx":6790
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
   Begin MSComCtl2.DTPicker date3 
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   1950
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   2190
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmpayapedit.frx":67CC
      Caption         =   "frmpayapedit.frx":67EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6858
      Keys            =   "frmpayapedit.frx":6876
      Spin            =   "frmpayapedit.frx":68B8
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
   Begin TDBNumber6Ctl.TDBNumber txtsisa 
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   450
      Calculator      =   "frmpayapedit.frx":68E0
      Caption         =   "frmpayapedit.frx":6900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":696C
      Keys            =   "frmpayapedit.frx":698A
      Spin            =   "frmpayapedit.frx":69CC
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmpayapedit.frx":69F4
      Caption         =   "frmpayapedit.frx":6A14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6A80
      Keys            =   "frmpayapedit.frx":6A9E
      Spin            =   "frmpayapedit.frx":6AE0
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
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   255
      Left            =   10080
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmpayapedit.frx":6B08
      Caption         =   "frmpayapedit.frx":6B28
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6B94
      Keys            =   "frmpayapedit.frx":6BB2
      Spin            =   "frmpayapedit.frx":6BF4
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
      ValueVT         =   935723013
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   255
      Left            =   9360
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmpayapedit.frx":6C1C
      Caption         =   "frmpayapedit.frx":6C3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6CA8
      Keys            =   "frmpayapedit.frx":6CC6
      Spin            =   "frmpayapedit.frx":6D08
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
      ValueVT         =   935723013
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   4110
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   1965
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
   Begin TDBText6Ctl.TDBText txtnotran 
      Height          =   285
      Left            =   1440
      TabIndex        =   41
      Top             =   1590
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmpayapedit.frx":6D30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayapedit.frx":6D9C
      Key             =   "frmpayapedit.frx":6DBA
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
   Begin VB.Label Label8 
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
      Left            =   60
      TabIndex        =   42
      Top             =   1620
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      TabIndex        =   38
      Top             =   6510
      Width           =   7575
   End
   Begin VB.Label Label6 
      Caption         =   "Currency"
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
      TabIndex        =   35
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   255
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
      Left            =   4680
      TabIndex        =   33
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Search by ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   165
      Width           =   255
   End
   Begin VB.Label lblbayar 
      Caption         =   "Bayar Apply : 0.00"
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
      Left            =   7920
      TabIndex        =   30
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
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
      TabIndex        =   28
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lbltotal 
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
      Left            =   7920
      TabIndex        =   20
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblsisa 
      Caption         =   "Sisa Bayar : 0.00"
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
      Left            =   7920
      TabIndex        =   26
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblapply 
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
      Left            =   7920
      TabIndex        =   25
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier"
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
      TabIndex        =   18
      Top             =   870
      Width           =   1095
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
      Left            =   4665
      TabIndex        =   17
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
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "frmpayapedit"
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
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtsisa <> 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid2.Rows = 2 Or Grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
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
            MsgBox "Data Entry Not Complite. No Giro harus diisi", vbExclamation, "Warning"
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
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
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
                MsgBox "Data Entry Not Complete, please refresh supplier.", vbExclamation, "Warning"
                OBJ.Close
                Exit Sub
            End If
            OBJ.Close
        End If
        
        grid2.Row = grid2.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM' and posted = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Data already posted, Update Aborted.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then date2 = RST!dateentry
    
    SQL = "delete from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashlin where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashsub where nobukti = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apopnfil where nobeli = '" & txtbukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
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
            SQL = SQL + ", kodesupp"
            SQL = SQL + ", kodebayar"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", jumlah"
            SQL = SQL + ", selisih"
            SQL = SQL + ", selisihkurs"
            SQL = SQL + ", potongan)"
            
            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
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
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtbukti = ""
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtsisa = 0
    txtketerangan = ""
    txtkurs = ""
    lblbase = ""
    txtnilaikurs = 0
    hapusgrid
    hapusgrid1
    cmdsearch.Enabled = True
    txtbukti.Enabled = True
    date1.Enabled = True
    txtbukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If

    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If

    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If

    OBJ.Open dsn
    SQL = "select * from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM' and posted = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Data already posted, Delete Aborted.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "delete from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete from am_apcashlin where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete from am_apcashsub where nobukti = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apopnfil where nobeli = '" & txtbukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbukti' from AM_apcashhdr where kodebayar = 'PM' and tglbkt >= '" & v_fstgl1 & "' and tglbkt <= '" & v_fstgl2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbukti' from AM_apcashhdr where kodebayar = 'PM'"
    End If
    namatabel = "Bayar Hutang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    hasil = ""
    hasil1 = ""
    Cariar
    txtketerangan.SetFocus
End Sub

Private Sub cmdswitch_Click()
    frmpayapeditsearch.Show vbModal
End Sub

Private Sub cmdswitch_GotFocus()
    If hasil3 = "" Then Exit Sub
    txtbukti = hasil3
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hasil3 = ""
    Cariar
    txtketerangan.SetFocus
End Sub

Private Sub date3_CloseUp()
    Grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")

    Grid1.SetFocus
    Grid1.Row = posrow
    date3.Visible = False
End Sub

Private Sub date3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date3.Visible = False
    
    If KeyCode = 13 Then
        Grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")
        
        Grid1.SetFocus
        Grid1.Row = posrow
        date3.Visible = False
    End If
End Sub

Private Sub date3_LostFocus()
    date3.Visible = False
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
    Grid1.TextMatrix(0, 3) = "J/T - Trans"
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
            If Grid1.TextMatrix(Grid1.Row, 1) <> "" Then
                If MsgBox("Apakah anda akan merubah type transaksi..?", vbQuestion + vbYesNo, AppName) = vbNo Then Exit Sub
            End If
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
            
            If date3.Visible = True Then Exit Sub
            
            date3.Width = Grid1.ColWidth(Grid1.Col) - 20
            date3.Height = 290
            If Grid1.TextMatrix(Grid1.Row, Grid1.Col) <> "" Then date3 = Grid1.TextMatrix(Grid1.Row, 3)
            date3.Left = Grid1.Left + Grid1.CellLeft - 10
            date3.Top = Grid1.Top + Grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
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
                        
            If date3.Visible = True Then Exit Sub
            
            date3.Width = Grid1.ColWidth(Grid1.Col) - 20
            date3.Height = 290
            If Grid1.TextMatrix(Grid1.Row, 3) <> "" Then date3 = Grid1.TextMatrix(Grid1.Row, 3)
            date3.Left = Grid1.Left + Grid1.CellLeft - 10
            date3.Top = Grid1.Top + Grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
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
            If Grid1.TextMatrix(Grid1.Row, 1) = "" And Grid1.Col = 7 Then Exit Sub

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

Private Sub Timer1_Timer()
    If Label5.Visible = True Then Label5.Visible = False Else Label5.Visible = True
    If Label2.Visible = True Then Label2.Visible = False Else Label2.Visible = True
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtketerangan.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    Cariar
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

Private Sub txtkurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
    KeyAscii = 0
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
    lblsisa = " Sisa bayar : " & Format(txtsisa, "###,###,##0.00")
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

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    Grid1.Row = idx
    Grid1.Col = 0
    If hapus Then Set Grid1.CellPicture = uncheck.Picture
    Grid1.Col = 1
End Sub

Private Sub Cariar()
    If txtbukti = "" Then Exit Sub

    hapusgrid
    hapusgrid1
    txtsup = ""
    lblsup = ""
    txtkurs = ""
    txtnilaikurs = 0
    lblbase = ""
    txtketerangan = ""
    date1 = Date

    OBJ.Open dsn
    SQL = "Select * From AM_apCashHdr Where Nobkt = '" & txtbukti & "' And kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbkt
        date2 = RST!tglbkt
        txtsup = RST!kodesupp
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtketerangan = RST!keterangan

        SQL = "Select * From AM_supplier Where kodesupp = '" & txtsup & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsup = RST!namasupp
        Else
            lblsup = ""
        End If
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!base = 1 Then
                lblbase = "1"
            Else
                lblbase = "0"
            End If
        Else
            lblbase = ""
        End If

        txtbukti.Enabled = False
        cmdsearch.Enabled = False
        date1.Enabled = False
        
        'keluarkan record dari cashlin
        grid2.Row = 1
        SQL1 = "SELECT * from AM_apCashLin WHERE NoBkt = '" & txtbukti & "' and kodebayar = 'PM'"
        Set RST1 = OBJ.Execute(SQL1)
        Do While Not RST1.EOF
            grid2.TextMatrix(grid2.Row, 0) = RST1!noapply
            grid2.TextMatrix(grid2.Row, 2) = Format(RST1!jumlah, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 3) = Format(RST1!potongan, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 4) = Format(RST1!selisih, "###,###,###,##0.00")

            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST1.MoveNext
        Loop
        
        'keluarkan notransaksi bank dari GL
        SQL1 = "Select * From gl_transaksi Where notrx = '" & txtbukti & "' and dbkrtrx='K'"
        Set RST1 = OBJ.Execute(SQL1)
        If RST1.EOF Then
        
        Else
        
        End If
        
        
        'keluarkan record dari aropnfil
        If lblbase = "1" Then
            SQL = "select NoApply, sum(case when kodecur <> '" & txtkurs & "' and (transtype = 'CI' or transtype = 'I' or transtype = 'DN' or transtype = 'CN') then (ppn*nilaikurs) when kodecur = '" & txtkurs & "' then ((Amount + potongan + PPN + selisih)*nilaikurs) end) as Total from AM_Apopnfil WHERE nobeli <> '" & txtbukti & "' and kodesupp = '" & txtsup & "' and tglbeli <= '" & tanggal1 & "' group by Noapply"
        Else
            SQL = "select NoApply, sum(Amount + potongan + selisih) as Total from AM_Apopnfil WHERE nobeli <> '" & txtbukti & "' and kodesupp = '" & txtsup & "' and kodecur = '" & txtkurs & "' and tglbeli <= '" & tanggal1 & "' group by Noapply"
        End If
        Set RST = OBJ.Execute(SQL)
        
        Do Until RST.EOF
            If Round(RST!Total, 0) <= 0 Then
                RST.MoveNext
                GoTo jump3
            End If
            'cek antara yg di grid ama aropnfil
            For i = 1 To grid2.Rows - 2
                If grid2.TextMatrix(i, 0) = RST!noapply Then
                    grid2.TextMatrix(i, 1) = Format(SpyRound(Format(RST!Total, "###,###,###,##0.00")), "###,###,###,##0.00")
                    grid2.TextMatrix(i, 5) = Format(SpyRound(Format(RST!Total - Val(Format(grid2.TextMatrix(i, 2), "general number")) - Val(Format(grid2.TextMatrix(i, 3), "general number")) + Val(Format(grid2.TextMatrix(i, 4), "general number")), "###,###,###,##0.00")), "###,###,###,##0.00")
                    RST.MoveNext
                    GoTo jump3
                End If
            Next i
            
            'cek yg tanggalnya lebih dari tanggal piutang
            If lblbase = "1" Then
                SQL1 = "select NoApply, sum(case when kodecur <> '" & txtkurs & "' and (transtype = 'CI' or transtype = 'I' or transtype = 'DN' or transtype = 'CN') then (ppn*nilaikurs) when kodecur = '" & txtkurs & "' then ((Amount + potongan + PPN + selisih)*nilaikurs) end) as Total from AM_Apopnfil WHERE noapply = '" & RST!noapply & "' and kodesupp = '" & txtsup & "' group by Noapply"
            Else
                SQL1 = "select NoApply, sum(Amount + potongan + selisih) as Total from AM_Apopnfil WHERE noapply = '" & RST!noapply & "' and kodesupp = '" & txtsup & "' and kodecur = '" & txtkurs & "' group by Noapply"
            End If
            Set RST1 = OBJ.Execute(SQL1)
            If Round(RST1!Total, 0) <= 0 Then
                RST.MoveNext
                GoTo jump3
            End If

            'kalo nga ada di grid nambah dari aropnfil
            grid2.TextMatrix(grid2.Row, 0) = RST!noapply
            grid2.TextMatrix(grid2.Row, 1) = Format(RST!Total, "###,###,###,##0.00")
            If grid2.TextMatrix(grid2.Row, 2) = "" Then grid2.TextMatrix(grid2.Row, 2) = "0.00"
            If grid2.TextMatrix(grid2.Row, 3) = "" Then grid2.TextMatrix(grid2.Row, 3) = "0.00"
            If grid2.TextMatrix(grid2.Row, 4) = "" Then grid2.TextMatrix(grid2.Row, 4) = "0.00"
            grid2.TextMatrix(grid2.Row, 5) = Format(RST!Total, "###,###,###,##0.00")

            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump3:
        Loop

        lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")

        grid2.Rows = grid2.Rows - 1
        grid2.Col = 0
        grid2.Sort = flexSortStringAscending
        grid2.Rows = grid2.Rows + 1

        Grid1.Row = 1

        OBJ1.Open dsn
        SQL1 = "SELECT * from AM_apCashsub WHERE NoBukti = '" & txtbukti & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            If RST1!Type = "TN" Then Grid1.TextMatrix(Grid1.Row, 1) = "Tunai"
            If RST1!Type = "C" Then Grid1.TextMatrix(Grid1.Row, 1) = "Cek"
            If RST1!Type = "G" Then Grid1.TextMatrix(Grid1.Row, 1) = "Giro"
            If RST1!Type = "TF" Then Grid1.TextMatrix(Grid1.Row, 1) = "Transfer"
            Grid1.TextMatrix(Grid1.Row, 2) = RST1!nogiro

            If RST1!Type <> "TN" Then Grid1.TextMatrix(Grid1.Row, 3) = Format(RST1!tgljt, "dd/MM/yyyy")

            Grid1.TextMatrix(Grid1.Row, 4) = RST1!bank
            Grid1.TextMatrix(Grid1.Row, 5) = RST1!acbank
            Grid1.TextMatrix(Grid1.Row, 6) = Format(RST1!jumlah, "###,###,###,##0.00")
            Grid1.TextMatrix(Grid1.Row, 7) = Format(RST1!byadmin, "###,###,###,##0.00")


            SetRow Grid1.Row, True

            Grid1.Rows = Grid1.Rows + 1
            Grid1.Row = Grid1.Row + 1
            RST1.MoveNext
        Loop
        OBJ1.Close

        lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")

        txtsisa = hitbayar1 - hitbayar
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtbukti = ""
        txtbukti.SetFocus
    End If
    OBJ.Close
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
            SQL2 = "select sum(amount+potongan+selisih)'net',sum((amount+potongan+selisih)*nilaikurs)'total' from am_apopnfil where kodecur = '" & txtkurs & "' and nobeli <> '" & txtbukti & "' and noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "'"
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


