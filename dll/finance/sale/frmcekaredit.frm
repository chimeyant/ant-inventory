VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmcekaredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Ganti Giro Tolak"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmcekaredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   300
      Left            =   7560
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   0
         TabIndex        =   9
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
      Left            =   7800
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   840
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
      Left            =   8520
      Picture         =   "frmcekaredit.frx":2372
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
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
      Left            =   8160
      Picture         =   "frmcekaredit.frx":2728
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   7800
      TabIndex        =   22
      Top             =   1200
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
      Format          =   102170625
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
      Caption         =   "frmcekaredit.frx":2ADE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":2B4A
      Key             =   "frmcekaredit.frx":2B68
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
      MaxLength       =   13
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
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
      Format          =   102170627
      CurrentDate     =   37421
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmcekaredit.frx":2BA4
      Caption         =   "frmcekaredit.frx":2BC4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":2C30
      Keys            =   "frmcekaredit.frx":2C4E
      Spin            =   "frmcekaredit.frx":2C90
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
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   9420
      _ExtentX        =   16616
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
      TabIndex        =   20
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Ganti"
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
      MICON           =   "frmcekaredit.frx":2CB8
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
      MICON           =   "frmcekaredit.frx":2FD2
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
      MICON           =   "frmcekaredit.frx":32EC
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
      Left            =   6360
      TabIndex        =   14
      Top             =   6120
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
      MICON           =   "frmcekaredit.frx":3606
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
      Left            =   5400
      TabIndex        =   13
      Top             =   6120
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
      MICON           =   "frmcekaredit.frx":3920
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
      Left            =   7560
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "frmcekaredit.frx":3C3A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":3CA6
      Key             =   "frmcekaredit.frx":3CC4
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
      Left            =   8880
      TabIndex        =   7
      Top             =   840
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
      Format          =   102170627
      CurrentDate     =   38767
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmcekaredit.frx":3D00
      Caption         =   "frmcekaredit.frx":3D20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":3D8C
      Keys            =   "frmcekaredit.frx":3DAA
      Spin            =   "frmcekaredit.frx":3DEC
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   10
      Top             =   1560
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
      Left            =   7560
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "frmcekaredit.frx":3E14
      Caption         =   "frmcekaredit.frx":3E34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":3EA0
      Keys            =   "frmcekaredit.frx":3EBE
      Spin            =   "frmcekaredit.frx":3F00
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
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcekaredit.frx":3F28
      Caption         =   "frmcekaredit.frx":3F48
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekaredit.frx":3FB4
      Keys            =   "frmcekaredit.frx":3FD2
      Spin            =   "frmcekaredit.frx":4014
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
   Begin VB.Label Label2 
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
      Width           =   1215
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
      TabIndex        =   34
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblbayar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Ganti Giro : 0.00"
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
      Left            =   2520
      TabIndex        =   31
      Top             =   6120
      Width           =   2655
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
      TabIndex        =   29
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Ganti : 0.00"
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
      Width           =   2655
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
      Left            =   2520
      TabIndex        =   27
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label lblapply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Giro : 0.00"
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
      TabIndex        =   26
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
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
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal"
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
      Left            =   3345
      TabIndex        =   18
      Top             =   150
      Width           =   975
   End
   Begin VB.Label lblsup 
      BackColor       =   &H80000004&
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
      TabIndex        =   17
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   -120
      TabIndex        =   30
      Top             =   6030
      Width           =   12135
   End
End
Attribute VB_Name = "frmcekaredit"
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

Dim str1, str2 As String
Dim posrow As String
Dim i As Integer

Private Sub cmbtype_Click()
    If cmbtype = "" Then Exit Sub
    
    grid1.Row = 1
    Do While cmbtype = "Tunai"
        If grid1.Row = grid1.Rows - 1 Then Exit Do
        
        If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then
            cmbtype = ""
            Frame2.Visible = False
            Exit Sub
        End If
        
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Row = posrow
    
    grid1.SetFocus
    grid1.TextMatrix(grid1.Row, 1) = cmbtype
    grid1.TextMatrix(grid1.Row, 6) = "0.00"
    cmbtype = ""
    Frame2.Visible = False
    
    grid1.Col = 0
    Set grid1.CellPicture = uncheck.Picture
                        
    If grid1.Row = (grid1.Rows - 1) Then grid1.Rows = grid1.Rows + 1
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then cmbtype_LostFocus
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    Frame2.Visible = False
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
    
    If txtsisa <> 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid2.Rows = 2 Or grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar='GT' and idupdate='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, data already process.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ1.Open dsn
    SQL1 = "select * from gl_transaksi where notrx = '" & txtbukti & "' and identry='auto'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close
        
        MsgBox "Data tidak bisa diupdate, transaksi ganti tolak ada di transaksi GL.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ1.Close
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            
        If grid1.TextMatrix(grid1.Row, 1) <> "Tunai" And grid1.TextMatrix(grid1.Row, 3) = "" Then
            MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If grid1.TextMatrix(grid1.Row, 1) <> "Tunai" Then
            OBJ.Open dsn
            SQL = "select * from gl_transaksi where cekbg = '" & grid1.TextMatrix(grid1.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                OBJ.Close
                MsgBox "Can not change, data already process.", vbExclamation, "Warning"
                Exit Sub
            End If
            OBJ.Close
        End If
        
        grid1.Row = grid1.Row + 1
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
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    str1 = ""
    OBJ.Open dsn
    SQL = "select * from am_cashsub Where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!Typebayar = "G" Then
            OBJ1.Open dsn
            SQL1 = "select * from am_cashlin Where kodebayar = 'GT' and noapply = '" & RST!nogiro & "' and nobkt <> '" & txtbukti & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                MsgBox "User can not update, there is already apply.", vbExclamation, "Warning"
                OBJ.Close
                OBJ1.Close
                Exit Sub
            End If
            OBJ1.Close
        End If
        
        RST.MoveNext
    Loop
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
            
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            OBJ1.Open dsn
            SQL1 = "select * from am_cashsub where nogiro = '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                MsgBox "Data Entry Not Complete, please refresh customer.", vbExclamation, "Warning"
                OBJ1.Close
                Exit Sub
            End If
            OBJ1.Close
        End If
        
        grid2.Row = grid2.Row + 1
    Loop
    
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!identry
        date2 = RST!dateentry
    End If
    
    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashsub where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO am_cashhdr"
    SQL = SQL + " (Kodecust"
    SQL = SQL + ", Nobkt"
    SQL = SQL + ", Tglbkt"
    SQL = SQL + ", kodebayar"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", Amount"
    SQL = SQL + ", posted"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", noac"
    SQL = SQL + ", kodecol"
    SQL = SQL + ", IdEntry"
    SQL = SQL + ", DateEntry"
    SQL = SQL + ", IdUpdate"
    SQL = SQL + ", DateUpdate)"
                
    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", 'GT'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ",Convert(Money," & hitbayar & ")"
    SQL = SQL + ", '9'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert(Money," & txtnilaikurs & ")"
    SQL = SQL + ", ''"
    SQL = SQL + ", ''"
    SQL = SQL + ", '" & str2 & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal2 & "')"
    SQL = SQL + ", '0'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid1.Row = 1
    OBJ.Open dsn
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        
        SQL = "INSERT INTO am_cashsub"
        SQL = SQL + " (Nobkt"
        SQL = SQL + ", tglbkt"
        SQL = SQL + ", typebayar"
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
        If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then SQL = SQL + ", 'TN'"
        If grid1.TextMatrix(grid1.Row, 1) = "Giro" Then SQL = SQL + ", 'G'"
        If grid1.TextMatrix(grid1.Row, 1) = "Transfer" Then SQL = SQL + ", 'TF'"
        SQL = SQL + ", '" & txtsup & "'"
        SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 2) & "'"
        If grid1.TextMatrix(grid1.Row, 3) = "" Then SQL = SQL + ",convert(datetime,'01/01/1900')"
        If grid1.TextMatrix(grid1.Row, 3) <> "" Then SQL = SQL + ",convert(datetime,'" & tanggalgrid & "')"
        SQL = SQL + ",convert(datetime,'01/01/1900')"
        SQL = SQL + ",convert(datetime,'01/01/1900')"
        SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 4) & "'"
        SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 5) & "'"
        SQL = SQL + ",convert(money,'" & Format(grid1.TextMatrix(grid1.Row, 6), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid1.Row = grid1.Row + 1
    Loop
    OBJ.Close
    
    grid2.Row = 1
    OBJ.Open dsn
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            
            SQL = "INSERT INTO am_cashlin"
            SQL = SQL + " (Nobkt"
            SQL = SQL + ", tglbkt"
            SQL = SQL + ", kodebayar"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", kodecust"
            SQL = SQL + ", jumlah"
            SQL = SQL + ", selisih"
            SQL = SQL + ", potongan)"
            
            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
            SQL = SQL + ", 'GT'"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") & "')"
            SQL = SQL + ",convert(money,'0')"
            SQL = SQL + ",convert(money,'0'))"
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
    txtnilaikurs = 0
    lblbase = ""
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
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar='GT' and idupdate='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, data already process.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ1.Open dsn
    SQL1 = "select * from gl_transaksi where notrx = '" & txtbukti & "' and identry='auto'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close
        
        MsgBox "Data tidak bisa dihapus, transaksi ganti tolak ada di transaksi GL.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ1.Close
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            
        If grid1.TextMatrix(grid1.Row, 1) <> "Tunai" And grid1.TextMatrix(grid1.Row, 3) = "" Then
            MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If grid1.TextMatrix(grid1.Row, 1) <> "Tunai" Then
            OBJ.Open dsn
            SQL = "select * from gl_transaksi where cekbg = '" & grid1.TextMatrix(grid1.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                OBJ.Close
                MsgBox "Can not change, data already process.", vbExclamation, "Warning"
                Exit Sub
            End If
            OBJ.Close
        End If
        
        grid1.Row = grid1.Row + 1
    Loop

    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If

    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If

    OBJ.Open dsn
    SQL = "select * from am_cashsub Where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!Typebayar = "G" Then
            OBJ1.Open dsn
            SQL1 = "select * from am_cashlin Where kodebayar = 'GT' and noapply = '" & RST!nogiro & "' and nobkt <> '" & txtbukti & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                MsgBox "User can not delete, there is already apply.", vbExclamation, "Warning"
                OBJ.Close
                OBJ1.Close
                Exit Sub
            End If
            OBJ1.Close
        End If
        
        RST.MoveNext
    Loop
    
    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete from am_cashsub where nobkt = '" & txtbukti & "'"
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
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbukti' from am_cashhdr where kodebayar = 'GT' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbukti' from am_cashhdr where kodebayar = 'GT'"
    End If
    namatabel = "Ganti Giro"
    
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

Private Sub date3_CloseUp()
    grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")

    grid1.SetFocus
    grid1.Row = posrow
    date3.Visible = False
End Sub

Private Sub date3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date3.Visible = False
    
    If KeyCode = 13 Then
        grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")
        
        grid1.SetFocus
        grid1.Row = posrow
        date3.Visible = False
    End If
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='132' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
     '   SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='133' and b.kodeuser = '1" & kuser & "'"
     '   Set RST = OBJ.Execute(SQL)
     '   If RST.EOF Then cmdel.Enabled = False
     '   OBJ.Close
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
    
    grid2.TextMatrix(0, 0) = "No Giro"
    grid2.TextMatrix(0, 1) = "Nilai Giro"
    grid2.TextMatrix(0, 2) = "Nilai Ganti"
    grid2.TextMatrix(0, 5) = "Sisa"
    
    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 0
    grid2.ColWidth(4) = 0
    grid2.ColWidth(5) = 1300
    
    grid1.TextMatrix(0, 1) = "Type Ganti"
    grid1.TextMatrix(0, 2) = "No Cek/Giro"
    grid1.TextMatrix(0, 3) = "J/T - Transfer"
    grid1.TextMatrix(0, 4) = "Bank"
    grid1.TextMatrix(0, 5) = "a/c Bank"
    grid1.TextMatrix(0, 6) = "Nilai"
    
    grid1.ColWidth(0) = 300
    grid1.ColWidth(1) = 1500
    grid1.ColWidth(2) = 1500
    grid1.ColWidth(3) = 1500
    grid1.ColWidth(4) = 1000
    grid1.ColWidth(5) = 1500
    grid1.ColWidth(6) = 1500
    
    grid1.RowHeightMin = 300
    grid2.RowHeightMin = 300
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 0
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.CellPicture = uncheck Then
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    hapusrow
                    Exit Sub
                End If
            End If
        Case 1
            If grid1.TextMatrix(grid1.Row, 1) <> "" Then Exit Sub
            If grid1.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If Frame2.Visible = True Then Exit Sub
            
            Frame2.Width = grid1.ColWidth(grid1.Col) - 20
            cmbtype.Width = grid1.ColWidth(grid1.Col) - 20
            cmbtype = grid1.TextMatrix(grid1.Row, grid1.Col)
            Frame2.Left = grid1.Left + grid1.CellLeft - 10
            Frame2.Top = grid1.Top + grid1.CellTop - 20
            Frame2.Visible = True
            cmbtype.SetFocus
        Case 3
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then Exit Sub
            
            If date3.Visible = True Then Exit Sub
            
            date3.Width = grid1.ColWidth(grid1.Col) - 20
            date3.Height = 290
            If grid1.TextMatrix(grid1.Row, grid1.Col) <> "" Then date3 = grid1.TextMatrix(grid1.Row, 3)
            date3.Left = grid1.Left + grid1.CellLeft - 10
            date3.Top = grid1.Top + grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
        Case 2, 4
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Transfer" And grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid1.ColWidth(grid1.Col) - 40
            txtket = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtket.Left = grid1.Left + grid1.CellLeft
            txtket.Top = grid1.Top + grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 5
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            carisql1 = "select acc,description from am_bank"
            namatabel = "Acc Sparta"

            frmsearch.Show vbModal
        Case 6
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai1.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai1 = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai1.Left = grid1.Left + grid1.CellLeft
            txtnilai1.Top = grid1.Top + grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_EnterCell()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 3
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            If date3.Visible = True Then Exit Sub
            
            date3.Width = grid1.ColWidth(grid1.Col) - 20
            date3.Height = 290
            If grid1.TextMatrix(grid1.Row, 3) <> "" Then date3 = grid1.TextMatrix(grid1.Row, 3)
            date3.Left = grid1.Left + grid1.CellLeft - 10
            date3.Top = grid1.Top + grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
        Case 2, 4
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Tunai" Then Exit Sub
            If grid1.TextMatrix(grid1.Row, 1) = "Transfer" And grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid1.ColWidth(grid1.Col) - 40
            txtket = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtket.Left = grid1.Left + grid1.CellLeft
            txtket.Top = grid1.Top + grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub

            txtnilai1.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai1 = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai1.Left = grid1.Left + grid1.CellLeft
            txtnilai1.Top = grid1.Top + grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid1.Col
        Case 5
            grid1.Row = posrow
            grid1.Col = 5
            grid1.CellAlignment = 1
            grid1.TextMatrix(grid1.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
    End Select
End Sub

Private Sub Grid1_Scroll()
    Frame2.Visible = False
    txtket.Visible = False
    txtnilai1.Visible = False
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Then Exit Sub
    posrow = grid2.Row

    Select Case grid2.Col
    Case 2
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
    If txtbukti = "" Or txtsup = "" Then Exit Sub

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
    Cariar
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then
        KeyAscii = 0
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    ElseIf KeyAscii = 13 Then
        Select Case grid1.Col
            Case 2
                For i = 1 To grid1.Rows - 2
                    If grid1.TextMatrix(i, 1) = "" Then Exit For
                    If grid1.TextMatrix(i, 2) = Trim(txtket) Then
                        txtket = ""
                        txtket.Visible = False
                        MsgBox "No Cek/Giro already exist.", vbExclamation, "Information"

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
                    MsgBox "No Cek/Giro already exist.", vbExclamation, "Information"

                    Exit Sub
                End If
                OBJ2.Close

                grid1.SetFocus
                grid1.TextMatrix(grid1.Row, 2) = Trim(txtket)
                txtket = ""
                txtket.Visible = False
            Case 4
                grid1.Row = posrow
                
                grid1.SetFocus
                grid1.Col = 4
                grid1.CellAlignment = 1
                grid1.TextMatrix(grid1.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False
            Case 5
                grid1.SetFocus
                grid1.TextMatrix(grid1.Row, grid1.Col) = txtket
                txtket = ""
                txtket.Visible = False
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
    If KeyAscii = 13 Then grid1.SetFocus
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

        lblbayar = "Ganti Giro : " & Format(hitbayar, "###,###,##0.00")

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
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, grid1.Col) = Format(txtnilai1, "###,###,##0.00")
        txtnilai1 = 0

        lbltotal = "Total Ganti : " & Format(hitbayar1, "###,###,##0.00")

        txtsisa = hitbayar1 - hitbayar

        txtnilai1.Visible = False
        grid1.SetFocus
        grid1.Row = posrow
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

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtsisa_Change()
    lblsisa = " Sisa : " & Format(txtsisa, "###,###,##0.00")
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalgrid()
    tanggalgrid = Month(grid1.TextMatrix(grid1.Row, 3)) & "/" & Day(grid1.TextMatrix(grid1.Row, 3)) & "/" & Year(grid1.TextMatrix(grid1.Row, 3))
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
    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 0
    grid2.ColWidth(4) = 0
    grid2.ColWidth(5) = 1300

    lblapply = "Total Giro : 0.00"
    lblbayar = "Ganti Giro : 0.00"
End Sub

Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.Col = 0
        Set grid1.CellPicture = blank
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    grid1.ColWidth(0) = 300
    grid1.ColWidth(1) = 1500
    grid1.ColWidth(2) = 1500
    grid1.ColWidth(3) = 1500
    grid1.ColWidth(4) = 1000
    grid1.ColWidth(5) = 1500
    grid1.ColWidth(6) = 1500

    lbltotal = "Total Ganti : 0.00"
End Sub

Private Sub hapusrow()
    grid1.TextMatrix(grid1.Row, 1) = ""
    grid1.TextMatrix(grid1.Row, 2) = ""
    grid1.TextMatrix(grid1.Row, 3) = ""
    grid1.TextMatrix(grid1.Row, 4) = ""
    grid1.TextMatrix(grid1.Row, 5) = ""
    grid1.TextMatrix(grid1.Row, 6) = ""
    Do While True
        If grid1.TextMatrix(grid1.Row + 1, 1) = "" Then
            grid1.TextMatrix(grid1.Row, 1) = ""
            grid1.TextMatrix(grid1.Row, 2) = ""
            grid1.TextMatrix(grid1.Row, 3) = ""
            grid1.TextMatrix(grid1.Row, 4) = ""
            grid1.TextMatrix(grid1.Row, 5) = ""
            grid1.TextMatrix(grid1.Row, 6) = ""
            Exit Do
        End If
        grid1.TextMatrix(grid1.Row, 1) = grid1.TextMatrix(grid1.Row + 1, 1)
        grid1.TextMatrix(grid1.Row, 2) = grid1.TextMatrix(grid1.Row + 1, 2)
        grid1.TextMatrix(grid1.Row, 3) = grid1.TextMatrix(grid1.Row + 1, 3)
        grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row + 1, 4)
        grid1.TextMatrix(grid1.Row, 5) = grid1.TextMatrix(grid1.Row + 1, 5)
        grid1.TextMatrix(grid1.Row, 6) = grid1.TextMatrix(grid1.Row + 1, 6)
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = grid1.Rows - 1
    grid1.Col = 0
    Set grid1.CellPicture = blank

    lbltotal = "Total Ganti : " & Format(hitbayar1, "###,###,##0.00")

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
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        hitbayar1 = Val(hitbayar1) + Val(Format(grid1.TextMatrix(grid1.Row, 6), "general number"))

        grid1.Row = grid1.Row + 1
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
    grid1.Row = idx
    grid1.Col = 0
    If hapus Then Set grid1.CellPicture = uncheck.Picture
    grid1.Col = 1
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
    SQL = "Select * From am_cashhdr Where Nobkt = '" & txtbukti & "' And kodebayar = 'GT'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbkt
        date2 = RST!tglbkt
        txtsup = RST!kodecust
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtketerangan = RST!keterangan

        SQL = "Select * From AM_customer Where kodecust = '" & txtsup & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsup = RST!namacust
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
        SQL1 = "SELECT * from am_cashlin WHERE NoBkt = '" & txtbukti & "' and kodebayar = 'GT'"
        Set RST1 = OBJ.Execute(SQL1)
        Do While Not RST1.EOF
            grid2.TextMatrix(grid2.Row, 0) = RST1!noapply
            grid2.TextMatrix(grid2.Row, 2) = Format(RST1!jumlah, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 3) = "0.00"
            grid2.TextMatrix(grid2.Row, 4) = "0.00"

            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST1.MoveNext
        Loop
        'keluarkan record dari aropnfil
        SQL = "select a.Nogiro, (a.jumlah-(select isnull(sum(b.jumlah),0) from am_cashlin b where b.kodebayar = 'GT' and b.noapply=a.nogiro and b.nobkt<>'" & txtbukti & "'))'sisagiro'"
        SQL = SQL + " from am_cashsub a left join am_cashhdr c on a.nobkt=c.nobkt"
        SQL = SQL + " Where Year(a.tglcair) = 1900 And Year(a.tgltolak) <> 1900"
        SQL = SQL + " and c.kodecur='" & txtkurs & "' and a.kodecust = '" & txtsup & "' and a.tgljt <= '" & tanggal1 & "'"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            If Round(RST!sisagiro, 0) = 0 Then
                RST.MoveNext
                GoTo jump3
            End If
            'cek antara yg di grid ama aropnfil
            For i = 1 To grid2.Rows - 2
                If grid2.TextMatrix(i, 0) = RST!nogiro Then
                    grid2.TextMatrix(i, 1) = Format(RST!sisagiro, "###,###,###,##0.00")
                    grid2.TextMatrix(i, 5) = Format(RST!sisagiro - Val(Format(grid2.TextMatrix(i, 2), "general number")), "###,###,###,##0.00")

                    RST.MoveNext
                    GoTo jump3
                End If
            Next i
            'yg tanggalnya lebih dari tanggal piutang
            SQL1 = "select a.Nogiro, (a.jumlah-(select isnull(sum(b.jumlah),0) from am_cashlin b where b.nobkt <> '" & txtbukti & "' and b.kodebayar = 'GT' and b.noapply=a.nogiro))'sisagiro'"
            SQL1 = SQL1 + " from am_cashsub a left join am_cashhdr c on a.nobkt=c.nobkt"
            SQL1 = SQL1 + " Where Year(a.tglcair) = 1900 And Year(a.tgltolak) <> 1900"
            SQL1 = SQL1 + " and c.kodecur='" & txtkurs & "' and a.kodecust = '" & txtsup & "' and a.nogiro = '" & RST!nogiro & "'"
            Set RST1 = OBJ.Execute(SQL1)
            If Round(RST1!sisagiro, 0) = 0 Then
                RST.MoveNext
                GoTo jump3
            End If

            'kalo nga ada di grid nambah dari aropnfil
            grid2.TextMatrix(grid2.Row, 0) = RST!nogiro
            grid2.TextMatrix(grid2.Row, 1) = Format(RST!sisagiro, "###,###,###,##0.00")
            If grid2.TextMatrix(grid2.Row, 2) = "" Then grid2.TextMatrix(grid2.Row, 2) = "0.00"
            If grid2.TextMatrix(grid2.Row, 3) = "" Then grid2.TextMatrix(grid2.Row, 3) = "0.00"
            If grid2.TextMatrix(grid2.Row, 4) = "" Then grid2.TextMatrix(grid2.Row, 4) = "0.00"
            grid2.TextMatrix(grid2.Row, 5) = Format(RST!sisagiro, "###,###,###,##0.00")

            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump3:
        Loop

        lblapply = "Total Giro : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Ganti Giro : " & Format(hitbayar, "###,###,##0.00")

        grid2.Rows = grid2.Rows - 1
        grid2.Col = 0
        grid2.Sort = flexSortStringAscending
        grid2.Rows = grid2.Rows + 1

        grid1.Row = 1

        OBJ1.Open dsn
        SQL1 = "SELECT * from am_cashsub WHERE NoBkt = '" & txtbukti & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            If RST1!Typebayar = "TN" Then grid1.TextMatrix(grid1.Row, 1) = "Tunai"
            If RST1!Typebayar = "G" Then grid1.TextMatrix(grid1.Row, 1) = "Giro"
            If RST1!Typebayar = "TF" Then grid1.TextMatrix(grid1.Row, 1) = "Transfer"
            grid1.TextMatrix(grid1.Row, 2) = RST1!nogiro

            If RST1!Typebayar <> "TN" Then grid1.TextMatrix(grid1.Row, 3) = Format(RST1!tgljt, "dd/MM/yyyy")

            grid1.TextMatrix(grid1.Row, 4) = RST1!bank
            grid1.TextMatrix(grid1.Row, 5) = RST1!acbank
            grid1.TextMatrix(grid1.Row, 6) = Format(RST1!jumlah, "###,###,###,##0.00")

            SetRow grid1.Row, True

            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
            RST1.MoveNext
        Loop
        OBJ1.Close

        lbltotal = "Total Ganti : " & Format(hitbayar1, "###,###,##0.00")

        txtsisa = hitbayar1 - hitbayar
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtbukti = ""
        txtbukti.SetFocus
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function
