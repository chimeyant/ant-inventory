VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmterimaedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Mutasi Barang"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   Icon            =   "frmterimaedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcust 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1920
      Width           =   5775
   End
   Begin VB.TextBox txtgudang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "frmterimaedit.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":23DE
      Key             =   "frmterimaedit.frx":23FC
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
      MaxLength       =   50
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
      Left            =   7440
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmterimaedit.frx":2438
      Caption         =   "frmterimaedit.frx":2458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":24C4
      Keys            =   "frmterimaedit.frx":24E2
      Spin            =   "frmterimaedit.frx":2524
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   1575
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
      Left            =   7560
      Picture         =   "frmterimaedit.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   1560
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
      Left            =   7320
      Picture         =   "frmterimaedit.frx":289A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   1560
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
      Left            =   7800
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   840
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
      Format          =   90439683
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   4800
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmterimaedit.frx":2B7C
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
      TabIndex        =   11
      Top             =   4800
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmterimaedit.frx":2E96
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
      TabIndex        =   9
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmterimaedit.frx":31B0
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
      TabIndex        =   10
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmterimaedit.frx":34CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   480
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
      MICON           =   "frmterimaedit.frx":37E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Customer"
      ENAB            =   0   'False
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
      MICON           =   "frmterimaedit.frx":3AFE
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
      Height          =   225
      Left            =   8400
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterimaedit.frx":3E18
      Caption         =   "frmterimaedit.frx":3E38
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":3EA4
      Keys            =   "frmterimaedit.frx":3EC2
      Spin            =   "frmterimaedit.frx":3F04
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   225
      Left            =   8400
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterimaedit.frx":3F2C
      Caption         =   "frmterimaedit.frx":3F4C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":3FB8
      Keys            =   "frmterimaedit.frx":3FD6
      Spin            =   "frmterimaedit.frx":4018
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
   Begin TDBNumber6Ctl.TDBNumber txtnil3 
      Height          =   225
      Left            =   8400
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterimaedit.frx":4040
      Caption         =   "frmterimaedit.frx":4060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":40CC
      Keys            =   "frmterimaedit.frx":40EA
      Spin            =   "frmterimaedit.frx":412C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   6720
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   90439683
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   6720
      TabIndex        =   29
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   90439683
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil4 
      Height          =   225
      Left            =   8400
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterimaedit.frx":4154
      Caption         =   "frmterimaedit.frx":4174
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimaedit.frx":41E0
      Keys            =   "frmterimaedit.frx":41FE
      Spin            =   "frmterimaedit.frx":4240
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
   Begin VB.Label Label8 
      Caption         =   "Desc/Reference"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Label lblcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Gudang"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   1200
      Width           =   4095
   End
   Begin MSForms.ComboBox cmbtype 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   735
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1296;503"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbltype 
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   150
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Type Transaksi"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   1950
      Width           =   1455
   End
End
Attribute VB_Name = "frmterimaedit"
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

Dim posrow, poscol As String
Dim i As Integer
Dim hitunginout As Boolean

Private Sub cmbtype_Change()
    hapusemua
    cmdsearch1.Enabled = False
    txtcust.Enabled = False
    txtnobukti = ""
    date1 = Date
    txtnobukti.SetFocus
    
    If cmbtype = "01" And cmbtype.ListIndex = 0 Then lbltype = "Produksi Harian Lem (In)"
    If cmbtype = "01" And cmbtype.ListIndex = 1 Then lbltype = "Produksi Harian Karet (In)"
    If cmbtype = "02" Then lbltype = "Terima Over Zak"
    If cmbtype = "03" Then lbltype = "Keluar Over Zak"
    If cmbtype = "04" Then
        lbltype = "Terima Retur Customer"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
    End If
    If cmbtype = "05" Then lbltype = "Barang Rusak (Out)"
    If cmbtype = "06" Then lbltype = "Keluar Sampel"
    If cmbtype = "07" Then lbltype = "Barang Bonus (In)"
    If cmbtype = "08" Then lbltype = "Dari Bahan Baku (In)"
    If cmbtype = "09" Then lbltype = "Terima Dari Pabrik (In)"
    If cmbtype = "10" Then lbltype = "Retur Ke Pabrik (Out)"
    
    If cmbtype = "12" Then
        lbltype = "Terima Pinjaman dr Customer (In)"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
    End If
    If cmbtype = "13" Then
        lbltype = "Kembali Pinjaman dr Customer (Out)"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
    End If
    If cmbtype = "14" Then lbltype = "Keluar Barang Bocor (Out)"
    If cmbtype = "15" Then lbltype = "Terima Tuangan Lem (In)"
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtnobukti.SetFocus
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtgudang = "" Or grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If cmbtype = "11" Then
        MsgBox "Invalid Type.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If (cmbtype = "04" Or cmbtype = "12" Or cmbtype = "13") And txtcust = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
       
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Or Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) < Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    If par1 = "1" Then hitunginout = True Else hitunginout = False
    
    If hitunginout Then
        Label2 = "Checking stock on hand ..."
        grid.Row = 1
        Do While grid.TextMatrix(grid.Row, 1) <> ""
            OBJ.Open dsn
            SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
            
            If par5 = "0" Then
                SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Else
                SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            End If
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
            
            SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj < '" & tanggalinv & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
            OBJ.Close
            
            If cmbtype = "01" Or cmbtype = "02" Or cmbtype = "04" Or cmbtype = "07" Or cmbtype = "08" Or cmbtype = "09" Or cmbtype = "12" Or cmbtype = "15" Then
                txtnil3 = txtnil1 - txtnil2 - txtnil4 + Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
            Else
                txtnil3 = txtnil1 - txtnil2 - txtnil4 - Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
            End If
            date2 = date1
            date3 = date1
            
            OBJ.Open dsn
            SQL = "select isnull(max(a.tglbpb),01/01/1900)'tanggal' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
            
            If par5 = "0" Then
                SQL = "select isnull(max(a.tglsj),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Else
                SQL = "select isnull(max(b.tglkirim),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            End If
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
            
            SQL = "select isnull(max(tglsj),01/01/1900)'tanggal' from am_sjsby where kodegudang = '" & txtgudang & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
            OBJ.Close
            
            Do While True
                OBJ.Open dsn
                SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.nobpb <> '" & txtnobukti & "' and a.tglbpb = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
                
                If par5 = "0" Then
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Else
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
                
                SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj = '" & tanggal2 & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
                OBJ.Close
                
                txtnil3 = txtnil3 + txtnil1 - txtnil2 - txtnil4
                
                If txtnil3 < 0 Then
                    MsgBox "Stock Limited on " & grid.TextMatrix(grid.Row, 2), vbOKOnly + vbExclamation, "Warning"
                    Label2 = ""
                    Exit Sub
                End If
                            
                If date2 = date3 Then Exit Do
                
                date2 = date2 + 1
            Loop
            
            grid.Row = grid.Row + 1
        Loop
        Label2 = "Calculating and make sure stock is available ..."
        
        OBJ.Open dsn
        SQL = "select * from am_bpblin where nobpb = '" & txtnobukti & "' order by lineitem"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            For i = 1 To grid.Rows - 2
                If RST!kodebarang = grid.TextMatrix(i, 1) Then GoTo balik3
            Next i
            
            OBJ1.Open dsn
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil1 = RST1!qty Else txtnil1 = 0
            
            If par5 = "0" Then
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Else
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            End If
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil2 = RST1!qty Else txtnil2 = 0
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj < '" & tanggalinv & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil4 = RST1!qty Else txtnil4 = 0
            OBJ1.Close
            
            txtnil3 = txtnil1 - txtnil2 - txtnil4
            date2 = date1
            date3 = date1
            
            OBJ1.Open dsn
            SQL1 = "select isnull(max(a.tglbpb),01/01/1900)'tanggal' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            
            If par5 = "0" Then
                SQL1 = "select isnull(max(a.tglsj),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Else
                SQL1 = "select isnull(max(b.tglkirim),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            End If
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            
            SQL1 = "select isnull(max(tglsj),01/01/1900)'tanggal' from am_sjsby where kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            OBJ1.Close
            
            Do While True
                OBJ1.Open dsn
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.nobpb <> '" & txtnobukti & "' and a.tglbpb = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil1 = RST1!qty Else txtnil1 = 0
                
                If par5 = "0" Then
                    SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                Else
                    SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                End If
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil2 = RST1!qty Else txtnil2 = 0
                
                SQL1 = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj = '" & tanggal2 & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil4 = RST1!qty Else txtnil4 = 0
                OBJ1.Close
                
                txtnil3 = txtnil3 + txtnil1 - txtnil2 - txtnil4
                
                If txtnil3 < 0 Then
                    MsgBox "Stock Limited on " & grid.TextMatrix(grid.Row, 2), vbOKOnly + vbExclamation, "Warning"
                    Label2 = ""
                    Exit Sub
                End If
                            
                If date2 = date3 Then Exit Do
                
                date2 = date2 + 1
            Loop
balik3:
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    
    Label2 = "Updating header and inserting data ..."
    OBJ.Open dsn
    SQL = "update am_bpbhdr set idupdate = '" & kuser & "',dateupdate = convert(datetime,'" & tanggalsekarang & "'),keterangan = '" & txtapply & "' where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_bpblin where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "insert into am_bpblin ("
        SQL = SQL + "type,"
        SQL = SQL + "nobpb,"
        SQL = SQL + "tglbpb,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values("
        SQL = SQL + "'" & cmbtype & "',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        If cmbtype = "03" Or cmbtype = "05" Or cmbtype = "06" Or cmbtype = "10" Or cmbtype = "13" Or cmbtype = "14" Then SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") * -1 & "'),"
        If cmbtype = "01" Or cmbtype = "02" Or cmbtype = "04" Or cmbtype = "07" Or cmbtype = "08" Or cmbtype = "09" Or cmbtype = "12" Or cmbtype = "15" Then SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    Label2 = "Proces complete ..."
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    cmdsearch1.Enabled = False
    txtcust.Enabled = False
    txtnobukti.Enabled = True
    cmdsearch.Enabled = True
    cmbtype.Enabled = True
    date1.Enabled = True
    date1 = Date
    txtnobukti = ""
    cmbtype = ""
    lbltype = ""
    cmbtype.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If txtnobukti = "" Or cmbtype = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complte", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    If par1 = "1" Then hitunginout = True Else hitunginout = False
    
    If hitunginout Then
        Label2 = "Checking stock on hand ..."
        OBJ.Open dsn
        SQL = "select * from am_bpblin where nobpb = '" & txtnobukti & "' order by lineitem"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil1 = RST1!qty Else txtnil1 = 0
            
            If par5 = "0" Then
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Else
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            End If
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil2 = RST1!qty Else txtnil2 = 0
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj < '" & tanggalinv & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnil4 = RST1!qty Else txtnil4 = 0
            OBJ1.Close
        
            txtnil3 = txtnil1 - txtnil2 - txtnil4
            date2 = date1
            date3 = date1
        
            OBJ1.Open dsn
            SQL1 = "select isnull(max(a.tglbpb),01/01/1900)'tanggal' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            
            If par5 = "0" Then
                SQL1 = "select isnull(max(a.tglsj),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            Else
                SQL1 = "select isnull(max(b.tglkirim),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
            End If
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            
            SQL1 = "select isnull(max(tglsj),01/01/1900)'tanggal' from am_sjsby where kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then If date3 < RST1!tanggal Then date3 = RST1!tanggal
            OBJ1.Close
        
            Do While True
                OBJ1.Open dsn
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.nobpb <> '" & txtnobukti & "' and a.tglbpb = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil1 = RST1!qty Else txtnil1 = 0
                
                If par5 = "0" Then
                    SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                Else
                    SQL1 = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST!kodebarang & "' and a.kodesatuan = '" & RST!kodesatuan & "'"
                End If
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil2 = RST1!qty Else txtnil2 = 0
                
                SQL1 = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj = '" & tanggal2 & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnil4 = RST1!qty Else txtnil4 = 0
                OBJ1.Close
        
                txtnil3 = txtnil3 + txtnil1 - txtnil2 - txtnil4
        
                If txtnil3 < 0 Then
                    MsgBox "Can not update data, quantity item Limited.", vbOKOnly + vbExclamation, "Warning"
                    Exit Sub
                End If
                            
                If date2 = date3 Then Exit Do
                
                date2 = date2 + 1
            Loop
            
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    
    Label2 = "Deleting data on database ..."
    OBJ.Open dsn
    SQL = "delete am_bpbhdr where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete am_bpblin where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    Label2 = "Proces complete ..."
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    If cmbtype = "" Then Exit Sub
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        If cmbtype = "01" And cmbtype.ListIndex = 0 Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PHL0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "01" And cmbtype.ListIndex = 1 Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PHK0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "02" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TOZ0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "03" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KOZ0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "04" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TR0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "05" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'BR0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "06" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KS0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "07" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'BB0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "08" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'DBB0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "09" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TDP0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "10" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'RKP0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        
        If cmbtype = "12" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TPC0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "13" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KPC0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "14" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KBB0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
        If cmbtype = "15" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TTL0-%' and type = '" & cmbtype & "' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
    Else
        If cmbtype = "01" And cmbtype.ListIndex = 0 Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PHL0-%' and type = '" & cmbtype & "'"
        If cmbtype = "01" And cmbtype.ListIndex = 1 Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PHK0-%' and type = '" & cmbtype & "'"
        If cmbtype = "02" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TOZ0-%' and type = '" & cmbtype & "'"
        If cmbtype = "03" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KOZ0-%' and type = '" & cmbtype & "'"
        If cmbtype = "04" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TR0-%' and type = '" & cmbtype & "'"
        If cmbtype = "05" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'BR0-%' and type = '" & cmbtype & "'"
        If cmbtype = "06" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KS0-%' and type = '" & cmbtype & "'"
        If cmbtype = "07" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'BB0-%' and type = '" & cmbtype & "'"
        If cmbtype = "08" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'DBB0-%' and type = '" & cmbtype & "'"
        If cmbtype = "09" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TDP0-%' and type = '" & cmbtype & "'"
        If cmbtype = "10" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'RKP0-%' and type = '" & cmbtype & "'"
        
        If cmbtype = "12" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TPC0-%' and type = '" & cmbtype & "'"
        If cmbtype = "13" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KPC0-%' and type = '" & cmbtype & "'"
        If cmbtype = "14" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'KBB0-%' and type = '" & cmbtype & "'"
        If cmbtype = "15" Then carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'TTL0-%' and type = '" & cmbtype & "'"
    End If
    namatabel = "Mutasi Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    carinvoice
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtcust = hasil
    lblcust = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtapply.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='102' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='103' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    cmbtype.Clear
    cmbtype.ColumnCount = 2
    cmbtype.ListWidth = "7 cm"
    cmbtype.ColumnWidths = "2 cm; 5 cm"
    
    cmbtype.AddItem "01"
    cmbtype.AddItem "01"
    cmbtype.AddItem "02"
    cmbtype.AddItem "03"
    cmbtype.AddItem "04"
    cmbtype.AddItem "05"
    cmbtype.AddItem "06"
    cmbtype.AddItem "07"
    cmbtype.AddItem "08"
    cmbtype.AddItem "09"
    cmbtype.AddItem "10"
    
    cmbtype.AddItem "12"
    cmbtype.AddItem "13"
    cmbtype.AddItem "14"
    cmbtype.AddItem "15"
    
    cmbtype.List(0, 1) = "Produksi Harian Lem (In)"
    cmbtype.List(1, 1) = "Produksi Harian Karet (In)"
    cmbtype.List(2, 1) = "Terima Over Zak"
    cmbtype.List(3, 1) = "Keluar Over Zak"
    cmbtype.List(4, 1) = "Terima Retur Customer"
    cmbtype.List(5, 1) = "Barang Rusak (Out)"
    cmbtype.List(6, 1) = "Keluar Sampel"
    cmbtype.List(7, 1) = "Barang Bonus (In)"
    cmbtype.List(8, 1) = "Dari Bahan Baku (In)"
    cmbtype.List(9, 1) = "Terima Dari Pabrik (In)"
    cmbtype.List(10, 1) = "Retur Ke Pabrik (Out)"
    
    cmbtype.List(11, 1) = "Terima Pinjaman dr Customer (In)"
    cmbtype.List(12, 1) = "Kembali Pinjaman dr Customer (Out)"
    cmbtype.List(13, 1) = "Keluar Barang Bocor (Out)"
    cmbtype.List(14, 1) = "Terima Tuangan Lem (In)"
    
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "K/Satuan"
    grid.TextMatrix(0, 5) = "N/Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.RowHeightMin = 300
    
    date1.Value = Date
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    If txtnobukti = "" Or txtgudang = "" Then Exit Sub
    posrow = grid.Row
    poscol = grid.Col
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
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 200 Then
                MsgBox "Maximum line 200", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtnilai.Visible = True Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtgudang = "" Then Exit Sub
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        If txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    If grid.Col = 4 Then
        grid.Row = 1
        Do While True
            If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 4) = hasil And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
            
                MsgBox "Kode Barang already exist.", vbInformation, "Information"
                hasil = ""
                grid.Row = posrow
                grid.Col = 4
                grid.SetFocus
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
    End If

    grid.Row = posrow
    grid.Col = poscol
    grid.CellAlignment = 1
    grid.TextMatrix(grid.Row, grid.Col) = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""

    If grid.Col = 1 Then
        OBJ.Open dsn
        If cmbtype = "07" Then
            SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodeproduk = 'C999'"
        Else
            SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        End If
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = "0.00"
            
            SetRow grid.Row, True
            lbltotal.Caption = "    Total Barang : " & grid.Rows - 1
            grid.SetFocus
            grid.Col = 2
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
        End If
        OBJ.Close
    End If

    If grid.Col = 4 Then
        OBJ.Open dsn
        SQL = "SELECT a.namabarang,b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            
            grid.SetFocus
            grid.Col = 5
        Else
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtcust_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtcust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtgudang_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtgudang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 1
                grid.Row = posrow
                
                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                If cmbtype = "07" Then
                    SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodeproduk = 'C999' and len(kodebarang)=8"
                Else
                    SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and len(kodebarang)=8"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    grid.TextMatrix(grid.Row, 2) = RST!namabarang
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    
                    OBJ.Close
    
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    
                    If cmbtype = "07" Then
                        carisql1 = "select kodebarang, namabarang from am_itemmst where kodeproduk = 'C999'"
                        namatabel = "Item "
                    Else
                        carisql1 = "select kodebarang, namabarang from am_itemmst"
                        namatabel = "Item"
                    End If
   
                    frmsearch.Show vbModal
                End If
                grid.Col = 1
            Case 4
                grid.Row = 1
                Do While True
                    If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                    If grid.TextMatrix(grid.Row, 4) = txtket And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
                    
                        MsgBox "Kode Barang already exist.", vbInformation, "Information"
                        txtket = ""
                        grid.Row = posrow
                        grid.Col = 4
                        grid.SetFocus
                        Exit Sub
                    End If
                    grid.Row = grid.Row + 1
                Loop
                
                grid.Row = posrow
                
                grid.SetFocus
                grid.Col = 4
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "SELECT namabarang,kodesatuan FROM AM_ITEMDTL WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    grid.TextMatrix(posrow, 4) = ""
                    
                    txtket = ""
                    
                    carisql1 = "SELECT b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b on a.kodesatuan = b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                    namatabel = "Satuan "
                        
                    frmsearch.Show vbModal
                Else
                    grid.TextMatrix(grid.Row, 2) = RST!namabarang
                    
                    SQL = "SELECT namasatuan FROM AM_unit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then grid.TextMatrix(grid.Row, 2) = RST!namasatuan
                End If
                OBJ.Close
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

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtnilai_LostFocus
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
bawah:
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
        grid.Col = poscol
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub hapusemua()
    txtcust = ""
    lblcust = ""
    
    txtgudang = ""
    lblgudang = ""
    txtapply = ""
    
    hapusgrid
    
    lbltotal.Caption = "    Total Barang : 0"
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
    
    If grid.Rows = 2 Then
        lbltotal.Caption = "    Total Barang : 0"
    Else
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    End If
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Or cmbtype = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusemua
    
    OBJ.Open dsn
    SQL = "select * from am_bpbhdr where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbpb
        txtgudang = RST!kodegudang
        txtapply = RST!keterangan
        txtcust = RST!noref
                
        SQL = "select * from am_gudang where kodegudang = '" & txtgudang & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblgudang = RST!namagudang
        
        SQL = "select * from am_customer where kodecust = '" & txtcust & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblcust = RST!namacust

        grid.Row = 1
        SQL = "select * from am_bpblin where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            If cmbtype = "03" Or cmbtype = "05" Or cmbtype = "06" Or cmbtype = "10" Or cmbtype = "13" Or cmbtype = "14" Then grid.TextMatrix(grid.Row, 3) = Format(RST!qty * -1, "###,###,##0.00")
            If cmbtype = "01" Or cmbtype = "02" Or cmbtype = "04" Or cmbtype = "07" Or cmbtype = "08" Or cmbtype = "09" Or cmbtype = "12" Or cmbtype = "15" Then grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
            
            OBJ1.Open dsn
            SQL1 = "select a.namabarang,b.namasatuan from am_itemdtl a left join am_unit b on a.kodesatuan = b.kodesatuan where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            grid.TextMatrix(grid.Row, 2) = RST1!namabarang
            grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            OBJ1.Close
                    
            SetRow grid.Row, True
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
        txtnobukti.Enabled = False
        cmdsearch.Enabled = False
        cmbtype.Enabled = False
        date1.Enabled = False
        txtgudang.SetFocus
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub
