VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmterima 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Mutasi Barang"
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
   Icon            =   "frmterima.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nomor Lot :"
      Height          =   750
      Left            =   6240
      TabIndex        =   32
      Top             =   90
      Visible         =   0   'False
      Width           =   3075
      Begin VB.TextBox txtnolot 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   105
         MaxLength       =   15
         TabIndex        =   33
         Top             =   255
         Width           =   2850
      End
   End
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
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmterima.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":23DE
      Key             =   "frmterima.frx":23FC
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
      Calculator      =   "frmterima.frx":2438
      Caption         =   "frmterima.frx":2458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":24C4
      Keys            =   "frmterima.frx":24E2
      Spin            =   "frmterima.frx":2524
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
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1920
      Width           =   5775
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
      Left            =   3360
      Picture         =   "frmterima.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
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
      Left            =   3600
      Picture         =   "frmterima.frx":289A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
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
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   480
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
      Format          =   144113667
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
      TabIndex        =   11
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
      MICON           =   "frmterima.frx":2B7C
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
      MICON           =   "frmterima.frx":2E96
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
      Left            =   6360
      TabIndex        =   9
      Top             =   4800
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmterima.frx":31B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Gudang"
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
      MICON           =   "frmterima.frx":34CA
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
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterima.frx":37E4
      Caption         =   "frmterima.frx":3804
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":3870
      Keys            =   "frmterima.frx":388E
      Spin            =   "frmterima.frx":38D0
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
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterima.frx":38F8
      Caption         =   "frmterima.frx":3918
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":3984
      Keys            =   "frmterima.frx":39A2
      Spin            =   "frmterima.frx":39E4
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   25
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
      MICON           =   "frmterima.frx":3A0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Calculator      =   "frmterima.frx":3D26
      Caption         =   "frmterima.frx":3D46
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":3DB2
      Keys            =   "frmterima.frx":3DD0
      Spin            =   "frmterima.frx":3E12
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
      Format          =   144113667
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
      Format          =   144113667
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil4 
      Height          =   225
      Left            =   8400
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmterima.frx":3E3A
      Caption         =   "frmterima.frx":3E5A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterima.frx":3EC6
      Keys            =   "frmterima.frx":3EE4
      Spin            =   "frmterima.frx":3F26
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
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4800
      Width           =   6015
   End
   Begin VB.Label lblcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lbltype 
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   150
      Width           =   4935
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
   Begin VB.Label Label2 
      Caption         =   "No Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Desc/Reference"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   1950
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Type Transaksi"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmterima"
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

Dim posrow, poscol As String
Dim str99 As String
Dim hitunginout As Boolean

Private Sub cmbtype_Change()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    hapusemua
    cmdsearch1.Enabled = False
    txtcust.Enabled = False
    txtnobukti = ""
    txtnobukti.SetFocus
    
    If cmbtype = "01" And cmbtype.ListIndex = 0 Then
        lbltype = "Produksi Harian Lem (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "PHL0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "PHL0-" & strformat & str99
    End If
    If cmbtype = "01" And cmbtype.ListIndex = 1 Then
        lbltype = "Produksi Harian Karet (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHK0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "PHK0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "PHK0-" & strformat & str99
    End If
    If cmbtype = "02" Then
        lbltype = "Terima Over Zak"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TOZ0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If
        OBJ.Close
        str99 = str99 + 1

        If Len(str99) = 1 Then txtnobukti = "TOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti = "TOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti = "TOZ0-" & strformat & str99
    End If

    If cmbtype = "03" Then
        lbltype = "Keluar Over Zak"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KOZ0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If
        OBJ.Close

        str99 = str99 + 1

        If Len(str99) = 1 Then txtnobukti = "KOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti = "KOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti = "KOZ0-" & strformat & str99
    End If
    If cmbtype = "04" Then
        lbltype = "Terima Retur Customer"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TR0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "TR0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "TR0-" & strformat & str99
        
        Frame1.Visible = True
    End If
    If cmbtype = "05" Then
        lbltype = "Barang Rusak (Out)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'BR0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "BR0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "BR0-" & strformat & str99
    End If
    If cmbtype = "06" Then
        lbltype = "Keluar Sampel"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KS0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "KS0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "KS0-" & strformat & str99
    End If
    If cmbtype = "07" Then
        lbltype = "Barang Bonus (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'BB0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "BB0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "BB0-" & strformat & str99
    End If
    If cmbtype = "08" Then
        lbltype = "Dari Bahan Baku (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'DBB0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "DBB0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "DBB0-" & strformat & str99
    End If
    If cmbtype = "09" Then
        lbltype = "Terima Dari Pabrik (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TDP0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "TDP0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "TDP0-" & strformat & str99
    End If
    If cmbtype = "10" Then
        lbltype = "Retur Ke Pabrik (Out)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'RKP0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "RKP0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "RKP0-" & strformat & str99
    End If
    If cmbtype = "12" Then
        lbltype = "Terima Pinjaman dr Customer (In)"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TPC0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "TPC0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "TPC0-" & strformat & str99
    End If
    If cmbtype = "13" Then
        lbltype = "Kembali Pinjaman dr Customer (Out)"
        cmdsearch1.Enabled = True
        txtcust.Enabled = True
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KPC0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "KPC0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "KPC0-" & strformat & str99
    End If
    If cmbtype = "14" Then
        lbltype = "Keluar Barang Bocor (Out)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KBB0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "KBB0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "KBB0-" & strformat & str99
    End If
    If cmbtype = "15" Then
        lbltype = "Terima Tuangan Lem (In)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TTL0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "TTL0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "TTL0-" & strformat & str99
    End If
    If cmbtype = "16" Then
        lbltype = "Keluar Ke Bahan Baku (Out)"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KKB0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 2)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "KKB0-" & strformat & "0" & str99
        If Len(str99) = 2 Then txtnobukti = "KKB0-" & strformat & str99
    End If
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtnobukti.SetFocus
End Sub

Private Sub cmdadd_Click()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If cmbtype = "11" Then
        MsgBox "Invalid Type.", vbExclamation, "Warning"
        Exit Sub
    End If

    If cmbtype = "" Or txtnobukti = "" Or txtgudang = "" Or grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If (cmbtype = "04" Or cmbtype = "12" Or cmbtype = "13") And txtcust = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not add, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
       
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop
    
    If par1 = "1" Then hitunginout = True Else hitunginout = False
    
    If hitunginout Then
        Label1 = "checking stock on hand ..."
        If cmbtype = "03" Or cmbtype = "05" Or cmbtype = "06" Or cmbtype = "10" Or cmbtype = "13" Or cmbtype = "14" Then
            grid.Row = 1
            Do While grid.TextMatrix(grid.Row, 1) <> ""
                OBJ.Open dsn
                SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
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
                
                txtnil3 = txtnil1 - txtnil2 - txtnil4 - Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
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
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
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
                        Label1 = ""
                        Exit Sub
                    End If
                                
                    If date2 = date3 Then Exit Do
                    
                    date2 = date2 + 1
                Loop
                
                grid.Row = grid.Row + 1
            Loop
        End If
    End If
    
    Label1 = "checking auto numbering format ..."
    OBJ.Open dsn
    SQL = "select * from am_bpbhdr where nobpb = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If cmbtype = "01" And cmbtype.ListIndex = 0 Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "PHL0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "PHL0-" & strformat & str99
        End If
        If cmbtype = "01" And cmbtype.ListIndex = 1 Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHK0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "PHK0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "PHK0-" & strformat & str99
        End If
        If cmbtype = "02" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TOZ0-' + '" + strformat + "%' order by dateentry desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 3)
            Else
                str99 = 0
            End If

            str99 = str99 + 1
            
            If Len(str99) = 1 Then txtnobukti = "TOZ0-" & strformat & "00" & str99
            If Len(str99) = 2 Then txtnobukti = "TOZ0-" & strformat & "0" & str99
            If Len(str99) = 3 Then txtnobukti = "TOZ0-" & strformat & str99
        End If
        If cmbtype = "03" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KOZ0-' + '" + strformat + "%' order by dateentry desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 3)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "KOZ0-" & strformat & "00" & str99
            If Len(str99) = 2 Then txtnobukti = "KOZ0-" & strformat & "0" & str99
            If Len(str99) = 3 Then txtnobukti = "KOZ0-" & strformat & str99
        End If
        If cmbtype = "04" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TR0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "TR0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "TR0-" & strformat & str99
        End If
        If cmbtype = "05" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'BR0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "BR0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "BR0-" & strformat & str99
        End If
        If cmbtype = "06" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KS0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "KS0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "KS0-" & strformat & str99
        End If
        If cmbtype = "07" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'BB0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "BB0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "BB0-" & strformat & str99
        End If
        If cmbtype = "08" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'DBB0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "DBB0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "DBB0-" & strformat & str99
        End If
        If cmbtype = "09" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TDP0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "TDP0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "TDP0-" & strformat & str99
        End If
        If cmbtype = "10" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'RKP0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "RKP0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "RKP0-" & strformat & str99
        End If
        If cmbtype = "12" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TPC0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "TPC0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "TPC0-" & strformat & str99
        End If
        If cmbtype = "13" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KPC0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
            
            If Len(str99) = 1 Then txtnobukti = "KPC0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "KPC0-" & strformat & str99
        End If
        If cmbtype = "14" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KBB0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
            
            If Len(str99) = 1 Then txtnobukti = "KBB0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "KBB0-" & strformat & str99
        End If
        If cmbtype = "15" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TTL0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
            
            If Len(str99) = 1 Then txtnobukti = "TTL0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "TTL0-" & strformat & str99
        End If
        If cmbtype = "16" Then
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KKB0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 2)
            Else
                str99 = 0
            End If
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "KKB0-" & strformat & "0" & str99
            If Len(str99) = 2 Then txtnobukti = "KKB0-" & strformat & str99
        End If
    End If
    OBJ.Close
    
    Label1 = "Inserting data to database ..."
    OBJ.Open dsn
    
    SQL = "Select CURRENT_TIMESTAMP'time'"
    Set RST = OBJ.Execute(SQL)
    Dim waktu As String
    waktu = RST!Time

    SQL = "insert into am_bpbhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nobpb,"
    SQL = SQL + "tglbpb,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "noref,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'" & cmbtype & "',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtgudang & "',"
    If cmbtype = "04" Then
        SQL = SQL + "'" & txtnolot & " : " & txtapply & "',"
    Else
        SQL = SQL + "'" & txtapply & "',"
    End If
    SQL = SQL + "'" & txtcust & "',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    'SQL = SQL + "'" & Format(waktu, "yyyy-mm-dd hh:mm:ss") & "',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
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
        If cmbtype = "03" Or cmbtype = "05" Or cmbtype = "06" Or cmbtype = "10" Or cmbtype = "13" Or cmbtype = "14" Or cmbtype = "16" Then SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") * -1 & "'),"
        If cmbtype = "01" Or cmbtype = "02" Or cmbtype = "04" Or cmbtype = "07" Or cmbtype = "08" Or cmbtype = "09" Or cmbtype = "12" Or cmbtype = "15" Then SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    Label1 = "Proces complete ..."
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    Label1 = ""
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    cmdsearch1.Enabled = False
    txtcust.Enabled = False
    txtnobukti = ""
    cmbtype = ""
    lbltype = ""
    txtnolot = ""
    Frame1.Visible = False
    cmbtype.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
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

Private Sub cmdsearch2_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    carigudang
    hasil = ""
    hasil1 = ""
    txtapply.SetFocus
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
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='101' and b.kodeuser = '1" & kuser & "'"
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
    cmbtype.AddItem "16"

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
    cmbtype.List(15, 1) = "Keluar Ke Bahan Baku (Out)"
    
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
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
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
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
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

Private Sub txtgudang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtapply.SetFocus
End Sub

Private Sub txtgudang_LostFocus()
    carigudang
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
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
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
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
                    
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

Private Sub txtnobukti_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date1.SetFocus
    KeyAscii = 0
End Sub

Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub carigudang()
    If txtgudang = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_gudang where kodegudang = '" & txtgudang & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblgudang = RST!namagudang
    Else
        MsgBox "Gudang " & txtgudang & " Not Found.", vbExclamation, "Warning"
        txtgudang = ""
        lblgudang = ""
        txtgudang.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Or cmbtype = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusemua
    
    OBJ.Open dsn
    SQL = "select * from am_bpbhdr where nobpb = '" & txtnobukti & "' and type = '" & cmbtype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnobukti.SetFocus
        txtnobukti = ""
    End If
    OBJ.Close
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub hapusemua()
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1 = Date
    End If
    
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
