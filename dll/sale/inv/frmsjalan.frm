VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsjalan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Surat Jalan"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
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
   Icon            =   "frmsjalan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date4 
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
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
      Format          =   135069699
      CurrentDate     =   37426
   End
   Begin VB.TextBox txtso 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtgudang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtvia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   7
      Top             =   2640
      Width           =   7575
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmsjalan.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":23DE
      Key             =   "frmsjalan.frx":23FC
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
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmsjalan.frx":2438
      Caption         =   "frmsjalan.frx":2458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":24C4
      Keys            =   "frmsjalan.frx":24E2
      Spin            =   "frmsjalan.frx":2524
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
   Begin VB.TextBox txtsales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1920
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
      Left            =   8040
      Picture         =   "frmsjalan.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
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
      Left            =   8040
      Picture         =   "frmsjalan.frx":289A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
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
      Left            =   8040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
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
      Format          =   135069699
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   9255
      _ExtentX        =   16325
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   5520
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
      MICON           =   "frmsjalan.frx":2B7C
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
      TabIndex        =   13
      Top             =   5520
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
      MICON           =   "frmsjalan.frx":2E96
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
      TabIndex        =   12
      Top             =   5520
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
      MICON           =   "frmsjalan.frx":31B0
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
      TabIndex        =   29
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmsjalan.frx":34CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   31
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Sales Order"
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
      MICON           =   "frmsjalan.frx":37E4
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
      Left            =   8280
      TabIndex        =   34
      Top             =   840
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsjalan.frx":3AFE
      Caption         =   "frmsjalan.frx":3B1E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":3B8A
      Keys            =   "frmsjalan.frx":3BA8
      Spin            =   "frmsjalan.frx":3BEA
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
      Left            =   8280
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsjalan.frx":3C12
      Caption         =   "frmsjalan.frx":3C32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":3C9E
      Keys            =   "frmsjalan.frx":3CBC
      Spin            =   "frmsjalan.frx":3CFE
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
      Left            =   8280
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsjalan.frx":3D26
      Caption         =   "frmsjalan.frx":3D46
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":3DB2
      Keys            =   "frmsjalan.frx":3DD0
      Spin            =   "frmsjalan.frx":3E12
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
   Begin TDBNumber6Ctl.TDBNumber txtnil4 
      Height          =   225
      Left            =   8280
      TabIndex        =   39
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsjalan.frx":3E3A
      Caption         =   "frmsjalan.frx":3E5A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalan.frx":3EC6
      Keys            =   "frmsjalan.frx":3EE4
      Spin            =   "frmsjalan.frx":3F26
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   7320
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
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
      Format          =   135069699
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   7320
      TabIndex        =   41
      Top             =   2280
      Visible         =   0   'False
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
      Format          =   135069699
      CurrentDate     =   37426
   End
   Begin VB.Label Label5 
      Caption         =   "Salesman"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Sales Order"
      Height          =   255
      Left            =   3120
      TabIndex        =   33
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label lblso 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   5040
      TabIndex        =   32
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblalamatcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label lblsat 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5730
      Width           =   5385
   End
   Begin VB.Label Label3 
      Caption         =   "Kirim Via"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Kirim"
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   24
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "No Surat Jalan"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "No PO Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal SJ"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Barang :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5490
      Width           =   5385
   End
End
Attribute VB_Name = "frmsjalan"
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

Dim CMD As New ADODB.Command
Dim vcmd(2) As Variant

Dim posrow, poscol, str99, str21, str1, str2 As String
Dim int3 As Integer
Dim boo1, hitunginout As Boolean
Dim strsj As String

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    txtnobukti = Trim(txtnobukti)
    
    If txtnobukti = "" Or txtsales = "" Or txtgudang = "" Or txtkodecust = "" Or txtso = "" Then
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
    
    If date4 < date1 Then
        MsgBox "Tanggal Kirim tidak boleh lebih kecil dari Tanggal Surat Jalan.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtapply = "" Then
        If MsgBox("Continue with blank PO number ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
         If grid.TextMatrix(grid.Row, 2) = "" Then
            MsgBox "Kolom keterangan harus diisi nomor lot", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Or Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) < Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_soapp where noso = '" & txtso & "' and tglso <= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Tanggal Surat Jalan Invalid.", vbExclamation, "Information"
                
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If par1 = "1" Then hitunginout = True Else hitunginout = False
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        'cek outstanding sj vs so
        If grid.TextMatrix(grid.Row, 3) <> "0.00" Then
            OBJ.Open dsn
            SQL = "select qty,bn from am_soapp where noso = '" & txtso & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                OBJ1.Open dsn
                SQL1 = "select isnull(sum(a.qty),0)'qtysj' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.noso = '" & txtso & "' and b.nosj <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) > RST!qty - RST1!qtysj Then
                    MsgBox "Line " & grid.Row & vbCrLf & _
                    "Sales Order - Sum of Surat Jalan, Qty max = " & (RST!qty - RST1!qtysj), vbExclamation, "Information"
            
                    OBJ.Close
                    OBJ1.Close
                    Exit Sub
                End If
                OBJ1.Close
            
                If grid.TextMatrix(grid.Row, 5) <> "0.00" Then
                    OBJ1.Open dsn
                    SQL1 = "select isnull(sum(a.bn),0)'bnsj' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.noso = '" & txtso & "' and b.nosj <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                    Set RST1 = OBJ1.Execute(SQL1)
                    If Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) > RST!bn - RST1!bnsj Then
                        MsgBox "Line " & grid.Row & vbCrLf & _
                        "Bonus Sales Order - Sum of Bonus Surat Jalan, Qty max = " & (RST!bn - RST1!bnsj), vbExclamation, "Information"
            
                        OBJ.Close
                        OBJ1.Close
                        Exit Sub
                    End If
                    OBJ1.Close
                End If
            End If
            OBJ.Close
        End If
        
        If hitunginout Then
            'check stock start
            
            'check stock dari mutasi (ok)
            OBJ.Open dsn
            SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
            
            'check stock dari surat jalan
            If par5 = "0" Then
                SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Else
                SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggal4 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            End If
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
            OBJ.Close
            
            txtnil3 = txtnil1 - txtnil2 - Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
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
                OBJ.Close
                
                txtnil3 = txtnil3 + txtnil1 - txtnil2
                
                If txtnil3 < 0 Then
                    MsgBox "Stock Limited on " & grid.TextMatrix(grid.Row, 1), vbOKOnly + vbExclamation, "Warning"
                    Exit Sub
                End If
                            
                If date2 = date3 Then Exit Do
                
                date2 = date2 + 1
            Loop
            'check stock end
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    int3 = 0
    OBJ.Open dsn
    SQL = "select nosj from am_sjhdr where nosj = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
            
        If Asc(Mid(txtnobukti, 1, 1)) = 80 Then
            OBJ.Open dsn
            SQL = "select top 1 nosj from am_sjhdr where nosj like 'PP" & str1 & "%' order by nosj desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nosj, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "PP" & str1 & "000" & str99
            If Len(str99) = 2 Then txtnobukti = "PP" & str1 & "00" & str99
            If Len(str99) = 3 Then txtnobukti = "PP" & str1 & "0" & str99
            If Len(str99) = 4 Then txtnobukti = "PP" & str1 & str99
        ElseIf Asc(Mid(txtnobukti, 1, 1)) = 76 Then
            OBJ.Open dsn
            SQL = "select top 1 nosj from am_sjhdr where nosj like 'LL" & str2 & "%' order by nosj desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nosj, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "LL" & str2 & "000" & str99
            If Len(str99) = 2 Then txtnobukti = "LL" & str2 & "00" & str99
            If Len(str99) = 3 Then txtnobukti = "LL" & str2 & "0" & str99
            If Len(str99) = 4 Then txtnobukti = "LL" & str2 & str99
        End If
        If Left(txtnobukti, 3) = "PPP" Then txtnobukti = Mid(txtnobukti, 2, 8)
        If Left(txtnobukti, 3) = "LLL" Then txtnobukti = Mid(txtnobukti, 2, 8)
        int3 = 1
    Else
        OBJ.Close
    End If
        
    OBJ.Open dsn
    SQL = "select * from am_sjhdr where nosj = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "insert into am_sjhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nosj,"
    SQL = SQL + "tglsj,"
    SQL = SQL + "kodecust,"
    SQL = SQL + "kodesales,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "nopo,"
    SQL = SQL + "noso,"
    SQL = SQL + "tglkirim,"
    SQL = SQL + "via,"
    SQL = SQL + "via2,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'01',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtkodecust & "',"
    SQL = SQL + "'" & txtsales & "',"
    SQL = SQL + "'" & txtgudang & "',"
    SQL = SQL + "'" & txtapply & "',"
    SQL = SQL + "'" & txtso & "',"
    SQL = SQL + "convert(datetime,'" & tanggal4 & "'),"
    SQL = SQL + "'" & txtvia & "',"
    SQL = SQL + "'0',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "insert into am_sjlin ("
        SQL = SQL + "nosj,"
        SQL = SQL + "tglsj,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "qtysj,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan,"
        SQL = SQL + "BN)"
        
        SQL = SQL + " values("
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    'update nosj
    'SQL = "Update am_sjlot set nosj='" & txtnobukti & "' Where nosj = '" & strsj & "'"
    'Set RST = OBJ.Execute(SQL)
    
    OBJ.Close
    
    If int3 = 1 Then
        MsgBox "Data already exist, data was saved with next number " & txtnobukti & vbCrLf & _
        "Click OK To Continue ...", vbExclamation, "Warning"
    Else
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    
    If par4 = "1" Then
        setup1 = txtnobukti
        frmsjalanshow.Show vbModal
    End If
        
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    'carisql1 = "Where kodeGudang <> 'G3'"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    carigudang
    hasil = ""
    hasil1 = ""
    txtvia.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    If txtnobukti = "" Then Exit Sub
    strsj = txtnobukti
    CMD.ActiveConnection = dsn
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "am_cancellist"
    vcmd(0) = GetTheComputerName
    vcmd(1) = Mid(txtnobukti, 1, 1)
    vcmd(2) = Format(date1, "yyyyMMdd")
    CMD.Execute , vcmd
    Set CMD = Nothing
    
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select noso, convert(char(11),tglso)'tglso',namacust"
        carisql1 = carisql1 + " from am_sjlist"
        carisql1 = carisql1 + " where tglso <= '" & tanggalinv & "' and tglso >= '" & batas1 & "' and tglso <= '" & batas2 & "'"
        carisql1 = carisql1 + " and idkomp = '" & GetTheComputerName & "'"
    Else
        carisql1 = "select noso, convert(char(11),tglso)'tglso',namacust"
        carisql1 = carisql1 + " from am_sjlist"
        carisql1 = carisql1 + " where tglso <= '" & tanggalinv & "' and idkomp = '" & GetTheComputerName & "'"
    End If
    namatabel = "Sales Order "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtso = hasil
    cariso
    hasil = ""
    hasil1 = ""
    txtapply.SetFocus
End Sub

Private Sub date1_Change()
    date4 = date1
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Keterangan"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "Satuan"
    grid.TextMatrix(0, 5) = "BN"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
    grid.ColWidth(5) = 800
    grid.ColWidth(6) = 0
    grid.RowHeightMin = 300
    
    date1.Value = Date
    date4.Value = Date
    
    OBJ.Open dsn
    SQL = "select id4, kode4, kode1 from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!id4 = "1" Then
            str1 = RST!kode4
            str2 = RST!kode1
        Else
            str1 = "00"
            str2 = "00"
        End If
    Else
        str1 = "00"
        str2 = "00"
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblitem = "    Nama Barang : " & RST!namabarang
            lblsat = "    Nama Satuan : " & RST!namasatuan
            lotbarang4 = RST!namasatuan
        Else
            lblitem = "    Nama Barang : "
            lblsat = "    Nama Satuan : "
            lotbarang4 = ""
        End If
        OBJ.Close
    End If
    If txtnobukti = "" Or txtkodecust = "" Or txtgudang = "" Or txtsales = "" Or txtso = "" Then Exit Sub
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    'Hapus nolot barang
                    OBJ.Open dsn
                    SQL = "Delete From am_sjlot Where kode='" & grid.TextMatrix(grid.Row, 6) & "' and kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
                    Set RST = OBJ.Execute(SQL)
                    OBJ.Close
                    
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If grid.Row <> 1 And grid.TextMatrix(grid.Row - 1, 1) = "" Then Exit Sub
        
            carisql1 = "select a.kodebarang, a.kodesatuan, b.namabarang from am_soapp a left join am_itemdtl b on a.kodebarang=b.kodebarang and a.kodesatuan=b.kodesatuan where a.noso = '" & txtso & "' and a.flag2<>'9'"
            namatabel = "Item on Sales Order"
            setup3 = txtgudang
        
            frmsearch.Show vbModal
            setup3 = ""
        Case 2
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            'txtket.Width = grid.ColWidth(grid.Col) - 40
            'txtket = grid.TextMatrix(grid.Row, grid.Col)
            'txtket.Left = grid.Left + grid.CellLeft
            'txtket.Top = grid.Top + grid.CellTop + 20
            'txtket.Visible = True
            'txtket.SetFocus
            
        'Exit Sub
            If grid.TextMatrix(grid.Row, 3) = "" Then
                MsgBox "Isi Qty Terlebih dahulu", vbExclamation, AppName
                Exit Sub
            End If
            If grid.TextMatrix(grid.Row, 2) <> "" Then
                'untuk overwrite hapus terlebih dahulu lot sebelumnya
                OBJ.Open dsn
                SQL = "Delete From am_sjlot Where kode='" & grid.TextMatrix(grid.Row, 6) & "' and kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST = OBJ.Execute(SQL)
                OBJ.Close
                grid.TextMatrix(grid.Row, 2) = ""
            End If
            lotbarang = grid.TextMatrix(grid.Row, 1)
            lotbarang1 = lblitem
            lotbarang2 = grid.TextMatrix(grid.Row, 3)
            lotbarang3 = txtnobukti
            lotbarang4 = grid.TextMatrix(grid.Row, 4)
            
            frmaddlot.Show vbModal
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            OBJ.Open dsn
            SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If RST!kodeproduk = "C999" Then
                    OBJ.Close
                    Exit Sub
                Else
                    OBJ.Close
                End If
            Else
                OBJ.Close
                Exit Sub
            End If

            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6
            'INSERT NOLOT
            carisql1 = "Select a.kodebarang,a.namabarang,a.qtyin,b.NamaSatuan,a.palet,a.nolot"
            carisql1 = carisql1 + " From am_stok a inner join am_unit b on a.kodesatuan = b.KodeSatuan"
            carisql1 = carisql1 + " Where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            namatabel = "Item on Lot"
            frmsearch.Show vbModal
            
            'pakai form aja seperti lot base
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtkodecust = "" Or txtgudang = "" Or txtsales = "" Or txtso = "" Then Exit Sub
    Select Case grid.Col
    Case 2
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
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
            
        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 5
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        OBJ.Open dsn
        SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!kodeproduk = "C999" Then
                OBJ.Close
                Exit Sub
            Else
                OBJ.Close
            End If
        Else
            OBJ.Close
            Exit Sub
        End If

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

    Select Case grid.Col
        Case 1
            grid.Row = 1
            Do While True
                If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                If grid.TextMatrix(grid.Row, 1) = hasil And grid.TextMatrix(grid.Row, 4) = hasil1 And posrow <> grid.Row Then
                
                    MsgBox "Item already exist.", vbInformation, "Information"
                    hasil = ""
                    hasil1 = ""
                    grid.Row = posrow
                    grid.Col = 1
                    grid.SetFocus
                    Exit Sub
                End If
                grid.Row = grid.Row + 1
            Loop
            
            grid.Row = posrow
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = hasil
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""

            OBJ1.Open dsn
            SQL1 = "select qty,keterangan,bn from am_soapp where noso = '" & txtso & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
                txtnil3 = RST1!bn
                grid.TextMatrix(grid.Row, 2) = RST1!keterangan
            Else
                txtnil1 = 0
                txtnil3 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty',isnull(sum(a.bn),0)'bn' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where b.noso = '" & txtso & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "' and (c.flag2 is null or c.flag2 <> '9')"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
                txtnil4 = RST1!bn
            Else
                txtnil2 = 0
                txtnil4 = 0
            End If
            OBJ1.Close
                        
            If txtnil1 - txtnil2 = 0 And txtnil3 - txtnil4 = 0 Then
                MsgBox "Sales Order required is complete.", vbExclamation, "Information"
                
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = ""
                grid.TextMatrix(grid.Row, 5) = ""
                grid.TextMatrix(grid.Row, 6) = ""
                Exit Sub
            End If
                        
            OBJ.Open dsn
            SQL = "SELECT * FROM AM_ITEMDTL WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 3) = Format(txtnil1 - txtnil2, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 5) = Format(txtnil3 - txtnil4, "###,###,##0.00")
                
                SetRow grid.Row, True
                If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                 
                grid.SetFocus
                grid.Col = 2
            Else
                MsgBox "Item Not Found", vbExclamation, "Warning"
                
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = ""
                grid.TextMatrix(grid.Row, 5) = ""
                grid.TextMatrix(grid.Row, 6) = ""
            End If
            OBJ.Close
        Case 2
            If hasil = "" Then
                'MsgBox kdlot
                grid.TextMatrix(grid.Row, 2) = ""
            Else
                grid.TextMatrix(grid.Row, 2) = hasil
                grid.TextMatrix(grid.Row, 6) = hasil1
            End If
            hasil = ""
    End Select
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub


Private Sub txtapply_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date4.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtgudang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtvia.SetFocus
End Sub

Private Sub txtgudang_LostFocus()
    carigudang
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 2
                grid.SetFocus
                grid.Col = 2
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 2) = txtket
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

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtnilai.Visible = False
    ElseIf KeyAscii = 13 Then
        If grid.Col = 3 Then
            OBJ1.Open dsn
            SQL1 = "select qty from am_soapp where noso = '" & txtso & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                OBJ2.Open dsn
                SQL2 = "select isnull(sum(a.qty),0)'qtysj' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.noso = '" & txtso & "' and b.nosj <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If txtnilai > RST1!qty - RST2!qtysj Then
                    MsgBox "Sales Order - Sum of Surat Jalan, Qty max = " & (RST1!qty - RST2!qtysj), vbExclamation, "Information"
                    
                    OBJ1.Close
                    OBJ2.Close
                    GoTo bawah
                End If
                OBJ2.Close
            End If
            OBJ1.Close
        ElseIf grid.Col = 5 Then
            OBJ1.Open dsn
            SQL1 = "select bn from am_soapp where noso = '" & txtso & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                OBJ2.Open dsn
                SQL2 = "select isnull(sum(a.bn),0)'bnsj' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.noso = '" & txtso & "' and b.nosj <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If txtnilai > RST1!bn - RST2!bnsj Then
                    MsgBox "Bonus Sales Order - Sum of Bonus Surat Jalan, Bonus max = " & (RST1!bn - RST2!bnsj), vbExclamation, "Information"
                    
                    OBJ1.Close
                    OBJ2.Close
                    GoTo bawah
                End If
                OBJ2.Close
            End If
            OBJ1.Close
        End If
        
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
    If KeyCode = 46 And par2 = "1" Then txtnobukti = ""
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date1.SetFocus
    If Len(txtnobukti) > 1 And par2 = "1" Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    par2 = "1"
    If KeyAscii = 80 And par2 = "1" Then
        OBJ.Open dsn
        SQL = "select top 1 nosj from am_sjhdr where nosj like 'PP" & str1 & "%' order by nosj desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nosj, 4)
        Else
            str99 = 0
        End If
        OBJ.Close

        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "PP" & str1 & "000" & str99
        If Len(str99) = 2 Then txtnobukti = "PP" & str1 & "00" & str99
        If Len(str99) = 3 Then txtnobukti = "PP" & str1 & "0" & str99
        If Len(str99) = 4 Then txtnobukti = "PP" & str1 & str99
    ElseIf KeyAscii = 76 And par2 = "1" Then
        OBJ.Open dsn
        SQL = "select top 1 nosj from am_sjhdr where nosj like 'LL" & str2 & "%' order by nosj desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nosj, 4)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "LL" & str2 & "000" & str99
        If Len(str99) = 2 Then txtnobukti = "LL" & str2 & "00" & str99
        If Len(str99) = 3 Then txtnobukti = "LL" & str2 & "0" & str99
        If Len(str99) = 4 Then txtnobukti = "LL" & str2 & str99
    ElseIf KeyAscii <> 76 And KeyAscii <> 80 And par2 = "1" Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtnobukti_KeyUp(KeyCode As Integer, Shift As Integer)
    If Left(txtnobukti, 3) = "PPP" Then txtnobukti = Mid(txtnobukti, 2, 8)
    If Left(txtnobukti, 3) = "LLL" Then txtnobukti = Mid(txtnobukti, 2, 8)
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Private Sub txtso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtso_LostFocus()
    cariso
End Sub

Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggal4()
    tanggal4 = Month(date4) & "/" & Day(date4) & "/" & Year(date4)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
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

Private Sub cariso()
    If txtso = "" Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    lblso = ""
    txtkodecust = ""
    lblnamacust = ""
    lblalamatcust = ""
    txtapply = ""
    txtvia = ""
    txtgudang = ""
    lblgudang = ""
    txtsales = ""
    lblsales = ""
    date4 = date1
    
    hapusgrid
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
    lbltotal.Caption = "    Total Barang : 0"
    
    If Mid(txtnobukti, 1, 1) <> Mid(txtso, 1, 1) Then
        MsgBox "Invalid Sales Order.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    
    'periksa limit customer baru
    
    SQL = "select top 1 * from am_soapp where noso = '" & txtso & "' and flag2 <> '9'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblso = Format(RST!tglso, "dd-MM-yyyy")
        txtkodecust = RST!kodecust
        txtapply = RST!nopo
        txtsales = RST!kodesales
        txtvia = RST!jasa
            
        SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
        Set RST = OBJ.Execute(SQL)
        lblnamacust = RST!namacust
        lblalamatcust = RST!alamatcust
        
        SQL = "select * from am_salesman where kodesales = '" & txtsales & "'"
        Set RST = OBJ.Execute(SQL)
        lblsales = RST!namasales
        OBJ.Close
    Else
        MsgBox "Sales Order Not Found.", vbExclamation, "Warning"
        OBJ.Close
        txtso = ""
        cmdsearch3_Click
        If hasil <> "" Then
            txtso = hasil
            
            OBJ.Open dsn
            SQL = "select * from am_soapp where noso = '" & txtso & "'"
            Set RST = OBJ.Execute(SQL)
            lblso = Format(RST!tglso, "dd-MM-yyyy")
            txtkodecust = RST!kodecust
            txtapply = RST!nopo
            txtsales = RST!kodesales
            txtvia = RST!jasa
                
            SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
            Set RST = OBJ.Execute(SQL)
            lblnamacust = RST!namacust
            lblalamatcust = RST!alamatcust
            
            SQL = "select * from am_salesman where kodesales = '" & txtsales & "'"
            Set RST = OBJ.Execute(SQL)
            lblsales = RST!namasales
            OBJ.Close
            
            hasil = ""
            hasil1 = ""
        End If
    End If
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusemua

    OBJ.Open dsn
    SQL = "select * from am_sjhdr where nosj = '" & txtnobukti & "'"
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
    
    txtkodecust = ""
    lblnamacust = ""
    lblalamatcust = ""
    txtsales = ""
    lblsales = ""
    txtgudang = ""
    lblgudang = ""
    txtapply = ""
    txtvia = ""
    txtso = ""
    lblso = ""
    date4 = Date
    
    hapusgrid
    
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
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
        grid.TextMatrix(grid.Row, 6) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 0
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
    If grid.Rows = 2 Then
        lbltotal.Caption = "    Total Barang : 0"
    Else
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    End If
End Sub
