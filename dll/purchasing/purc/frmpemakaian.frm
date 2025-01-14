VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmpemakaian 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Pemakaian Bahan Baku"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
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
   Icon            =   "frmpemakaian.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpemakaian.frx":2372
      Caption         =   "frmpemakaian.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":23FE
      Keys            =   "frmpemakaian.frx":241C
      Spin            =   "frmpemakaian.frx":245E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
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
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
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
      Left            =   3360
      Picture         =   "frmpemakaian.frx":2486
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   120
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
      Picture         =   "frmpemakaian.frx":27D4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   120
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
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
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
      Format          =   135528451
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   7095
      _ExtentX        =   12515
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3360
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
      MICON           =   "frmpemakaian.frx":2AB6
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
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
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
      MICON           =   "frmpemakaian.frx":2DD0
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
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
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
      MICON           =   "frmpemakaian.frx":30EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   225
      Left            =   6240
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3404
      Caption         =   "frmpemakaian.frx":3424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":3490
      Keys            =   "frmpemakaian.frx":34AE
      Spin            =   "frmpemakaian.frx":34F0
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,###,##0.000;(##,###,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,###,##0.000;(##,###,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   225
      Left            =   6240
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3518
      Caption         =   "frmpemakaian.frx":3538
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":35A4
      Keys            =   "frmpemakaian.frx":35C2
      Spin            =   "frmpemakaian.frx":3604
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil3 
      Height          =   225
      Left            =   6240
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":362C
      Caption         =   "frmpemakaian.frx":364C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":36B8
      Keys            =   "frmpemakaian.frx":36D6
      Spin            =   "frmpemakaian.frx":3718
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   3960
      TabIndex        =   17
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
      Format          =   135528451
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil4 
      Height          =   225
      Left            =   6240
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3740
      Caption         =   "frmpemakaian.frx":3760
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":37CC
      Keys            =   "frmpemakaian.frx":37EA
      Spin            =   "frmpemakaian.frx":382C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3960
      TabIndex        =   19
      Top             =   120
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
      Format          =   135528451
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil5 
      Height          =   225
      Left            =   6240
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3854
      Caption         =   "frmpemakaian.frx":3874
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":38E0
      Keys            =   "frmpemakaian.frx":38FE
      Spin            =   "frmpemakaian.frx":3940
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil6 
      Height          =   225
      Left            =   5400
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3968
      Caption         =   "frmpemakaian.frx":3988
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":39F4
      Keys            =   "frmpemakaian.frx":3A12
      Spin            =   "frmpemakaian.frx":3A54
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,###,##0.000;(##,###,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,###,##0.000;(##,###,###,##0.000)"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -99999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil7 
      Height          =   225
      Left            =   5400
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3A7C
      Caption         =   "frmpemakaian.frx":3A9C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":3B08
      Keys            =   "frmpemakaian.frx":3B26
      Spin            =   "frmpemakaian.frx":3B68
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil8 
      Height          =   225
      Left            =   5400
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpemakaian.frx":3B90
      Caption         =   "frmpemakaian.frx":3BB0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpemakaian.frx":3C1C
      Keys            =   "frmpemakaian.frx":3C3A
      Spin            =   "frmpemakaian.frx":3C7C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.000;(##,###,##0.000)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin akProgress.akProgressBar pro1 
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BackColour      =   16744576
      FontColour      =   4210752
      BarColour       =   16761024
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   0
      LineWidth       =   3
      Caption         =   0
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   0
      Alignment       =   2
   End
   Begin akProgress.akProgressBar pro2 
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BackColour      =   16744576
      FontColour      =   4210752
      BarColour       =   16761024
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   0
      LineWidth       =   3
      Caption         =   0
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   0
      Alignment       =   2
   End
   Begin VB.Label Label1 
      Caption         =   "No Order Prod."
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "No BPB"
      Height          =   255
      Left            =   255
      TabIndex        =   12
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal BPB"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   510
      Width           =   1455
   End
End
Attribute VB_Name = "frmpemakaian"
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

Dim posrow As String

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    pro1.Max = grid.Rows - 2
    pro1.Value = 0
    pro1.Visible = True
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 3) = "0.000" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        OBJ1.Open dsn
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbpb < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil1 = RST1!qty
        Else
            txtnil1 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbeli < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil2 = RST1!qty
        Else
            txtnil2 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and dateentry < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil3 = RST1!qty
        Else
            txtnil3 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and tglretur < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil5 = RST1!qty
        Else
            txtnil5 = 0
        End If
        
        SQL1 = "select isnull(sum(qtyawal),0)'qty' from am_invloc where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and tglupdate < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil6 = RST1!qty
        Else
            txtnil6 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut < '" & tanggalpo & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil7 = RST1!qty
        Else
            txtnil7 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut < '" & tanggalpo & "' and (b.type = '02' or b.type = '03')"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil8 = RST1!qty
        Else
            txtnil8 = 0
        End If
        OBJ1.Close
        
        txtnil4 = txtnil6 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8 - Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
        date2 = date1
        date3 = date1
        OBJ1.Open dsn
        SQL1 = "select isnull(max(b.tglbpb),01/01/1900)'tanggal' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglbeli),01/01/1900)'tanggal' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(dateentry),01/01/1900)'tanggal' from am_usesisa where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(tglretur),01/01/1900)'tanggal' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and (b.type = '02' or b.type = '03')"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        OBJ1.Close
        
        pro2.Max = (date3 - date2) + 1
        pro2.Value = 0
        pro2.Visible = True
        Do While True
            OBJ1.Open dsn
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbpb = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbeli = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and dateentry = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil3 = RST1!qty
            Else
                txtnil3 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and tglretur = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil5 = RST1!qty
            Else
                txtnil5 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut = '" & tanggal2 & "' and b.type = '01'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil7 = RST1!qty
            Else
                txtnil7 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut = '" & tanggal2 & "' and (b.type = '02' or b.type = '03')"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil8 = RST1!qty
            Else
                txtnil8 = 0
            End If
            OBJ1.Close
            
            txtnil4 = txtnil4 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
            
            If txtnil4 < 0 Then
                MsgBox "Stock Limited on " & grid.TextMatrix(grid.Row, 2), vbOKOnly + vbExclamation, "Warning"
                Exit Sub
            End If
            pro2.Value = pro2.Value + 1
            
            If date2 = date3 Then Exit Do
            
            date2 = date2 + 1
        Loop
        pro2.Visible = False
        pro1.Value = pro1.Value + 1
        
        grid.Row = grid.Row + 1
    Loop
    pro1.Visible = False

    OBJ.Open dsn
    SQL = "select * from am_usehdr where nobpb = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "insert into am_usehdr ("
    SQL = SQL + "nobpb, "
    SQL = SQL + "tglbpb, "
    SQL = SQL + "noorder)"

    SQL = SQL + " values('" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
    SQL = SQL + "'" & txtsj & "')"
    Set RST = OBJ.Execute(SQL)

    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do

        SQL = "insert into am_uselin ("
        SQL = SQL + "nobpb, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "lineitem, "
        SQL = SQL + "kodesatuan)"

        SQL = SQL + " values('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)

        grid.Row = grid.Row + 1
    Loop
    OBJ.Close

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    
    txtnobukti = ""
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    txtsj = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
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
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='131' and b.kodeuser = '2" & kuser & "'"
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
   
    
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Qty"
    grid.TextMatrix(0, 4) = "K/Sat."
    grid.TextMatrix(0, 5) = "Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 800
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 0
    grid.ColWidth(7) = 0
    
    grid.RowHeightMin = 300
    
    date1.Value = Date
    
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1.MinDate = RST!tanggal1
        date1.MaxDate = RST!tanggal2
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 100 Then
                MsgBox "Maximum line 100", vbExclamation, "Warning"
                Exit Sub
            End If
            If grid.Row <> 1 And grid.TextMatrix(grid.Row - 1, 1) = "" Then Exit Sub
            
            carisql1 = "select a.kodebarang, a.namabarang,"
            carisql1 = carisql1 + "((select isnull(sum(b.qtyuse),0) from am_belilin b where a.kodebarang=b.kodebarang)-"
            carisql1 = carisql1 + "(select isnull(sum(c.qty),0) from am_uselin c where a.kodebarang=c.kodebarang)-"
            carisql1 = carisql1 + "(select isnull(sum(e.qty),0) from am_beliretur e where a.kodebarang=e.kodebarang)+"
            carisql1 = carisql1 + "(select isnull(sum(f.qtyawal),0) from am_invloc f where a.kodebarang=f.kodebarang)+"
            carisql1 = carisql1 + "(select isnull(sum(case when g.type ='01' then g.qty else (g.qty * -1) end),0) from am_mutlin g where a.kodebarang=g.kodebarang)+"
            carisql1 = carisql1 + "(select isnull(sum(d.qty),0) from am_usesisa d where a.kodebarang=d.kodebarang))'onhand'"
            carisql1 = carisql1 + " from am_apitemmst a"
            namatabel = "Stock Bahan Baku"
                
            frmsearch.Show vbModal
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    Select Case grid.Col
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
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
                If grid.TextMatrix(grid.Row, 1) = hasil And posrow <> grid.Row Then
                
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
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            'semua tanggal
            OBJ1.Open dsn
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbpb <= '" & tanggalpo & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbeli <= '" & tanggalpo & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and dateentry <= '" & tanggalpo & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil3 = RST1!qty
            Else
                txtnil3 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and tglretur <= '" & tanggalpo & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil5 = RST1!qty
            Else
                txtnil5 = 0
            End If
            
            SQL1 = "select isnull(sum(qtyawal),0)'qty' from am_invloc where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and tglupdate <= '" & tanggalpo & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                
                    txtnil6 = RST1!qty
              
            Else
                txtnil6 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut <= '" & tanggalpo & "' and b.type = '01'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil7 = RST1!qty
            Else
                txtnil7 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglmut <= '" & tanggalpo & "' and (b.type = '02' or b.type = '03')"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil8 = RST1!qty
            Else
                txtnil8 = 0
            End If
            OBJ1.Close
            
            txtnil4 = txtnil6 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
            
            If txtnil4 < 0 Then
                MsgBox "Minus Stock.", vbExclamation, "Information"
                grid.TextMatrix(grid.Row, 1) = ""
                Exit Sub
            ElseIf txtnil4 = 0 Then
                MsgBox "Stock empty.", vbExclamation, "Information"
                grid.TextMatrix(grid.Row, 1) = ""
                Exit Sub
            End If
                        
            grid.TextMatrix(grid.Row, 6) = Format(txtnil4, "###,##0.000")
            
            OBJ.Open dsn
            SQL = "select a.namabarang,a.kodesatuanmutasi,b.namasatuan from am_apitemmst a left join am_apunit b on a.kodesatuanmutasi=b.kodesatuan where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = "0.000"
                grid.TextMatrix(grid.Row, 4) = RST!kodesatuanmutasi
                grid.TextMatrix(grid.Row, 5) = RST!namasatuan
                
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
                grid.TextMatrix(grid.Row, 7) = ""
            End If
            OBJ.Close
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtnilai > Format(grid.TextMatrix(grid.Row, 6), "general number") Then
            MsgBox "Stock Limited, Qty max = " & grid.TextMatrix(grid.Row, 6), vbExclamation, "Information"
        Else
            grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.000")
        End If
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai_LostFocus
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

Private Sub txtsj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub

    hapusgrid
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    txtsj = ""

    OBJ.Open dsn
    SQL = "select * from am_usehdr where nobpb = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnobukti = ""
        txtnobukti.SetFocus
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
        grid.TextMatrix(grid.Row, 7) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 800
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 0
    grid.ColWidth(7) = 0
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Function tanggalpo()
    tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
