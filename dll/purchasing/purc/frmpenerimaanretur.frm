VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmpenerimaanretur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retur Penerimaan Barang"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
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
   Icon            =   "frmpenerimaanretur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date4 
      Height          =   285
      Left            =   3360
      TabIndex        =   28
      Top             =   840
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
      Format          =   137363459
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   3360
      TabIndex        =   26
      Top             =   480
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
      Format          =   137363459
      CurrentDate     =   37426
   End
   Begin VB.TextBox txtsj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtbukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtkend 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmpenerimaanretur.frx":2372
      Caption         =   "frmpenerimaanretur.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":23FE
      Keys            =   "frmpenerimaanretur.frx":241C
      Spin            =   "frmpenerimaanretur.frx":245E
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
   Begin VB.TextBox txtpo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   2
      Top             =   960
      Width           =   1815
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
      Left            =   5520
      Picture         =   "frmpenerimaanretur.frx":2486
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
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
      Left            =   5760
      Picture         =   "frmpenerimaanretur.frx":27D4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
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
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
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
      Format          =   137363459
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2160
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No LPB"
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
      MICON           =   "frmpenerimaanretur.frx":2AB6
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
      Left            =   6000
      TabIndex        =   11
      Top             =   4440
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
      MICON           =   "frmpenerimaanretur.frx":2DD0
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
      TabIndex        =   10
      Top             =   4440
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
      MICON           =   "frmpenerimaanretur.frx":30EA
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
      Left            =   3120
      TabIndex        =   8
      Top             =   4440
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
      MICON           =   "frmpenerimaanretur.frx":3404
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
      Left            =   4080
      TabIndex        =   9
      Top             =   4440
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
      MICON           =   "frmpenerimaanretur.frx":371E
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
      Left            =   6120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":3A38
      Caption         =   "frmpenerimaanretur.frx":3A58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":3AC4
      Keys            =   "frmpenerimaanretur.frx":3AE2
      Spin            =   "frmpenerimaanretur.frx":3B24
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
      Left            =   6120
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":3B4C
      Caption         =   "frmpenerimaanretur.frx":3B6C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":3BD8
      Keys            =   "frmpenerimaanretur.frx":3BF6
      Spin            =   "frmpenerimaanretur.frx":3C38
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
      Left            =   6120
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":3C60
      Caption         =   "frmpenerimaanretur.frx":3C80
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":3CEC
      Keys            =   "frmpenerimaanretur.frx":3D0A
      Spin            =   "frmpenerimaanretur.frx":3D4C
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
      Left            =   6120
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":3D74
      Caption         =   "frmpenerimaanretur.frx":3D94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":3E00
      Keys            =   "frmpenerimaanretur.frx":3E1E
      Spin            =   "frmpenerimaanretur.frx":3E60
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
      Left            =   3360
      TabIndex        =   25
      Top             =   120
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
      Format          =   137363459
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil5 
      Height          =   225
      Left            =   6120
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":3E88
      Caption         =   "frmpenerimaanretur.frx":3EA8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":3F14
      Keys            =   "frmpenerimaanretur.frx":3F32
      Spin            =   "frmpenerimaanretur.frx":3F74
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Retur"
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
      MICON           =   "frmpenerimaanretur.frx":3F9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil7 
      Height          =   225
      Left            =   5280
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":42B6
      Caption         =   "frmpenerimaanretur.frx":42D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":4342
      Keys            =   "frmpenerimaanretur.frx":4360
      Spin            =   "frmpenerimaanretur.frx":43A2
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
      MinValue        =   0
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
   Begin TDBNumber6Ctl.TDBNumber txtnil8 
      Height          =   225
      Left            =   5280
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":43CA
      Caption         =   "frmpenerimaanretur.frx":43EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":4456
      Keys            =   "frmpenerimaanretur.frx":4474
      Spin            =   "frmpenerimaanretur.frx":44B6
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
      MinValue        =   0
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
   Begin TDBNumber6Ctl.TDBNumber txtnil6 
      Height          =   225
      Left            =   6120
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanretur.frx":44DE
      Caption         =   "frmpenerimaanretur.frx":44FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanretur.frx":456A
      Keys            =   "frmpenerimaanretur.frx":4588
      Spin            =   "frmpenerimaanretur.frx":45CA
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
   Begin akProgress.akProgressBar pro1 
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      TabIndex        =   34
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "No Surat Jalan"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label31 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No P.O."
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Retur"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   23
      Top             =   4200
      Width           =   7095
   End
End
Attribute VB_Name = "frmpenerimaanretur"
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

Dim str1 As String
Dim i As Integer
Dim posrow As String

Private Sub cmdadd_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If Trim(txtnobukti) = Trim(txtbukti) Then
        MsgBox "NoRetur dan NoLPB tidak boleh sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtpo = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalpo & "' and tanggal2 >= '" & tanggalpo & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not save, the return date is out of range. (Period on Process)", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
        
    If MsgBox("Are you sure want to retur ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_beliretur where noretur = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    grid.Row = 1
    pro1.Max = grid.Rows - 2
    pro1.Value = 0
    pro1.Visible = True
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 3) = "0.00" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        OBJ1.Open dsn
        SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and nobeli = '" & txtnobukti & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil1 = RST1!qty
        Else
            txtnil1 = 0
        End If

        SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli = '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil2 = RST1!qty
        Else
            txtnil2 = 0
        End If
        OBJ1.Close
        If txtnil2 - txtnil1 < Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) Then
            MsgBox "Can not retur, quantity item Limited. (penerimaan < retur)", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        
        '== cek udakepake brapa/diterima brapa/sisa brapa ==
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
        
        txtnil4 = txtnil6 + txtnil2 - txtnil5 - txtnil1 + txtnil3 + txtnil7 - txtnil8 - Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
        
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
                MsgBox "Can not save data, quantity item Limited.", vbOKOnly + vbExclamation, "Warning"
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
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        OBJ1.Open dsn
        SQL1 = "select b.kodesupp,b.kodecur,b.nilaikurs,a.price from am_polin a left join am_pohdr b on a.nopo=b.nopo where b.nopo = '" & txtpo & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        SQL = "insert into am_beliretur ("
        SQL = SQL + "noretur, "
        SQL = SQL + "tglretur, "
        SQL = SQL + "nobeli, "
        SQL = SQL + "nopo, "
        SQL = SQL + "kodesupp, "
        SQL = SQL + "kodecur, "
        SQL = SQL + "nilaikurs, "
        SQL = SQL + "keterangan, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "price, "
        SQL = SQL + "kodesatuan, "
        SQL = SQL + "lineitem, "
        SQL = SQL + "flag1, "
        SQL = SQL + "flag2)"
        
        SQL = SQL + " values ('" & txtbukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "'',"
        SQL = SQL + "'" & RST1!kodesupp & "',"
        SQL = SQL + "'" & RST1!kodecur & "',"
        SQL = SQL + "convert(money,'" & RST1!nilaikurs & "'),"
        SQL = SQL + "'" & txtkend & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "convert(money,'" & RST1!Price & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'0',"
        SQL = SQL + "'0')"
        Set RST = OBJ.Execute(SQL)
        
        OBJ1.Close
        
        grid.Row = grid.Row + 1
    Loop
    'SIMPAN KE am_belireturtemp

    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        OBJ1.Open dsn
        SQL1 = "Select Qty From am_beliapp where NoPO = '" & txtpo & "' and KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
    
        SQL = "Insert into am_belireturtemp ("
        SQL = SQL + "noretur, "
        SQL = SQL + "nobeli, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "qtyuse)"
        
        SQL = SQL + " Values ('" & txtbukti & "',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & RST1!qty - Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "convert(money,'" & RST1!qty - Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        'SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        Set RST = OBJ.Execute(SQL)
        
        OBJ1.Close
        
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    
    txtbukti.Enabled = True
    cmdsearch.Enabled = True
    date1.Enabled = True
    txtbukti = ""
    txtnobukti = ""
    date1.Value = Date
    txtpo = ""
    txtsj = ""
    txtkend = ""
    
    txtbukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtbukti = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_beliretur where noretur = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data not found.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_period where tanggal1 <= '" & tanggalpo & "' and tanggal2 >= '" & tanggalpo & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
        
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_beliretur where noretur = '" & txtbukti & "' and flag1 = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete, data already export.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
        
    OBJ.Open dsn
    SQL = "delete am_beliretur where noretur = '" & txtbukti & "'"
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
        carisql1 = "select distinct noretur, convert(char(11),tglretur)'tglretur' from am_beliretur where tglretur >= '" & batas1 & "' and tglretur <= '" & batas2 & "'"
    Else
        carisql1 = "select distinct noretur, convert(char(11),tglretur)'tglretur' from am_beliretur"
    End If
    namatabel = "Retur Penerimaan Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    carinvoice
    hasil = ""
    hasil1 = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobeli, convert(char(11),tglbeli)'tglbeli' from am_belihdr where tglbeli <= '" & tanggalpo & "' and tglbeli >= '" & batas1 & "' and tglbeli <= '" & batas2 & "'"
    Else
        carisql1 = "select nobeli, convert(char(11),tglbeli)'tglbeli' from am_belihdr where tglbeli <= '" & tanggalpo & "'"
    End If
    namatabel = "Penerimaan Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    hasil = ""
    hasil1 = ""
    
    OBJ.Open dsn
    SQL = "select nopo,nosj from am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtpo = RST!nopo
        txtsj = RST!nosj
    End If
    OBJ.Close
    txtpo.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='251' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='253' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
        
    '   If cmdel.Enabled = False And cmdadd.Enabled = False Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
        
    '       Unload Me
    '    End If
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Retur"
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
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
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
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            If grid.Row <> 1 And grid.TextMatrix(grid.Row - 1, 1) = "" Then Exit Sub
            
            carisql1 = "select a.kodebarang, a.kodesatuanuse, b.namabarang from am_belilin a left join am_apitemmst b on a.kodebarang=b.kodebarang where a.nobeli = '" & txtnobukti & "'"
            namatabel = "Item on LPB"
                
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
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
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
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            
            OBJ1.Open dsn
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and nobeli = '" & txtnobukti & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli = '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            OBJ1.Close
            
            If txtnil1 - txtnil2 = 0 Then
                MsgBox "Item return complete.", vbExclamation, "Information"
                
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 4) = ""
            
                Exit Sub
            End If
            
            grid.TextMatrix(grid.Row, 6) = Format(txtnil2 - txtnil1, "###,##0.00")
            
            OBJ.Open dsn
            SQL = "select a.namabarang from am_apitemmst a where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = "0.00"
                
                SQL = "select b.namasatuan from am_apunit b where b.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then grid.TextMatrix(grid.Row, 5) = RST!namasatuan Else grid.TextMatrix(grid.Row, 5) = ""
                
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
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    carinvoice
End Sub

Private Sub txtkend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtnilai > Val(Format(grid.TextMatrix(grid.Row, 6), "general number")) Then
            MsgBox "Quantity return max = " & grid.TextMatrix(grid.Row, 6), vbExclamation, "Information"
        Else
            grid.TextMatrix(grid.Row, 3) = Format(txtnilai, "###,###,##0.00")
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

Private Sub txtnobukti_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtpo.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtpo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsj.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsj_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtsj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkend.SetFocus
    KeyAscii = 0
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

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Function tanggalpo()
      tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub carinvoice()
    If txtbukti = "" Then Exit Sub
    If txtbukti.SelLength <> 0 Then Exit Sub

    hapusgrid
    date1 = Date
    txtnobukti = ""
    txtpo = ""
    txtsj = ""
    txtkend = ""

    OBJ.Open dsn
    SQL = "select * from am_beliretur where noretur = '" & txtbukti & "' order by lineitem"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglretur
        txtnobukti = RST!NoBeli
        txtkend = RST!keterangan
        
        OBJ1.Open dsn
        SQL1 = "select nopo,nosj from am_belihdr where nobeli = '" & txtnobukti & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtpo = RST1!nopo
            txtsj = RST1!nosj
        End If
        OBJ1.Close

        grid.Row = 1
        Do While Not RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan

            OBJ1.Open dsn
            SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 2) = RST1!namabarang

            SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and nobeli = '" & txtnobukti & "' and noretur <> '" & txtbukti & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli = '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            OBJ1.Close
            
            grid.TextMatrix(grid.Row, 6) = Format(txtnil2 - txtnil1, "###,##0.00")

            SetRow grid.Row, True
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
        txtbukti.Enabled = False
        cmdsearch.Enabled = False
        date1.Enabled = False
        txtnobukti.SetFocus
    Else
        If MsgBox("Data not found, add new data?", vbQuestion + vbYesNo, "Question") = vbNo Then
            txtbukti = ""
            txtbukti.SetFocus
        Else
            date1.SetFocus
        End If
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

