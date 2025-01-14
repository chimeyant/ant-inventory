VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmpenerimaanedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Penerimaan Barang"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
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
   Icon            =   "frmpenerimaanedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Note"
      Height          =   2175
      Left            =   6960
      TabIndex        =   32
      Top             =   120
      Width           =   2895
      Begin VB.Label Label5 
         Caption         =   "Qty (PO) : Quantity diisi sesuai dengan quantity yang di PO."
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Qty (USE) : Quantity diisi sesuai dengan quantity yang nantinya dipakai di modul Pemakaian."
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Jika Qty (PO) dengan Qty (Use) adalah sama maka Satuan (PO) dan Satuan (Use) diharuskan sama juga."
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   2655
      End
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   4200
      TabIndex        =   27
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
      Format          =   134283267
      CurrentDate     =   37426
   End
   Begin VB.TextBox txtsj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtdriver 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox txtkend 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1680
      Width           =   5415
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpenerimaanedit.frx":2372
      Caption         =   "frmpenerimaanedit.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":23FE
      Keys            =   "frmpenerimaanedit.frx":241C
      Spin            =   "frmpenerimaanedit.frx":245E
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
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   0
      Top             =   120
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
      Left            =   3960
      Picture         =   "frmpenerimaanedit.frx":2486
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
      Left            =   4200
      Picture         =   "frmpenerimaanedit.frx":27D4
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
      Left            =   3720
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
      Format          =   134283267
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   105
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
      MICON           =   "frmpenerimaanedit.frx":2AB6
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
      Left            =   8760
      TabIndex        =   11
      Top             =   4680
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
      MICON           =   "frmpenerimaanedit.frx":2DD0
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
      Left            =   7800
      TabIndex        =   10
      Top             =   4680
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
      MICON           =   "frmpenerimaanedit.frx":30EA
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
      Left            =   5880
      TabIndex        =   8
      Top             =   4680
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
      MICON           =   "frmpenerimaanedit.frx":3404
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
      Left            =   6840
      TabIndex        =   9
      Top             =   4680
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
      MICON           =   "frmpenerimaanedit.frx":371E
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
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3A38
      Caption         =   "frmpenerimaanedit.frx":3A58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":3AC4
      Keys            =   "frmpenerimaanedit.frx":3AE2
      Spin            =   "frmpenerimaanedit.frx":3B24
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
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3B4C
      Caption         =   "frmpenerimaanedit.frx":3B6C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":3BD8
      Keys            =   "frmpenerimaanedit.frx":3BF6
      Spin            =   "frmpenerimaanedit.frx":3C38
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
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3C60
      Caption         =   "frmpenerimaanedit.frx":3C80
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":3CEC
      Keys            =   "frmpenerimaanedit.frx":3D0A
      Spin            =   "frmpenerimaanedit.frx":3D4C
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
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3D74
      Caption         =   "frmpenerimaanedit.frx":3D94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":3E00
      Keys            =   "frmpenerimaanedit.frx":3E1E
      Spin            =   "frmpenerimaanedit.frx":3E60
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
      Left            =   4200
      TabIndex        =   26
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
      Format          =   134283267
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil5 
      Height          =   225
      Left            =   6120
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3E88
      Caption         =   "frmpenerimaanedit.frx":3EA8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":3F14
      Keys            =   "frmpenerimaanedit.frx":3F32
      Spin            =   "frmpenerimaanedit.frx":3F74
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
   Begin TDBNumber6Ctl.TDBNumber txtnil6 
      Height          =   225
      Left            =   6120
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":3F9C
      Caption         =   "frmpenerimaanedit.frx":3FBC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":4028
      Keys            =   "frmpenerimaanedit.frx":4046
      Spin            =   "frmpenerimaanedit.frx":4088
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
   Begin TDBNumber6Ctl.TDBNumber txtnil7 
      Height          =   225
      Left            =   5280
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":40B0
      Caption         =   "frmpenerimaanedit.frx":40D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":413C
      Keys            =   "frmpenerimaanedit.frx":415A
      Spin            =   "frmpenerimaanedit.frx":419C
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
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaanedit.frx":41C4
      Caption         =   "frmpenerimaanedit.frx":41E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanedit.frx":4250
      Keys            =   "frmpenerimaanedit.frx":426E
      Spin            =   "frmpenerimaanedit.frx":42B0
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
   Begin akProgress.akProgressBar pro1 
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
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
      TabIndex        =   37
      Top             =   4800
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
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
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2070
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "No Surat Jalan"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label31 
      Caption         =   "Driver / No.Kendaraan"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No P.O."
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal LPB"
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
      TabIndex        =   24
      Top             =   4440
      Width           =   9975
   End
End
Attribute VB_Name = "frmpenerimaanedit"
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
Dim bo1 As Boolean

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
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
        MsgBox "Can not update, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_pohdr where nopo = '" & txtpo & "' and flag = '0'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, Purchase Order already close/cancel.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_beliretur where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not update, data already return.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_beliapp where nobeli='" & txtnobukti & "' and flag2='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not update, data already confirm.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    bo1 = True
    OBJ1.Open dsn
    SQL1 = "select *,len(kode)'lebar' from am_nomax"
    Set RST1 = OBJ1.Execute(SQL1)
    Do While Not RST1.EOF
        If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
        (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
        (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
            bo1 = False
            Exit Do
        End If
        RST1.MoveNext
    Loop
    OBJ1.Close
    
    grid.Row = 1
    pro1.Max = grid.Rows - 2
    pro1.Value = 0
    pro1.Visible = True
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 6) = "0.00" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        If (grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row, 7)) And (Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) <> Val(Format(grid.TextMatrix(grid.Row, 6), "general number"))) Then
            MsgBox "Qty <> QtyUse , on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        If bo1 Then
            OBJ1.Open dsn
            SQL1 = "select qty from am_polin where nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nopo = '" & txtpo & "' and b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            OBJ1.Close
                        
            If Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) > (txtnil1 - txtnil2) Then
                MsgBox "Purchase Order required, Qty max = " & (txtnil1 - txtnil2), vbExclamation, "Information"
                Exit Sub
            End If
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

        SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbeli < '" & tanggalpo & "'"
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
        
        txtnil4 = txtnil6 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8 + Val(Format(grid.TextMatrix(grid.Row, 3), "general number"))
        
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
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.tglbeli = '" & tanggal2 & "'"
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
                MsgBox txtnil4
                MsgBox "Can not update data, quantity item Limited.", vbOKOnly + vbExclamation, "Warning"
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
    SQL = "select count(nobeli)'totalputar' from am_belilin where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    pro1.Value = 0
    If Not RST.EOF Then pro1.Max = RST!totalputar Else pro1.Max = 0
    pro1.Visible = True
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_belilin where nobeli = '" & txtnobukti & "' order by lineitem"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        For i = 1 To grid.Rows - 2
            If RST!kodebarang = grid.TextMatrix(i, 1) Then GoTo balik1
        Next i
        
        OBJ1.Open dsn
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "' and b.tglbpb < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil1 = RST1!qty
        Else
            txtnil1 = 0
        End If

        SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & RST!kodebarang & "' and b.tglbeli < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil2 = RST1!qty
        Else
            txtnil2 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & RST!kodebarang & "' and dateentry < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil3 = RST1!qty
        Else
            txtnil3 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & RST!kodebarang & "' and tglretur < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil5 = RST1!qty
        Else
            txtnil5 = 0
        End If
        
        SQL1 = "select isnull(sum(qtyawal),0)'qty' from am_invloc where kodebarang = '" & RST!kodebarang & "' and tglupdate < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil6 = RST1!qty
        Else
            txtnil6 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut < '" & tanggalpo & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil7 = RST1!qty
        Else
            txtnil7 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut < '" & tanggalpo & "' and (b.type = '02' or b.type = '03')"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil8 = RST1!qty
        Else
            txtnil8 = 0
        End If
        OBJ1.Close
        
        txtnil4 = txtnil6 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
        date2 = date1
        date3 = date1
        OBJ1.Open dsn
        SQL1 = "select isnull(max(b.tglbpb),01/01/1900)'tanggal' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglbeli),01/01/1900)'tanggal' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(dateentry),01/01/1900)'tanggal' from am_usesisa where kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(tglretur),01/01/1900)'tanggal' from am_beliretur where kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and (b.type = '02' or b.type = '03')"
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
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "' and b.tglbpb = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & RST!kodebarang & "' and b.tglbeli = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & RST!kodebarang & "' and dateentry = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil3 = RST1!qty
            Else
                txtnil3 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & RST!kodebarang & "' and tglretur = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil5 = RST1!qty
            Else
                txtnil5 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut = '" & tanggal2 & "' and b.type = '01'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil7 = RST1!qty
            Else
                txtnil7 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut = '" & tanggal2 & "' and (b.type = '02' or b.type = '03')"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil8 = RST1!qty
            Else
                txtnil8 = 0
            End If
            OBJ1.Close
            
            txtnil4 = txtnil4 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
            
            If txtnil4 < 0 Then
                OBJ.Close
                MsgBox "Can not update data, quantity item Limited.", vbOKOnly + vbExclamation, "Warning"
                Exit Sub
            End If
            pro2.Value = pro2.Value + 1
                        
            If date2 = date3 Then Exit Do
            
            date2 = date2 + 1
        Loop
        pro2.Visible = False
balik1:
        pro1.Value = pro1.Value + 1
        RST.MoveNext
    Loop
    OBJ.Close
    pro1.Visible = False
    
    OBJ.Open dsn
    SQL = "delete from am_beliapp where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
        
    SQL = "delete from am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_belilin where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "insert into am_belihdr ("
    SQL = SQL + "nobeli, "
    SQL = SQL + "tglbeli, "
    SQL = SQL + "nopo, "
    SQL = SQL + "nosj, "
    SQL = SQL + "nokend, "
    SQL = SQL + "driver, "
    SQL = SQL + "terima)"
    
    SQL = SQL + " values('" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
    SQL = SQL + "'" & txtpo & "',"
    SQL = SQL + "'" & txtsj & "',"
    SQL = SQL + "'" & txtkend & "',"
    SQL = SQL + "'" & txtdriver & "',"
    SQL = SQL + "'" & str1 & "')"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                        
        SQL = "insert into am_belilin ("
        SQL = SQL + "nobeli, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "kodesatuan, "
        SQL = SQL + "qtyUse, "
        SQL = SQL + "kodesatuanuse, "
        SQL = SQL + "lineitem)"

        SQL = SQL + " values ('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    
    '================================== (untuk update otomatis) ==================================
    
    SQL = "delete FROM am_belirev WHERE nobeli = '" & txtnobukti & "' and flag1='0' and flag2='0'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "SELECT b.nobeli,b.tglbeli,b.nopo,c.kodecur,c.nilaikurs,a.kodebarang,a.qty,d.price,a.kodesatuan,a.lineitem,c.kodesupp,b.driver,c.ket1,c.ket2,c.ket3 FROM am_belilin a left join AM_belihdr b on a.nobeli=b.nobeli left join am_pohdr c on b.nopo=c.nopo left join am_polin d on a.kodebarang=d.kodebarang and d.nopo=b.nopo WHERE b.nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "INSERT INTO AM_belirev"
        SQL1 = SQL1 + " (noBeli"
        SQL1 = SQL1 + ", TglBeli"
        SQL1 = SQL1 + ", nopo"
        SQL1 = SQL1 + ", ref1"
        SQL1 = SQL1 + ", ref2"
        SQL1 = SQL1 + ", kodesupp"
        SQL1 = SQL1 + ", kodecur"
        SQL1 = SQL1 + ", nilaikurs"
        SQL1 = SQL1 + ", Kodebarang"
        SQL1 = SQL1 + ", qty"
        SQL1 = SQL1 + ", Price"
        SQL1 = SQL1 + ", kodesatuan"
        SQL1 = SQL1 + ", keterangan"
        SQL1 = SQL1 + ", keterangan2"
        SQL1 = SQL1 + ", keterangan3"
        SQL1 = SQL1 + ", keterangan4"
        SQL1 = SQL1 + ", ppn"
        SQL1 = SQL1 + ", lineitem"
        SQL1 = SQL1 + ", flag1"   '0 untuk update 1 untuk delete
        SQL1 = SQL1 + ", flag2)"
    
        SQL1 = SQL1 + " VALUES"
        SQL1 = SQL1 + " ('" & txtnobukti & "'"
        SQL1 = SQL1 + ",Convert(dateTime, '" & tanggalpo & "')"
        SQL1 = SQL1 + ", '" & txtpo & "'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
        SQL1 = SQL1 + ", '" & RST!kodecur & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
        SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
        SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
        SQL1 = SQL1 + ", '" & RST!driver & "'"
        SQL1 = SQL1 + ", '" & RST!ket1 & "'"
        SQL1 = SQL1 + ", '" & RST!ket2 & "'"
        SQL1 = SQL1 + ", '" & RST!ket3 & "'"
        SQL1 = SQL1 + ",Convert (Money, '0')"
        SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
        SQL1 = SQL1 + ", '0'"   '0 untuk update 1 untuk delete
        SQL1 = SQL1 + ", '0')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
    
        RST.MoveNext
    Loop
    '================================== (untuk update otomatis) ==================================
    
    'simpan ke table beli apply
    SQL = "SELECT b.nobeli,b.tglbeli,b.nopo,c.kodecur,c.nilaikurs,a.kodebarang,a.qty,d.price,a.kodesatuan,a.lineitem,c.kodesupp,b.driver,c.ket1,c.ket2,c.ket3 FROM am_belilin a left join AM_belihdr b on a.nobeli=b.nobeli left join am_pohdr c on b.nopo=c.nopo left join am_polin d on a.kodebarang=d.kodebarang and d.nopo=b.nopo WHERE b.nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_beliapp where nobeli = '" & RST!NoBeli & "' and kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_beliapp"
            SQL1 = SQL1 + " (noBeli"
            SQL1 = SQL1 + ", TglBeli"
            SQL1 = SQL1 + ", nopo"
            SQL1 = SQL1 + ", ref1"
            SQL1 = SQL1 + ", ref2"
            SQL1 = SQL1 + ", kodesupp"
            SQL1 = SQL1 + ", kodecur"
            SQL1 = SQL1 + ", nilaikurs"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", Price"
            SQL1 = SQL1 + ", kodesatuan"
            SQL1 = SQL1 + ", keterangan"
            SQL1 = SQL1 + ", keterangan2"
            SQL1 = SQL1 + ", keterangan3"
            SQL1 = SQL1 + ", keterangan4"
            SQL1 = SQL1 + ", ppn"
            SQL1 = SQL1 + ", lineitem"
            SQL1 = SQL1 + ", flag1"
            SQL1 = SQL1 + ", flag2)"
        
            SQL1 = SQL1 + " VALUES"
            SQL1 = SQL1 + " ('" & RST!NoBeli & "'"
            SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglbeli) & "/" & Day(RST!tglbeli) & "/" & Year(RST!tglbeli) & "')"
            SQL1 = SQL1 + ", '" & RST!nopo & "'"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
            SQL1 = SQL1 + ", '" & RST!kodecur & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
            SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
            SQL1 = SQL1 + ", '" & RST!driver & "'"
            SQL1 = SQL1 + ", '" & RST!ket1 & "'"
            SQL1 = SQL1 + ", '" & RST!ket2 & "'"
            SQL1 = SQL1 + ", '" & RST!ket3 & "'"
            SQL1 = SQL1 + ",Convert (Money, '0')"
            SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
            SQL1 = SQL1 + ", '1'"
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        RST.MoveNext
        DoEvents
    Loop
    
    OBJ.Close
    'Cetak_Bukti
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    Cetak_Bukti
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    
    txtnobukti.Enabled = True
    cmdsearch1.Enabled = True
    date1.Enabled = True
    txtnobukti = ""
    date1.Value = Date
    txtpo = ""
    txtsj = ""
    txtkend = ""
    txtdriver = ""
    
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
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
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalpo & "' and tanggal2 >= '" & tanggalpo & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_beliapp where nobeli='" & txtnobukti & "' and flag2='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete,Data already confirm.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
        
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_beliretur where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete, data already return.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select count(nobeli)'totalputar' from am_belilin where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    pro1.Value = 0
    If Not RST.EOF Then pro1.Max = RST!totalputar Else pro1.Max = 0
    pro1.Visible = True
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select kodebarang from am_belilin where nobeli = '" & txtnobukti & "' order by lineitem"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "' and b.tglbpb < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil1 = RST1!qty
        Else
            txtnil1 = 0
        End If

        SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & RST!kodebarang & "' and b.tglbeli < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil2 = RST1!qty
        Else
            txtnil2 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & RST!kodebarang & "' and dateentry < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil3 = RST1!qty
        Else
            txtnil3 = 0
        End If
        
        SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & RST!kodebarang & "' and tglretur < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil5 = RST1!qty
        Else
            txtnil5 = 0
        End If
        
        SQL1 = "select isnull(sum(qtyawal),0)'qty' from am_invloc where kodebarang = '" & RST!kodebarang & "' and tglupdate < '" & tanggalpo & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil6 = RST1!qty
        Else
            txtnil6 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut < '" & tanggalpo & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil7 = RST1!qty
        Else
            txtnil7 = 0
        End If
        
        SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut < '" & tanggalpo & "' and (b.type = '02' or b.type = '03')"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnil8 = RST1!qty
        Else
            txtnil8 = 0
        End If
        OBJ1.Close
        
        txtnil4 = txtnil6 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
        date2 = date1
        date3 = date1
        
        OBJ1.Open dsn
        SQL1 = "select isnull(max(b.tglbpb),01/01/1900)'tanggal' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglbeli),01/01/1900)'tanggal' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where a.kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(dateentry),01/01/1900)'tanggal' from am_usesisa where kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(tglretur),01/01/1900)'tanggal' from am_beliretur where kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.type = '01'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If date3 < RST1!tanggal Then date3 = RST1!tanggal
        End If
        SQL1 = "select isnull(max(b.tglmut),01/01/1900)'tanggal' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and (b.type = '02' or b.type = '03')"
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
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_uselin a left join am_usehdr b on a.nobpb=b.nobpb where a.kodebarang = '" & RST!kodebarang & "' and b.tglbpb = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
    
            SQL1 = "select isnull(sum(a.qtyuse),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and a.kodebarang = '" & RST!kodebarang & "' and b.tglbeli = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_usesisa where kodebarang = '" & RST!kodebarang & "' and dateentry = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil3 = RST1!qty
            Else
                txtnil3 = 0
            End If
            
            SQL1 = "select isnull(sum(qty),0)'qty' from am_beliretur where kodebarang = '" & RST!kodebarang & "' and tglretur = '" & tanggal2 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil5 = RST1!qty
            Else
                txtnil5 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut = '" & tanggal2 & "' and b.type = '01'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil7 = RST1!qty
            Else
                txtnil7 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_mutlin a left join am_muthdr b on a.nomut=b.nomut and a.type=b.type where a.kodebarang = '" & RST!kodebarang & "' and b.tglmut = '" & tanggal2 & "' and (b.type = '02' or b.type = '03')"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil8 = RST1!qty
            Else
                txtnil8 = 0
            End If
            OBJ1.Close
                    
            txtnil4 = txtnil4 + txtnil2 - txtnil1 + txtnil3 - txtnil5 + txtnil7 - txtnil8
            
            If txtnil4 < 0 Then
                OBJ.Close
                MsgBox "Can not update data, quantity item Limited.", vbOKOnly + vbExclamation, "Warning"
                Exit Sub
            End If
            pro2.Value = pro2.Value + 1
            
            If date2 = date3 Then Exit Do
            
            date2 = date2 + 1
        Loop
        pro2.Visible = False
        pro1.Value = pro1.Value + 1
        
        RST.MoveNext
    Loop
    OBJ.Close
    pro1.Visible = False
    
    OBJ.Open dsn
    SQL = "select * from am_pohdr where nopo = '" & txtpo & "' and flag = '0'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Purchase Order already close/cancel.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "delete from am_beliapp where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    '================================== (untuk update otomatis) ==================================
    SQL = "SELECT b.nobeli,b.tglbeli,b.nopo,c.kodecur,c.nilaikurs,a.kodebarang,a.qty,d.price,a.kodesatuan,a.lineitem,c.kodesupp,b.driver,c.ket1,c.ket2,c.ket3 FROM am_belilin a left join AM_belihdr b on a.nobeli=b.nobeli left join am_pohdr c on b.nopo=c.nopo left join am_polin d on a.kodebarang=d.kodebarang and d.nopo=b.nopo WHERE b.nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "INSERT INTO AM_belirev"
        SQL1 = SQL1 + " (noBeli"
        SQL1 = SQL1 + ", TglBeli"
        SQL1 = SQL1 + ", nopo"
        SQL1 = SQL1 + ", ref1"
        SQL1 = SQL1 + ", ref2"
        SQL1 = SQL1 + ", kodesupp"
        SQL1 = SQL1 + ", kodecur"
        SQL1 = SQL1 + ", nilaikurs"
        SQL1 = SQL1 + ", Kodebarang"
        SQL1 = SQL1 + ", qty"
        SQL1 = SQL1 + ", Price"
        SQL1 = SQL1 + ", kodesatuan"
        SQL1 = SQL1 + ", keterangan"
        SQL1 = SQL1 + ", keterangan2"
        SQL1 = SQL1 + ", keterangan3"
        SQL1 = SQL1 + ", keterangan4"
        SQL1 = SQL1 + ", ppn"
        SQL1 = SQL1 + ", lineitem"
        SQL1 = SQL1 + ", flag1"   '0 untuk update 1 untuk delete
        SQL1 = SQL1 + ", flag2)"
    
        SQL1 = SQL1 + " VALUES"
        SQL1 = SQL1 + " ('" & txtnobukti & "'"
        SQL1 = SQL1 + ",Convert(dateTime, '" & tanggalpo & "')"
        SQL1 = SQL1 + ", '" & txtpo & "'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
        SQL1 = SQL1 + ", '" & RST!kodecur & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
        SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
        SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
        SQL1 = SQL1 + ", '" & RST!driver & "'"
        SQL1 = SQL1 + ", '" & RST!ket1 & "'"
        SQL1 = SQL1 + ", '" & RST!ket2 & "'"
        SQL1 = SQL1 + ", '" & RST!ket3 & "'"
        SQL1 = SQL1 + ",Convert (Money, '0')"
        SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
        SQL1 = SQL1 + ", '1'"   '0 untuk update 1 untuk delete
        SQL1 = SQL1 + ", '0')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
    
        RST.MoveNext
    Loop
    '================================== (untuk update otomatis) ==================================
    
    SQL = "delete am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete am_belilin where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobeli, convert(char(11),tglbeli)'tglbeli' from am_belihdr where tglbeli >= '" & batas1 & "' and tglbeli <= '" & batas2 & "'"
    Else
        carisql1 = "select nobeli, convert(char(11),tglbeli)'tglbeli' from am_belihdr"
    End If
    namatabel = "Penerimaan Barang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    carinvoice
    hasil = ""
    hasil1 = ""
    txtsj.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='122' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='123' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
        
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Qty (PO)"
    grid.TextMatrix(0, 4) = "K/Sat."
    grid.TextMatrix(0, 5) = "Satuan"
    grid.TextMatrix(0, 6) = "Qty (USE)"
    grid.TextMatrix(0, 7) = "K/Sat."
    grid.TextMatrix(0, 8) = "Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1000
    grid.ColWidth(7) = 800
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 0
    
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
            
            carisql1 = "select a.kodebarang, a.kodesatuan, b.namabarang from am_polin a left join am_apitemmst b on a.kodebarang=b.kodebarang where a.nopo = '" & txtpo & "'"
            namatabel = "Item on PO"
                
            frmsearch.Show vbModal
        Case 3, 6
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
    Case 3, 6
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
            
            bo1 = True
            OBJ1.Open dsn
            SQL1 = "select *,len(kode)'lebar' from am_nomax"
            Set RST1 = OBJ1.Execute(SQL1)
            Do While Not RST1.EOF
                If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
                (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
                (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
                    bo1 = False
                    Exit Do
                End If
                RST1.MoveNext
            Loop
            OBJ1.Close
            
            If bo1 Then
                OBJ1.Open dsn
                SQL1 = "select qty from am_polin where nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil1 = RST1!qty
                Else
                    txtnil1 = 0
                End If
                
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and b.nopo = '" & txtpo & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil2 = RST1!qty
                Else
                    txtnil2 = 0
                End If
                OBJ1.Close
                
                If txtnil1 - txtnil2 = 0 Then
                    MsgBox "Purchase Order required is complete", vbExclamation, "Information"
                    
                    grid.TextMatrix(grid.Row, 1) = ""
                    grid.TextMatrix(grid.Row, 4) = ""
                    grid.TextMatrix(grid.Row, 7) = ""
                
                    Exit Sub
                End If
                
                grid.TextMatrix(grid.Row, 9) = Format(txtnil1 - txtnil2, "###,##0.00")
            Else
                grid.TextMatrix(grid.Row, 9) = "0.00"
            End If
            
            OBJ.Open dsn
            SQL = "select a.namabarang,a.kodesatuanmutasi,b.namasatuan,(c.namasatuan)'satmutasi' from am_apitemmst a left join am_apunit b on a.kodesatuan=b.kodesatuan left join am_apunit c on a.kodesatuanmutasi=c.kodesatuan where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = "0.00"
                grid.TextMatrix(grid.Row, 5) = RST!namasatuan
                grid.TextMatrix(grid.Row, 6) = "0.00"
                grid.TextMatrix(grid.Row, 7) = RST!kodesatuanmutasi
                grid.TextMatrix(grid.Row, 8) = RST!satmutasi
                
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
                grid.TextMatrix(grid.Row, 8) = ""
                grid.TextMatrix(grid.Row, 9) = ""
            End If
            OBJ.Close
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtdriver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtkend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtdriver.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        bo1 = True
        OBJ1.Open dsn
        SQL1 = "select *,len(kode)'lebar' from am_nomax"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
            (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
            (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
                bo1 = False
                Exit Do
            End If
            RST1.MoveNext
        Loop
        OBJ1.Close
        
        If bo1 Then
            If grid.Col = 3 Then
                If txtnilai > Val(Format(grid.TextMatrix(grid.Row, 9), "general number")) Then
                    MsgBox "Purchase Order required, Qty max = " & grid.TextMatrix(grid.Row, 9), vbExclamation, "Information"
                Else
                    grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
                End If
            Else
                grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
            End If
        Else
            grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
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

Private Sub txtpo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsj.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsj_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkend.SetFocus
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
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1000
    grid.ColWidth(7) = 800
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 0
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture
    grid.Col = 1
End Sub

Function tanggalpo()
      tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub

    hapusgrid
    date1 = Date
    txtpo = ""
    txtsj = ""
    txtkend = ""
    txtdriver = ""

    OBJ.Open dsn
    SQL = "select * from am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbeli
        txtpo = RST!nopo
        txtsj = RST!nosj
        txtkend = RST!nokend
        txtdriver = RST!driver
        str1 = RST!terima
        
        bo1 = True
        OBJ1.Open dsn
        SQL1 = "select *,len(kode)'lebar' from am_nomax"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
            (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
            (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
                bo1 = False
                Exit Do
            End If
            RST1.MoveNext
        Loop
        OBJ1.Close

        grid.Row = 1
        SQL = "select kodebarang,qty,isnull(qtyuse,0)'qtyuse',kodesatuan,kodesatuanuse from am_belilin where nobeli = '" & txtnobukti & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 6) = Format(RST!qtyuse, "###,###,##0.00")
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
            grid.Col = 7
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 7) = RST!kodesatuanuse

            OBJ1.Open dsn
            SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 2) = RST1!namabarang

            SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            
            SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 7) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 8) = RST1!namasatuan
            OBJ1.Close
            
            If bo1 Then
                OBJ1.Open dsn
                SQL1 = "select qty from am_polin where nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil1 = RST1!qty
                Else
                    txtnil1 = 0
                End If
                
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nobeli <> '" & txtnobukti & "' and b.nopo = '" & txtpo & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil2 = RST1!qty
                Else
                    txtnil2 = 0
                End If
                OBJ1.Close
            
                grid.TextMatrix(grid.Row, 9) = Format(txtnil1 - txtnil2, "###,##0.00")
            Else
                grid.TextMatrix(grid.Row, 9) = "0.00"
            End If
            
            SetRow grid.Row, True
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
        txtnobukti.Enabled = False
        cmdsearch1.Enabled = False
        date1.Enabled = False
        txtsj.SetFocus
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub Cetak_Bukti()
    With rptbpb
         SQL1 = "Exec am_printbpb '" & txtnobukti & "'"
        .DataControl1.Source = SQL1
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub
