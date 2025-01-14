VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmmutstokbylot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Stock by Lot"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9660
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtcari 
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
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtnobpb 
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
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtsaldo 
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
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   29
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtsisa 
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
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   24
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtout 
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
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   23
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtin 
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
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   22
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtkode 
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
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtnolot 
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
      TabIndex        =   13
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtgudang 
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   8895
      _Version        =   851970
      _ExtentX        =   15690
      _ExtentY        =   661
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Qty Lot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   3615
      Begin VB.OptionButton optqty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty > 0"
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
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optnull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty = 0"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optminus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty < 0"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.RadioButton optlem 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   855
      _Version        =   851970
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " LEM"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
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
   Begin XtremeSuiteControls.RadioButton optkarpet 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   855
      _Version        =   851970
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   " KARPET"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7320
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
      MICON           =   "frmmutstokbylot.frx":0000
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
      Left            =   7680
      TabIndex        =   4
      Top             =   7320
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
      MICON           =   "frmmutstokbylot.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show"
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
      MICON           =   "frmmutstokbylot.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   5640
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
      CustomFormat    =   "ddMMMMyyyy"
      Format          =   142934019
      CurrentDate     =   42052
   End
   Begin Chameleon.chameleonButton cmdgudang 
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   120
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
      MICON           =   "frmmutstokbylot.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtqty 
      Height          =   285
      Left            =   5640
      TabIndex        =   21
      Top             =   6600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   503
      Calculator      =   "frmmutstokbylot.frx":0C68
      Caption         =   "frmmutstokbylot.frx":0C88
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmmutstokbylot.frx":0CF4
      Keys            =   "frmmutstokbylot.frx":0D12
      Spin            =   "frmmutstokbylot.frx":0D54
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   8454143
      BorderStyle     =   1
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
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   6720
      TabIndex        =   32
      Top             =   7320
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
      MICON           =   "frmmutstokbylot.frx":0D7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlot 
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Nomor Lot"
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
      MICON           =   "frmmutstokbylot.frx":1096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   375
      Left            =   8640
      TabIndex        =   42
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Search.."
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
      MICON           =   "frmmutstokbylot.frx":13B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblitem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   5280
      TabIndex        =   43
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label10 
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Apabila nomor lot kosong, harus di isi sesuai dengan nomor lot yang tersedia di gudang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblhppperkg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3600
      TabIndex        =   37
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblkgperpalet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2520
      TabIndex        =   36
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblkg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1440
      TabIndex        =   35
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblkdsatuan 
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
      Left            =   8640
      TabIndex        =   34
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label lblsatuan 
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
      Left            =   8640
      TabIndex        =   33
      Top             =   6600
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   5520
      X2              =   5520
      Y1              =   6960
      Y2              =   5640
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "No Bukti"
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
      TabIndex        =   31
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
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
      Left            =   7560
      TabIndex        =   28
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Out"
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
      Left            =   6600
      TabIndex        =   27
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "In"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Lot"
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
      Left            =   4800
      TabIndex        =   25
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblnmbrg 
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
      Left            =   1440
      TabIndex        =   20
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode/Item"
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
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label lblgudang 
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
      Left            =   2160
      TabIndex        =   17
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
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
      Left            =   4800
      TabIndex        =   16
      Top             =   5685
      Width           =   975
   End
   Begin VB.Label lblrow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "."
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
      Left            =   7440
      TabIndex        =   11
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7095
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmmutstokbylot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim SQLcari As String
Dim str99 As String
Dim jumlah, baris, hint As Integer

Private Sub cmdclear_Click()
    hapusgrid
    optkarpet.Value = False
    optlem.Value = False
    optminus.Value = False
    optnull.Value = False
    optqty.Value = False
    txtgudang = ""
    lblgudang = ""
    lblrow = "0 Lot"
    txtnolot = ""
    txtkode = ""
    lblnmbrg = ""
    lblkdsatuan = ""
    lblsatuan = ""
    txtin = ""
    txtout = ""
    txtsaldo = ""
    txtsisa = ""
    lblkg = ""
    lblhppperkg = ""
    lblkgperpalet = ""
    txtqty = 0
    txtcari = ""
    lblitem = ""
    'ambil kode pindah gudang baru
    strformat = Format(Date1, "yymm")

    OBJ.Open dsn
    SQL = "select top 1 ref from am_stokgudang where ref like 'OPN-' + '" + strformat + "%' order by ref desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!ref, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
    str99 = str99 + 1
    If Len(str99) = 1 Then txtnobpb = "OPN-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobpb = "OPN-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobpb = "OPN-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobpb = "OPN-" & strformat & str99
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdgudang_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdgudang_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    lblgudang = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdlot_Click()
    hasil8 = grid.TextMatrix(grid.Row, 1)
    carisql1 = "Select distinct nolot,tanggal,kg,kgperpalet,hppperkg from am_stokgudang"
    namatabel = "Daftar Lot"
    frmsearch.Show vbModal
End Sub

Private Sub cmdlot_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hasil3 = ""
    hasil4 = ""
    hasil8 = ""
End Sub

Private Sub cmdsearch_Click()
    If txtcari <> "" Then GoTo lewatisql:
    If optminus.Value = True Then
        SQLcari = SQLcari + " and c.stok < 0 and a.gudang= '" & txtgudang & "'"
        SQLcari = SQLcari + " group by a.nolot,a.kodebarang,b.namabarang,c.stok"
    ElseIf optnull.Value = True Then
        SQLcari = SQLcari + " and c.stok = 0 and a.gudang= '" & txtgudang & "'"
        SQLcari = SQLcari + " group by a.nolot,a.kodebarang,b.namabarang,c.stok"
    ElseIf optqty.Value = True Then
        SQLcari = SQLcari + " and c.stok > 0 and a.gudang= '" & txtgudang & "'"
        SQLcari = SQLcari + " group by a.nolot,a.kodebarang,b.namabarang,c.stok"
    End If
lewatisql:
    carisql1 = "Select distinct a.kodebarang,a.namabarang from (" + SQLcari + ") a"
    namatabel = "Daftar Item"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtcari = hasil
    lblitem = hasil1
    hasil = ""
    hasil1 = ""
    opendata
End Sub

Private Sub opendata()
    Screen.MousePointer = vbHourglass
    hapusgrid
    OBJ.Open dsn
    
    SQL = "Select COUNT(nolot)'jml' from (Select a.nolot,a.kodebarang,b.namabarang,sum(a.qin)'in',sum(a.qout)'out',"
    SQL = SQL + "sum(a.qin)-sum(a.qout)'stok' From am_stokgudang a"
    SQL = SQL + " inner join am_itemmst b on a.kodebarang=b.kodebarang"
    SQL = SQL + " Where a.kodebarang='" & txtcari & "' and"
    SQL = SQL + " a.gudang='" & txtgudang & "' Group By a.nolot,a.kodebarang,b.namabarang)a"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    Pg.text = "Please Wait..."
    
    SQL = "Select a.nolot,a.kodebarang,b.namabarang,sum(a.qin)'in',sum(a.qout)'out',"
    SQL = SQL + "sum(a.qin)-sum(a.qout)'stok' From am_stokgudang a"
    SQL = SQL + " inner join am_itemmst b on a.kodebarang=b.kodebarang"
    SQL = SQL + " Where a.kodebarang='" & txtcari & "' and"
    SQL = SQL + " a.gudang='" & txtgudang & "' Group By a.nolot,a.kodebarang,b.namabarang"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = RST!nolot
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang
        grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
        grid.TextMatrix(grid.Row, 3) = Format(RST!In, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 4) = Format(RST!out, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!stok, "##,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        lblrow = Pg.Value & " Lot"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    Pg.text = ""
    Pg.Visible = False
    lblrow = grid.Row - 1 & " Lot"
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdview_Click()
    If txtgudang = "" Then
        MsgBox "Pilih gudang terlebih dahulu", vbExclamation, AppName
        Exit Sub
    End If
    
    If optkarpet.Value = False And optlem.Value = False Then
        MsgBox "Tentukan opsi terlebih dahulu (LEM / KARPET)", vbExclamation, AppName
        Exit Sub
    End If
    
    If optminus.Value = False And optnull.Value = False And optqty.Value = False Then
        MsgBox "Tentukan opsi terlebih dahulu (</=/>)", vbExclamation, AppName
        Exit Sub
    End If
    
    showdata
End Sub

Private Sub showdata()
    Screen.MousePointer = vbHourglass

    hapusgrid
    OBJ.Open dsn
    
    SQL = "Select COUNT(y.nolot)'jml' From (Select a.nolot,a.kodebarang,b.namabarang,c.stok,sum(a.qin)'in',sum(a.qout)'out' from am_stokgudang a"
    SQL = SQL + " left join (Select nolot,kodebarang,sum(qin)-sum(qout)'stok' From am_stokgudang"
    If optlem.Value = True Then
        SQL = SQL + " Where kodebarang like 'L%' group by nolot,kodebarang) c on a.nolot=c.nolot and a.kodebarang=c.kodebarang"
        SQL = SQL + " inner join am_itemmst b on a.kodebarang = b.kodebarang"
        SQL = SQL + " Where a.kodebarang like 'L%'"
    ElseIf optkarpet.Value = True Then
        SQL = SQL + " Where kodebarang like 'K%' group by nolot,kodebarang) c on a.nolot=c.nolot and a.kodebarang=c.kodebarang"
        SQL = SQL + " inner join am_itemmst b on a.kodebarang = b.kodebarang"
        SQL = SQL + " Where a.kodebarang like 'K%'"
    End If
    
    If optminus.Value = True Then
        SQL = SQL + "  and c.stok < 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok)y"
    ElseIf optnull.Value = True Then
        SQL = SQL + "  and c.stok = 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok)y"
    ElseIf optqty.Value = True Then
        SQL = SQL + "  and c.stok > 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok)y"
    End If
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    Pg.text = "Please Wait..."
    
    SQL = "Select a.nolot,a.kodebarang,b.namabarang,c.stok,sum(a.qin)'in',sum(a.qout)'out' from am_stokgudang a"
    SQL = SQL + " left join (Select nolot,kodebarang,sum(qin)-sum(qout)'stok' From am_stokgudang"
    If optlem.Value = True Then
        SQL = SQL + " Where kodebarang like 'L%' group by nolot,kodebarang) c on a.nolot=c.nolot and a.kodebarang=c.kodebarang"
        SQL = SQL + " inner join am_itemmst b on a.kodebarang = b.kodebarang"
        SQL = SQL + " Where a.kodebarang like 'L%'"
    ElseIf optkarpet.Value = True Then
        SQL = SQL + " Where kodebarang like 'K%' group by nolot,kodebarang) c on a.nolot=c.nolot and a.kodebarang=c.kodebarang"
        SQL = SQL + " inner join am_itemmst b on a.kodebarang = b.kodebarang"
        SQL = SQL + " Where a.kodebarang like 'K%'"
    End If
    SQLcari = SQL
    If optminus.Value = True Then
        SQL = SQL + "  and c.stok < 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok order by nolot asc"
    ElseIf optnull.Value = True Then
        SQL = SQL + "  and c.stok = 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok order by nolot asc"
    ElseIf optqty.Value = True Then
        SQL = SQL + "  and c.stok > 0 and a.gudang= '" & txtgudang & "' group by a.nolot,a.kodebarang,b.namabarang,c.stok order by nolot asc"
    End If
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = RST!nolot
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang
        grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
        grid.TextMatrix(grid.Row, 3) = Format(RST!In, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 4) = Format(RST!out, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!stok, "##,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        lblrow = Pg.Value & " Lot"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    Pg.text = ""
    Pg.Visible = False
    lblrow = grid.Row - 1 & " Lot"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    grid.Cols = 6
    grid.TextMatrix(0, 0) = "No Lot"
    grid.TextMatrix(0, 1) = "Kode"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "In"
    grid.TextMatrix(0, 4) = "Out"
    grid.TextMatrix(0, 5) = "Stok"
    
    grid.ColWidth(0) = 1800
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2600
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1000
    hint = 0
    Date1 = Date
    
    'ambil kode pindah gudang baru
    strformat = Format(Date1, "yymm")

    OBJ.Open dsn
    SQL = "select top 1 ref from am_stokgudang where ref like 'OPN-' + '" + strformat + "%' order by ref desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!ref, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
    str99 = str99 + 1
    If Len(str99) = 1 Then txtnobpb = "OPN-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobpb = "OPN-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobpb = "OPN-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobpb = "OPN-" & strformat & str99
    
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    setAlternatingGridBg grid.Row
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As Control
  
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSHFlexGrid Then
          If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        End If
    Next ctl
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    txtnolot = grid.TextMatrix(grid.Row, 0)
    txtkode = grid.TextMatrix(grid.Row, 1)
    lblnmbrg = grid.TextMatrix(grid.Row, 2)
    txtin = grid.TextMatrix(grid.Row, 3)
    txtout = grid.TextMatrix(grid.Row, 4)
    txtsisa = grid.TextMatrix(grid.Row, 5)
    txtqty = 0
    txtsaldo = ""
    
    OBJ.Open dsn
    SQL = "Select * From am_stokgudang Where nolot = '" & grid.TextMatrix(grid.Row, 0) & "'"
    SQL = SQL + " and kodebarang ='" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        MsgBox "Data yang anda pilih tidak lengkap" & vbCrLf & "Silahkan hubungi Administrator", vbCritical, AppName
        OBJ.Close
        Exit Sub
    End If
    
    lblkdsatuan = RST!kdsatuan
    lblsatuan = RST!satuan
    lblkg = RST!kg
    lblkgperpalet = RST!kgperpalet
    If RST!hppperkg = "0.00" Or RST!hppperkg = "0" Then
        SQL = "Select kg,SUM(hppperkg)/COUNT(kodebarang)'perkg' from am_stokgudang"
        SQL = SQL + " Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "' and hppperkg <> '0.00'"
        SQL = SQL + " Group By kg"
        Set RST = OBJ.Execute(SQL)
        
        If RST.EOF Then
            'ambil dari list_hpp_produksi
            lblhppperkg = "0.00"
            lblkg = "0.00"
        Else
            lblhppperkg = Format(RST!perkg, "##,##0.00")
            lblkg = Format(RST!kg, "##,##0.00")
        End If
    Else
        lblhppperkg = Format(RST!hppperkg, "##,##0.00")
    End If
    OBJ.Close
    
    setAlternatingGrid grid.Row
    hint = baris
    
    baris = grid.Row
    If hint = 0 Then Exit Sub
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        If grid.Row = hint Then
            setAlternatingGridBg hint
            Exit Do
        End If
        grid.Row = grid.Row + 1
    Loop
    
    txtqty.SetFocus
End Sub

Private Function setAlternatingGrid(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To 5
        grid.Col = j
        grid.CellBackColor = &H80000010
    Next
End Function

Private Function setAlternatingGridBg(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To 5
        grid.Col = j
        grid.CellBackColor = &H80000005
    Next
End Function

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtsaldo = txtsisa + txtqty
        txtsaldo = Format(txtsaldo, "##,##0.00")
        lblkgperpalet = txtqty * CDbl(lblkg)
        lblkgperpalet = Format(lblkgperpalet, "##,##0.00")
        
        cmdSave.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
    If txtgudang = "" Then
        MsgBox "Kolom gudang tidak boleh kosong", vbCritical, AppName
        Exit Sub
    End If
    
    If txtkode = "" Then
        MsgBox "Kolom kode tidak boleh kosong", vbCritical, AppName
        Exit Sub
    End If
    
    If txtqty = 0 Or IsNull(txtqty) Then
        MsgBox "Kolom qty belum diisi", vbCritical, AppName
        Exit Sub
    End If
    
    'If txtnolot = "" Then   'untuk opname stok data tanpa lot harus kosongkan lot jadi sementara disabled dulu
        'MsgBox "Nomor Lot harus di isi", vbExclamation, AppName
        'Exit Sub
    'End If
    
    strformat = Format(Date1, "yymm")

    OBJ.Open dsn
    SQL = "select top 1 ref from am_stokgudang where ref like 'OPN-' + '" + strformat + "%' order by ref desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!ref, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
    str99 = str99 + 1
    If Len(str99) = 1 Then txtnobpb = "OPN-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobpb = "OPN-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobpb = "OPN-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobpb = "OPN-" & strformat & str99
    
    OBJ.Open dsn
    SQL = "Select * From am_stokgudang Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot
        !palet = "01" & txtnolot
        !tanggal = Format(Date1, "yyyy/MM/dd")
        !ref = txtnobpb
        !keterangan = "Penyesuaian Stok"
        !kodebarang = txtkode
        !NamaBarang = lblnmbrg
        !kg = lblkg
        !kgperpalet = lblkgperpalet
        !hppperkg = lblhppperkg
        If txtqty > 0 Then
            !qin = txtqty
            !qout = "0.00"
            !flag = "0"
        Else
            !qin = "0.00"
            !qout = txtqty * -1
            !flag = "1"
        End If
        !kdsatuan = lblkdsatuan
        !satuan = lblsatuan
        !gudang = txtgudang
        !UserName = nmuser
        
        .Update
    End With
    OBJ.Close
    
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    'cmdclear_Click
    txtnolot = ""
    txtkode = ""
    lblnmbrg = ""
    lblkdsatuan = ""
    lblsatuan = ""
    txtin = ""
    txtout = ""
    txtsaldo = ""
    txtsisa = ""
    lblkg = ""
    lblhppperkg = ""
    lblkgperpalet = ""
    txtqty = 0
    txtcari = ""
    lblitem = ""
    'ambil kode pindah gudang baru
    strformat = Format(Date1, "yymm")

    OBJ.Open dsn
    SQL = "select top 1 ref from am_stokgudang where ref like 'OPN-' + '" + strformat + "%' order by ref desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!ref, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
    str99 = str99 + 1
    If Len(str99) = 1 Then txtnobpb = "OPN-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobpb = "OPN-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobpb = "OPN-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobpb = "OPN-" & strformat & str99
    
    Call showdata
End Sub
