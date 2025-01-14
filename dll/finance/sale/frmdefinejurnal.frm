VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmdefinejurnal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Journal"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
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
   Icon            =   "frmdefinejurnal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5760
      Left            =   75
      TabIndex        =   1
      Top             =   495
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10160
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Jurnal Piutang"
      TabPicture(0)   =   "frmdefinejurnal.frx":2372
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label28"
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(3)=   "Label30"
      Tab(0).Control(4)=   "Label31"
      Tab(0).Control(5)=   "Label32"
      Tab(0).Control(6)=   "Label33"
      Tab(0).Control(7)=   "lblacc01"
      Tab(0).Control(8)=   "lblacc02"
      Tab(0).Control(9)=   "lblacc03"
      Tab(0).Control(10)=   "lblacc04"
      Tab(0).Control(11)=   "lblacc05"
      Tab(0).Control(12)=   "lblacc06"
      Tab(0).Control(13)=   "Label20"
      Tab(0).Control(14)=   "lblacc07"
      Tab(0).Control(15)=   "lblacc00"
      Tab(0).Control(16)=   "Label22"
      Tab(0).Control(17)=   "cmdacc00"
      Tab(0).Control(18)=   "cmdacc07"
      Tab(0).Control(19)=   "cmdacc06"
      Tab(0).Control(20)=   "cmdacc05"
      Tab(0).Control(21)=   "cmdacc04"
      Tab(0).Control(22)=   "cmdacc03"
      Tab(0).Control(23)=   "cmdacc02"
      Tab(0).Control(24)=   "cmdacc01"
      Tab(0).Control(25)=   "cmdadd"
      Tab(0).Control(26)=   "txtacc01"
      Tab(0).Control(27)=   "txtacc02"
      Tab(0).Control(28)=   "txtacc03"
      Tab(0).Control(29)=   "txtacc04"
      Tab(0).Control(30)=   "txtacc05"
      Tab(0).Control(31)=   "txtacc06"
      Tab(0).Control(32)=   "txtacc07"
      Tab(0).Control(33)=   "txtacc00"
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Jurnal Bayar Piutang"
      TabPicture(1)   =   "frmdefinejurnal.frx":238E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(5)=   "lbldesc1"
      Tab(1).Control(6)=   "lbldesc2"
      Tab(1).Control(7)=   "lbldesc3"
      Tab(1).Control(8)=   "lbldesc5"
      Tab(1).Control(9)=   "Label26"
      Tab(1).Control(10)=   "lbldesc6"
      Tab(1).Control(11)=   "Label35"
      Tab(1).Control(12)=   "cmdsearch6"
      Tab(1).Control(13)=   "cmdsearch5"
      Tab(1).Control(14)=   "cmdsearch3"
      Tab(1).Control(15)=   "cmdsearch2"
      Tab(1).Control(16)=   "cmdsearch1"
      Tab(1).Control(17)=   "cmdadd1"
      Tab(1).Control(18)=   "txtacc2"
      Tab(1).Control(19)=   "txtacc1"
      Tab(1).Control(20)=   "txtacc3"
      Tab(1).Control(21)=   "txtacc5"
      Tab(1).Control(22)=   "txtacc6"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Proses Jurnal Piutang"
      TabPicture(2)   =   "frmdefinejurnal.frx":23AA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtnilai3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "date3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdmanual"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdproses1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "date2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "date1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtno1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtkode1"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtnilai1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "pro1"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "grid2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtacc"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Proses Jurnal Bayar Piutang"
      TabPicture(3)   =   "frmdefinejurnal.frx":23C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pro2"
      Tab(3).Control(1)=   "txtkode2"
      Tab(3).Control(2)=   "cmdproses2"
      Tab(3).Control(3)=   "date4"
      Tab(3).Control(4)=   "date5"
      Tab(3).Control(5)=   "grid1"
      Tab(3).Control(6)=   "cmdverify"
      Tab(3).Control(7)=   "txtnilai2"
      Tab(3).Control(8)=   "txtnilai4"
      Tab(3).Control(9)=   "date10"
      Tab(3).Control(10)=   "Label27"
      Tab(3).Control(11)=   "Label9"
      Tab(3).Control(12)=   "Label8"
      Tab(3).Control(13)=   "Label7"
      Tab(3).Control(14)=   "Label6"
      Tab(3).Control(15)=   "Label5"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "Proses Jurnal Ganti Tolak"
      TabPicture(4)   =   "frmdefinejurnal.frx":23E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label10"
      Tab(4).Control(1)=   "Label11"
      Tab(4).Control(2)=   "Label14"
      Tab(4).Control(3)=   "Label1"
      Tab(4).Control(4)=   "cmdproses3"
      Tab(4).Control(5)=   "cmdverifyGT"
      Tab(4).Control(6)=   "date7"
      Tab(4).Control(7)=   "date6"
      Tab(4).Control(8)=   "txtkode3"
      Tab(4).Control(9)=   "pro3"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Proses Jurnal Giro"
      TabPicture(5)   =   "frmdefinejurnal.frx":23FE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label23"
      Tab(5).Control(1)=   "Label24"
      Tab(5).Control(2)=   "Label12"
      Tab(5).Control(3)=   "txtnilai5"
      Tab(5).Control(4)=   "cmdproses4"
      Tab(5).Control(5)=   "date9"
      Tab(5).Control(6)=   "date8"
      Tab(5).Control(7)=   "txtkode4"
      Tab(5).Control(8)=   "txtno2"
      Tab(5).Control(9)=   "pro4"
      Tab(5).ControlCount=   10
      Begin TDBText6Ctl.TDBText txtacc 
         Height          =   255
         Left            =   7920
         TabIndex        =   114
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "frmdefinejurnal.frx":241A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":2486
         Key             =   "frmdefinejurnal.frx":24A4
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   2835
         Left            =   60
         TabIndex        =   112
         Top             =   2355
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
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
      Begin XtremeSuiteControls.ProgressBar pro4 
         Height          =   330
         Left            =   -74970
         TabIndex        =   111
         Top             =   4335
         Width           =   11070
         _Version        =   851970
         _ExtentX        =   19526
         _ExtentY        =   582
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.ProgressBar pro3 
         Height          =   330
         Left            =   -74925
         TabIndex        =   110
         Top             =   4230
         Width           =   11025
         _Version        =   851970
         _ExtentX        =   19447
         _ExtentY        =   582
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ProgressBar pro2 
         Height          =   330
         Left            =   -74970
         TabIndex        =   109
         Top             =   4290
         Width           =   4395
         _Version        =   851970
         _ExtentX        =   7752
         _ExtentY        =   582
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ProgressBar pro1 
         Height          =   330
         Left            =   30
         TabIndex        =   108
         Top             =   5280
         Width           =   11040
         _Version        =   851970
         _ExtentX        =   19473
         _ExtentY        =   582
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.TextBox txtno2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72120
         MaxLength       =   9
         TabIndex        =   50
         Text            =   "YYMM/9999"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtkode4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "MB"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtacc6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtacc5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtacc00 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtacc07 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtacc06 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtacc05 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtacc04 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtacc03 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtacc02 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtacc01 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtacc3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   27
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtacc1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtacc2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtkode3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   44
         Text            =   "TP"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtkode2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "BP"
         Top             =   1440
         Width           =   375
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai1 
         Height          =   255
         Left            =   5280
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":24E0
         Caption         =   "frmdefinejurnal.frx":2500
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":256C
         Keys            =   "frmdefinejurnal.frx":258A
         Spin            =   "frmdefinejurnal.frx":25CC
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.TextBox txtkode1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   32
         Text            =   "JP"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtno1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   33
         Text            =   "YYMM/9999"
         Top             =   1440
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   600
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   960
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdproses1 
         Height          =   375
         Left            =   4920
         TabIndex        =   35
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Piutang"
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
         MICON           =   "frmdefinejurnal.frx":25F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdproses2 
         Height          =   375
         Left            =   -73080
         TabIndex        =   40
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Bayar Piutang"
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
         MICON           =   "frmdefinejurnal.frx":290E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdmanual 
         Height          =   375
         Left            =   1455
         TabIndex        =   34
         Top             =   1905
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verify Account Customer + Transaksi GL"
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
         MICON           =   "frmdefinejurnal.frx":2C28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   4920
         TabIndex        =   60
         Top             =   1080
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdadd 
         Height          =   375
         Left            =   -66120
         TabIndex        =   18
         Top             =   4185
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Simpan Jurnal Piutang"
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
         MICON           =   "frmdefinejurnal.frx":2F42
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai3 
         Height          =   255
         Left            =   5280
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":325C
         Caption         =   "frmdefinejurnal.frx":327C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":32E8
         Keys            =   "frmdefinejurnal.frx":3306
         Spin            =   "frmdefinejurnal.frx":3348
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSComCtl2.DTPicker date4 
         Height          =   285
         Left            =   -73560
         TabIndex        =   36
         Top             =   600
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date5 
         Height          =   285
         Left            =   -73560
         TabIndex        =   37
         Top             =   960
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdadd1 
         Height          =   375
         Left            =   -66360
         TabIndex        =   29
         Top             =   3720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Simpan Jurnal Bayar Piutang"
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
         MICON           =   "frmdefinejurnal.frx":3370
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date6 
         Height          =   285
         Left            =   -73560
         TabIndex        =   42
         Top             =   600
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date7 
         Height          =   285
         Left            =   -73560
         TabIndex        =   43
         Top             =   960
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   3975
         Left            =   -70560
         TabIndex        =   41
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   3
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
         _Band(0).Cols   =   3
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   315
         Left            =   -70560
         TabIndex        =   24
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":368A
         PICN            =   "frmdefinejurnal.frx":36A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   315
         Left            =   -70560
         TabIndex        =   26
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":5528
         PICN            =   "frmdefinejurnal.frx":5544
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch3 
         Height          =   315
         Left            =   -70560
         TabIndex        =   28
         Top             =   2865
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":73C6
         PICN            =   "frmdefinejurnal.frx":73E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdverify 
         Height          =   615
         Left            =   -74280
         TabIndex        =   39
         Top             =   2265
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Verify Account Customer + Different Currency + Verify Transaksi GL"
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
         MICON           =   "frmdefinejurnal.frx":9264
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai2 
         Height          =   255
         Left            =   -71640
         TabIndex        =   79
         Top             =   720
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":957E
         Caption         =   "frmdefinejurnal.frx":959E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":960A
         Keys            =   "frmdefinejurnal.frx":9628
         Spin            =   "frmdefinejurnal.frx":966A
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
         ValueVT         =   1992425477
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Chameleon.chameleonButton cmdacc01 
         Height          =   315
         Left            =   -70440
         TabIndex        =   5
         Top             =   1305
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":9692
         PICN            =   "frmdefinejurnal.frx":96AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc02 
         Height          =   315
         Left            =   -71280
         TabIndex        =   7
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":BE60
         PICN            =   "frmdefinejurnal.frx":BE7C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc03 
         Height          =   315
         Left            =   -70440
         TabIndex        =   9
         Top             =   2145
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":E62E
         PICN            =   "frmdefinejurnal.frx":E64A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc04 
         Height          =   315
         Left            =   -70440
         TabIndex        =   11
         Top             =   2625
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":10DFC
         PICN            =   "frmdefinejurnal.frx":10E18
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc05 
         Height          =   315
         Left            =   -70440
         TabIndex        =   13
         Top             =   2985
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":135CA
         PICN            =   "frmdefinejurnal.frx":135E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc06 
         Height          =   315
         Left            =   -70440
         TabIndex        =   15
         Top             =   3345
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":15D98
         PICN            =   "frmdefinejurnal.frx":15DB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc07 
         Height          =   315
         Left            =   -70440
         TabIndex        =   17
         Top             =   3720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":18566
         PICN            =   "frmdefinejurnal.frx":18582
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdacc00 
         Height          =   315
         Left            =   -70440
         TabIndex        =   3
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":1AD34
         PICN            =   "frmdefinejurnal.frx":1AD50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date8 
         Height          =   285
         Left            =   -73560
         TabIndex        =   47
         Top             =   600
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date9 
         Height          =   285
         Left            =   -73560
         TabIndex        =   48
         Top             =   960
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai4 
         Height          =   255
         Left            =   -71640
         TabIndex        =   99
         Top             =   960
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":1D502
         Caption         =   "frmdefinejurnal.frx":1D522
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":1D58E
         Keys            =   "frmdefinejurnal.frx":1D5AC
         Spin            =   "frmdefinejurnal.frx":1D5EE
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
         ValueVT         =   1992425477
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSComCtl2.DTPicker date10 
         Height          =   285
         Left            =   -72600
         TabIndex        =   100
         Top             =   3600
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
         Format          =   134479875
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdsearch5 
         Height          =   315
         Left            =   -70560
         TabIndex        =   20
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":1D616
         PICN            =   "frmdefinejurnal.frx":1D632
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch6 
         Height          =   315
         Left            =   -70560
         TabIndex        =   22
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MPTR            =   1
         MICON           =   "frmdefinejurnal.frx":1F4B4
         PICN            =   "frmdefinejurnal.frx":1F4D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdverifyGT 
         Height          =   375
         Left            =   -72975
         TabIndex        =   45
         Top             =   2160
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verify Transaksi GL"
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
         MICON           =   "frmdefinejurnal.frx":21352
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdproses3 
         Height          =   375
         Left            =   -72960
         TabIndex        =   46
         Top             =   2625
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Ganti Tolak"
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
         MICON           =   "frmdefinejurnal.frx":2166C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdproses4 
         Height          =   375
         Left            =   -70815
         TabIndex        =   51
         Top             =   1920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Giro Tolak dan Cair"
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
         MICON           =   "frmdefinejurnal.frx":21986
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai5 
         Height          =   255
         Left            =   -71400
         TabIndex        =   107
         Top             =   600
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":21CA0
         Caption         =   "frmdefinejurnal.frx":21CC0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":21D2C
         Keys            =   "frmdefinejurnal.frx":21D4A
         Spin            =   "frmdefinejurnal.frx":21D8C
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label12 
         Caption         =   "Kode Transaksi            (2 character)                         (YY=Year, MM=Month, 9999=Counter)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   106
         Top             =   1470
         Width           =   6615
      End
      Begin VB.Label Label1 
         Caption         =   "No Transaksi di GL sesuai dengan No Bukti Pembayaran"
         Height          =   255
         Left            =   -74760
         TabIndex        =   105
         Top             =   1830
         Width           =   4095
      End
      Begin VB.Label Label35 
         Caption         =   "N.S.F.r .............................."
         Height          =   255
         Left            =   -74400
         TabIndex        =   104
         Top             =   1350
         Width           =   2415
      End
      Begin VB.Label lbldesc6 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69960
         TabIndex        =   103
         Top             =   1350
         Width           =   5895
      End
      Begin VB.Label Label26 
         Caption         =   "P.D.C.r .............................."
         Height          =   255
         Left            =   -74400
         TabIndex        =   102
         Top             =   990
         Width           =   2415
      End
      Begin VB.Label lbldesc5 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69960
         TabIndex        =   101
         Top             =   990
         Width           =   5895
      End
      Begin VB.Label Label24 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   97
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Discount Penjualan (Debet) :  -  Lem ...."
         Height          =   255
         Left            =   -74760
         TabIndex        =   96
         Top             =   990
         Width           =   2895
      End
      Begin VB.Label lblacc00 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   95
         Top             =   990
         Width           =   5775
      End
      Begin VB.Label lblacc07 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   94
         Top             =   3750
         Width           =   5775
      End
      Begin VB.Label Label20 
         Caption         =   "Penjualan WaterProof (Kredit)"
         Height          =   255
         Left            =   -74040
         TabIndex        =   93
         Top             =   3750
         Width           =   2175
      End
      Begin VB.Label lblacc06 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   92
         Top             =   3390
         Width           =   5775
      End
      Begin VB.Label lblacc05 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   91
         Top             =   3030
         Width           =   5775
      End
      Begin VB.Label lblacc04 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   90
         Top             =   2670
         Width           =   5775
      End
      Begin VB.Label lblacc03 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   89
         Top             =   2190
         Width           =   5775
      End
      Begin VB.Label lblacc02 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -70680
         TabIndex        =   88
         Top             =   1710
         Width           =   6615
      End
      Begin VB.Label lblacc01 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69840
         TabIndex        =   87
         Top             =   1350
         Width           =   5775
      End
      Begin VB.Label Label33 
         Caption         =   "Penjualan R.Material (Kredit)"
         Height          =   255
         Left            =   -74040
         TabIndex        =   86
         Top             =   3390
         Width           =   2175
      End
      Begin VB.Label Label32 
         Caption         =   "Penjualan Lem (Kredit) ......."
         Height          =   255
         Left            =   -74040
         TabIndex        =   85
         Top             =   3030
         Width           =   2175
      End
      Begin VB.Label Label31 
         Caption         =   "Penjualan Karet (Kredit) ......"
         Height          =   255
         Left            =   -74040
         TabIndex        =   84
         Top             =   2670
         Width           =   2175
      End
      Begin VB.Label Label30 
         Caption         =   "PPn Keluaran (Kredit) ........."
         Height          =   255
         Left            =   -74040
         TabIndex        =   83
         Top             =   2190
         Width           =   2175
      End
      Begin VB.Label Label29 
         Caption         =   "Bonus Penjualan (Debet) ..."
         Height          =   255
         Left            =   -74760
         TabIndex        =   82
         Top             =   1710
         Width           =   2055
      End
      Begin VB.Label Label28 
         Caption         =   "-  Karet .."
         Height          =   255
         Left            =   -72600
         TabIndex        =   81
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Account Piutang (Debet) ....   Define Account Customer"
         Height          =   255
         Left            =   -74760
         TabIndex        =   80
         Top             =   630
         Width           =   4095
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Penjualan dan Pembayaran yang berbeda Currency."
         Height          =   255
         Left            =   -68280
         TabIndex        =   78
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lbldesc3 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69960
         TabIndex        =   77
         Top             =   2910
         Width           =   5895
      End
      Begin VB.Label lbldesc2 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69960
         TabIndex        =   76
         Top             =   2550
         Width           =   5895
      End
      Begin VB.Label lbldesc1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -69960
         TabIndex        =   75
         Top             =   1710
         Width           =   5895
      End
      Begin VB.Label Label19 
         Caption         =   "Selisih Bayar (Debet/Kredit) ......"
         Height          =   255
         Left            =   -74400
         TabIndex        =   74
         Top             =   2910
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Discount Bayar (Debet) ................."
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   1710
         Width           =   2775
      End
      Begin VB.Label Label17 
         Caption         =   "Kas/Bank/ (Debet) .......................  Define Account Bank/Cash"
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   630
         Width           =   4935
      End
      Begin VB.Label Label16 
         Caption         =   "Selisih Kurs (Debet/Kredit) ........"
         Height          =   255
         Left            =   -74400
         TabIndex        =   71
         Top             =   2550
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "Account Piutang (Kredit) ....  Define Account Customer"
         Height          =   255
         Left            =   -74040
         TabIndex        =   70
         Top             =   2070
         Width           =   6255
      End
      Begin VB.Label Label14 
         Caption         =   "Kode Transaksi            (2 character)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   1470
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "No Transaksi di GL sesuai dengan No Bukti Pembayaran"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   1830
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Kode Transaksi            (2 character)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   1470
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -67680
         TabIndex        =   62
         Top             =   1470
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Kode Transaksi            (2 character)                         (YY=Year, MM=Month, 9999=Counter)"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1470
         Width           =   6615
      End
      Begin VB.Label Label3 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   630
         Width           =   1455
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   10260
      TabIndex        =   53
      Top             =   6330
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
      MICON           =   "frmdefinejurnal.frx":21DB4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtkodecomp 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmdefinejurnal.frx":220CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdefinejurnal.frx":2213A
      Key             =   "frmdefinejurnal.frx":22158
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
      MaxLength       =   4
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   57
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Company"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmdefinejurnal.frx":22194
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpreview 
      Height          =   375
      Left            =   75
      TabIndex        =   52
      Top             =   6360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview (Penjualan dan pembayaran yang berbeda currency)"
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
      MICON           =   "frmdefinejurnal.frx":224AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   4980
      Top             =   6345
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdmanualjurnal 
      Height          =   375
      Left            =   6585
      TabIndex        =   113
      Top             =   6330
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Manual Jurnal"
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
      MICON           =   "frmdefinejurnal.frx":227C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblnamacomp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   58
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmdefinejurnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ0 As New ADODB.Connection
Dim RST0 As New ADODB.Recordset
Dim SQL0 As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim OBJ3 As New ADODB.Connection
Dim RST3 As New ADODB.Recordset
Dim SQL3 As String

Dim SP As New ADODB.Command
Dim vsp(2) As Variant

Dim posrow As String
Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, str15 As String
Dim int2, jml As Integer

Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdacc00_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc00_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc00 = hasil
    lblacc00 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc01.SetFocus
End Sub

Private Sub cmdacc01_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc01_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc01 = hasil
    lblacc01 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc02.SetFocus
End Sub

Private Sub cmdacc02_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc02_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc02 = hasil
    lblacc02 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc03.SetFocus
End Sub

Private Sub cmdacc03_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc03_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc03 = hasil
    lblacc03 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc04.SetFocus
End Sub

Private Sub cmdacc04_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc04_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc04 = hasil
    lblacc04 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc05.SetFocus
End Sub

Private Sub cmdacc05_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc05_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc05 = hasil
    lblacc05 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc06.SetFocus
End Sub

Private Sub cmdacc06_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc06_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc06 = hasil
    lblacc06 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc07.SetFocus
End Sub

Private Sub cmdacc07_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdacc07_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc07 = hasil
    lblacc07 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    cmdadd.SetFocus
End Sub

Private Sub cmdadd_Click()
    If txtkodecomp = "" Or txtacc00 = "" Or txtacc01 = "" Or txtacc02 = "" Or txtacc03 = "" Or txtacc04 = "" Or txtacc05 = "" Or txtacc06 = "" Or txtacc07 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "select ac_cust from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str2 = RST!ac_cust
        
    'piutang
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & str2 & "',"
    SQL = SQL + "'D',"
    SQL = SQL + "'cust',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'1'))"
    Set RST = OBJ.Execute(SQL)
    'discount lem
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc00 & "',"
    SQL = SQL + "'D',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'2'))"
    Set RST = OBJ.Execute(SQL)
    'discount karet
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc01 & "',"
    SQL = SQL + "'D',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'3'))"
    Set RST = OBJ.Execute(SQL)
    'bonus
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc02 & "',"
    SQL = SQL + "'D',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'4'))"
    Set RST = OBJ.Execute(SQL)
    'ppn
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc03 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'5'))"
    Set RST = OBJ.Execute(SQL)
    'jual karet
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc04 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'K',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'6'))"
    Set RST = OBJ.Execute(SQL)
    'jual lem
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc05 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'L',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'7'))"
    Set RST = OBJ.Execute(SQL)
    'jual material
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc06 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'R',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'8'))"
    Set RST = OBJ.Execute(SQL)
    'jual water
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'piutang',"
    SQL = SQL + "'" & txtacc07 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'W',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'9'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdadd1_Click()
    If txtkodecomp = "" Or txtacc1 = "" Or txtacc2 = "" Or txtacc3 = "" Or txtacc5 = "" Or txtacc6 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang'"
    Set RST = OBJ.Execute(SQL)
    
    'kas/bank
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'',"
    SQL = SQL + "'D',"
    SQL = SQL + "'bank',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'1'))"
    Set RST = OBJ.Execute(SQL)
    'disc bayar
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & txtacc1 & "',"
    SQL = SQL + "'D',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'2'))"
    Set RST = OBJ.Execute(SQL)
    'piutang
    SQL = "select ac_cust from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str2 = RST!ac_cust
    
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & str2 & "',"
    SQL = SQL + "'K',"
    SQL = SQL + "'cust',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'3'))"
    Set RST = OBJ.Execute(SQL)
    'selisih kurs
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & txtacc2 & "',"
    SQL = SQL + "'A',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'4'))"
    Set RST = OBJ.Execute(SQL)
    'selisih bayar
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & txtacc3 & "',"
    SQL = SQL + "'A',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'5'))"
    Set RST = OBJ.Execute(SQL)
    'pdc
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & txtacc5 & "',"
    SQL = SQL + "'A',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'7'))"
    Set RST = OBJ.Execute(SQL)
    'nsf
    SQL = "insert into am_autoo ("
    SQL = SQL + "kodecomp, "
    SQL = SQL + "jurnal_, "
    SQL = SQL + "noacc, "
    SQL = SQL + "dk, "
    SQL = SQL + "kdkurs, "
    SQL = SQL + "nanti, "
    SQL = SQL + "line)"

    SQL = SQL + " values('" & txtkodecomp & "',"
    SQL = SQL + "'bpiutang',"
    SQL = SQL + "'" & txtacc6 & "',"
    SQL = SQL + "'A',"
    SQL = SQL + "'one',"
    SQL = SQL + "'',"
    SQL = SQL + "convert(numeric,'8'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdmanual_Click()
    If txtkodecomp = "" Then Exit Sub
    Dim jml As Integer
    
    If date1 > date2 Then
        MsgBox "Invalid date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date1.Month <> date2.Month Or date1.Year <> date2.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Verify Account berlangsung." & vbCrLf & _
    "Lanjutkan Proses Verify Account ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
        
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_postingpenjualan"
    SP.CommandTimeout = 1000
    vsp(0) = Format(date1, "yyyyMMdd")
    vsp(1) = Format(date2, "yyyyMMdd")
    vsp(2) = txtkodecomp
    SP.Execute , vsp
    Set SP = Nothing
    
    OBJ.Open dsn
    SQL = "select distinct kodecust from am_manualpostjual where noacc = ''"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "There is undefine account customer, please run Define account Customer.", vbInformation, "Information"
        Exit Sub
    End If
    
    
    SQL = "select * from am_manualpostjual"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "There is no transaction to proses.", vbInformation, "Information"
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Verify account complete, continue verify Transaksi GL.", vbInformation, "Information"
    
    pro1.Visible = True
    pro1.Value = 0
    
    OBJ.Open dsn
    SQL = "select distinct nobkt from am_manualpostjual where noacc <> ''"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
    Loop
    pro1.Max = jml
    RST.MoveFirst
    End If
    
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode1 & "' and left(desctrx,8)='" & RST!nobkt & "' and flag<>'B'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
        MsgBox RST1!notrx
            OBJ1.Close
            OBJ.Close
            
            pro1.Visible = False
            pro1.Value = 0
            pro1.Max = 0

            MsgBox "Transaksi penjualan sudah ada di GL dan sudah terposting." & vbCrLf & _
            "proses di batalkan, unposting dahulu transaksi GL.", vbInformation, "Proses Batal"
            
            Exit Sub
        End If
        OBJ1.Close
        
        pro1.Value = pro1.Value + 1
        If pro1.Value = jml Then pro1.Value = 0
        RST.MoveNext
        DoEvents
    Loop
    OBJ.Close
    
    pro1.Visible = False
    pro1.Value = 0
    pro1.Max = 0
    
    MsgBox "Verify Transaksi GL complete.", vbInformation, "Information"
        
    cmdproses1.Enabled = True
    cmdmanual.Enabled = False
    cmdproses1.SetFocus
End Sub

Private Sub cmdmanualjurnal_Click()
    Dim jml As Integer
    grid2.Clear
    grid2.Rows = 2
    grid2.Row = 1
    setgrid2
    
    OBJ.Open dsn
    SQL = "Select count(nobkt)'jml' from am_cashhdr"
    SQL = SQL + " Where tglbkt >= '" & tanggal1 & "' and tglbkt <= '" & tanggal2 & "'"
    SQL = SQL + " and kodebayar in('CN','DN')"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        pro1.Visible = True
        pro1.Max = RST!jml
    Else
        pro1.Visible = False
        pro1.Max = 0
    End If
    pro1.Value = 0
    
    SQL = "Select nobkt,tglbkt,kodebayar,keterangan,amount from am_cashhdr"
    SQL = SQL + " Where tglbkt >= '" & tanggal1 & "' and tglbkt <= '" & tanggal2 & "'"
    SQL = SQL + " and kodebayar in('CN','DN')"
    SQL = SQL + " Order By tglbkt,nobkt"
    Set RST = OBJ.Execute(SQL)
    
    Do While Not RST.EOF
        grid2.TextMatrix(grid2.Row, 0) = RST!nobkt
        grid2.TextMatrix(grid2.Row, 1) = Format(RST!tglbkt, "dd MMM yyyy")
        grid2.TextMatrix(grid2.Row, 2) = RST!kodebayar
        grid2.TextMatrix(grid2.Row, 3) = RST!keterangan
        grid2.TextMatrix(grid2.Row, 4) = Format(RST!amount, "#,##0.00")
        grid2.TextMatrix(grid2.Row, 5) = ""
        
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        pro1.Value = pro1.Value + 1
        RST.MoveNext
    Loop
    pro1.Visible = False
    pro1.Value = 0
    OBJ.Close
End Sub

Private Sub cmdpreview_Click()
    If MsgBox("Continue Preview ?", vbQuestion + vbYesNo, "Preview") = vbNo Then Exit Sub
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_listbeda)"
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\listbeda.rpt"
    Crystal.ParameterFields(0) = "@tanggal1;" & Format(Date, "yyyy0101") & ";true"
    Crystal.ParameterFields(1) = "@tanggal2;" & Format(Date, "yyyy1231") & ";true"
    Crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdproses1_Click()
    If date1 > date2 Then
        MsgBox "Invalid date range, posting aborted.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If date1.Month <> date2.Month Or date1.Year <> date2.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkode1 = "" Or txtkodecomp = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Piutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Piutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode1 & "' and notrx like '" & Format(date1, "YYMM/") & "%' order by notrx desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str7 = Right(RST!notrx, 3) Else str7 = 0
    OBJ.Close
    
    str7 = str7 + 1
    If Len(str7) = 1 Then str8 = "000" & str7
    If Len(str7) = 2 Then str8 = "00" & str7
    If Len(str7) = 3 Then str8 = "0" & str7
    If Len(str7) = 4 Then str8 = str7
    
    OBJ.Open dsn
    SQL = "select count(nobkt)'hitnojual' from am_manualpostjual where noacc <> ''"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        pro1.Visible = True
        pro1.Max = RST!hitnojual
    Else
        pro1.Visible = False
        pro1.Max = 0
    End If
    pro1.Value = 0
    
    SQL = "select * from am_manualpostjual where noacc <> '' order by nobkt"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        date3 = RST!tglbkt
        txtnilai3 = RST!nilaikurs
        str3 = RST!noacc
        str9 = RST!nobkt
        str11 = RST!kodecust
        int2 = 1
        
        'cek yang sama atau pengulangan overwrite atau skip
        OBJ1.Open dsn
        SQL1 = "select kdtrx,notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx = '" & txtkode1 & "' and left(desctrx,8)='" & str9 & "' and identry='auto'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            str10 = RST1!notrx
            OBJ1.Close
            
            If MsgBox("Transaksi PENJUALAN sudah ada di GL." & vbCrLf & _
            "klik YES untuk Overwrite atau klik NO untuk Skip.", vbQuestion + vbYesNo, "Overwrite / Skip") = vbYes Then
                OBJ1.Open dsn
                SQL1 = "delete from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode1 & "' and notrx='" & str10 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            Else
                Do While Not RST.EOF
                    If RST!nobkt <> str9 Then GoTo yangskip1
                    
                    pro1.Value = pro1.Value + 1
                    RST.MoveNext
                Loop
                If RST.EOF Then GoTo yangskip2
            End If
        Else
            OBJ1.Close
        End If
        
        'piutang
        OBJ1.Open dsn
        SQL1 = "select sum(nilaijual - nilaibn - nilaidisc + nilaippn)'nilpiutang' from am_manualpostjual where nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai1 = RST1!nilpiutang
        OBJ1.Close
        
        OBJ1.Open dsn
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "lineitem)"
        
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtkodecomp & "',"  'kode company
        SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "')," 'tgltrx
        SQL1 = SQL1 + "'" & txtkode1 & "'," 'kdtrx
        SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str3 & "',"
        
        OBJ2.Open dsn
        SQL2 = "select kodecust,namacust from am_customer where kodecust='" & str11 & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            SQL1 = SQL1 + "'" & str9 & " " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
        Else
            SQL1 = SQL1 + "'" & str9 & "',"
        End If
        OBJ2.Close
        
        SQL1 = SQL1 + "'D',"
        SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
        SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        'jurnal discount
        OBJ1.Open dsn
        SQL1 = "select nilaidisc,substring(kodebarang,1,1)'kode' from am_manualpostjual where nobkt = '" & RST!nobkt & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            txtnilai1 = RST1!nilaidisc
                        
            If txtnilai1 > 0 Then
                int2 = int2 + 1
                
                OBJ3.Open dsn
                If RST1!kode = "K" Then
                    SQL3 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=3"
                Else
                    SQL3 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=2"
                End If
                Set RST3 = OBJ3.Execute(SQL3)
                If Not RST3.EOF Then str3 = RST3!noacc
                OBJ3.Close
            
                OBJ3.Open dsn
                SQL3 = "insert into gl_transaksi "
                SQL3 = SQL3 + "(kdcomp, "
                SQL3 = SQL3 + "tgltrx, "
                SQL3 = SQL3 + "kdtrx, "
                SQL3 = SQL3 + "notrx, "
                SQL3 = SQL3 + "kurs, "
                SQL3 = SQL3 + "noactrx, "
                SQL3 = SQL3 + "desctrx, "
                SQL3 = SQL3 + "dbkrtrx, "
                SQL3 = SQL3 + "amounttrx, "
                SQL3 = SQL3 + "nilaitrx, "
                SQL3 = SQL3 + "currtrx, "
                SQL3 = SQL3 + "flag, "
                SQL3 = SQL3 + "flagprint, "
                SQL3 = SQL3 + "flagadjustment, "
                SQL3 = SQL3 + "cekbg, "
                SQL3 = SQL3 + "identry, "
                SQL3 = SQL3 + "idupdate, "
                SQL3 = SQL3 + "dateentry, "
                SQL3 = SQL3 + "dateupdate, "
                SQL3 = SQL3 + "lineitem)"
                
                SQL3 = SQL3 + " values"
                SQL3 = SQL3 + "('" & txtkodecomp & "',"
                SQL3 = SQL3 + "convert(datetime,'" & tanggal3 & "'),"
                SQL3 = SQL3 + "'" & txtkode1 & "',"
                SQL3 = SQL3 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
                SQL3 = SQL3 + "convert(money,'1'),"
                SQL3 = SQL3 + "'" & str3 & "',"
            
                OBJ2.Open dsn
                SQL2 = "select kodecust,namacust from am_customer where kodecust='" & str11 & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    SQL3 = SQL3 + "'" & str9 & " " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
                Else
                    SQL3 = SQL3 + "'" & str9 & "',"
                End If
                OBJ2.Close
            
                SQL3 = SQL3 + "'D',"
                SQL3 = SQL3 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
                SQL3 = SQL3 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
                SQL3 = SQL3 + "'" & str1 & "',"
                SQL3 = SQL3 + "'B',"
                SQL3 = SQL3 + "'J',"
                SQL3 = SQL3 + "'0',"
                SQL3 = SQL3 + "'',"
                SQL3 = SQL3 + "'auto',"
                SQL3 = SQL3 + "'',"
                SQL3 = SQL3 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL3 = SQL3 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL3 = SQL3 + "convert(numeric,'" & int2 & "'))"
                Set RST3 = OBJ3.Execute(SQL3)
                OBJ3.Close
            End If
            
            RST1.MoveNext
        Loop
        OBJ1.Close
        
        'jurnal bonus
        OBJ1.Open dsn
        SQL1 = "select sum(nilaibn)'nilbn' from am_manualpostjual where nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai1 = RST1!nilbn Else txtnilai1 = 0
        
        SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=4"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str3 = RST1!noacc
        OBJ1.Close
        
        If txtnilai1 > 0 Then
            int2 = int2 + 1
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "'),"
            SQL1 = SQL1 + "'" & txtkode1 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            OBJ2.Open dsn
            SQL2 = "select kodecust,namacust from am_customer where kodecust='" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then
                SQL1 = SQL1 + "'" & str9 & " " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
            Else
                SQL1 = SQL1 + "'" & str9 & "',"
            End If
            OBJ2.Close
            
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        
        'jurnal ppn keluaran
        OBJ1.Open dsn
        SQL1 = "select sum(nilaippn)'nilppn' from am_manualpostjual where nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai1 = RST1!nilppn Else txtnilai1 = 0
        
        SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=5"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str3 = RST1!noacc
        OBJ1.Close
        
        If txtnilai1 > 0 Then
            int2 = int2 + 1
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "'),"
            SQL1 = SQL1 + "'" & txtkode1 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            OBJ2.Open dsn
            SQL2 = "select kodecust,namacust from am_customer where kodecust='" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then
                SQL1 = SQL1 + "'" & str9 & " " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
            Else
                SQL1 = SQL1 + "'" & str9 & "',"
            End If
            OBJ2.Close
            
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        
        'jurnal penjualan
        Do While Not RST.EOF
            If RST!nobkt <> str9 Then Exit Do
            
            str3 = Left(RST!kodebarang, 1)
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and kdkurs='" & str3 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc
            OBJ1.Close
        
            int2 = int2 + 1
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "'),"
            SQL1 = SQL1 + "'" & txtkode1 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            OBJ2.Open dsn
            SQL2 = "select kodecust,namacust from am_customer where kodecust='" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then
                SQL1 = SQL1 + "'" & str9 & " " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
            Else
                SQL1 = SQL1 + "'" & str9 & "',"
            End If
            OBJ2.Close
            
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST!nilaijual * txtnilai3 & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!nilaijual * txtnilai3 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        
            pro1.Value = pro1.Value + 1
            RST.MoveNext
        Loop
yangskip1:
        OBJ1.Open dsn
        SQL1 = "update am_invhdr set posted = '1' where nobkt = '" & str9 & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
                
        OBJ1.Open dsn
        SQL1 = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode1 & "' and notrx like '" & Format(date1, "YYMM/") & "%' order by notrx desc"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str7 = Right(RST1!notrx, 3) Else str7 = 0
        OBJ1.Close
        
        str7 = str7 + 1
        If Len(str7) = 1 Then str8 = "000" & str7
        If Len(str7) = 2 Then str8 = "00" & str7
        If Len(str7) = 3 Then str8 = "0" & str7
        If Len(str7) = 4 Then str8 = str7
        DoEvents
    Loop
yangskip2:
    OBJ.Close
    
    MsgBox "Proses Complete.", vbInformation, "Information"
    pro1.Visible = False
    pro1.Value = 0
    Unload Me
End Sub

Private Sub cmdproses2_Click()
    If date4 > date5 Then
        MsgBox "Invalid date range, posting aborted.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If date4.Month <> date5.Month Or date4.Year <> date5.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkode2 = "" Or txtkodecomp = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Bayar Piutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Bayar Piutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT count(NoBkt)'hitnobkt' FROM AM_cashhdr WHERE kodebayar='PM' and"
    SQL = SQL + " tglbkt>='" & tanggal4 & "' and tglbkt<='" & tanggal5 & "' and posted='0' and idupdate='0'"
    Set RST = OBJ.Execute(SQL)
    
    
    If Not RST.EOF Then pro2.Max = RST!hitnobkt Else pro2.Max = 0
    If pro2.Max = 0 Then pro2.Visible = False Else pro2.Visible = True
    pro2.Value = 0
    OBJ.Close
    
    int2 = 1
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_cashhdr"
    SQL = SQL + " WHERE kodebayar='PM' and tglbkt>='" & tanggal4 & "' and tglbkt<='" & tanggal5 & "'"
    SQL = SQL + " and posted='0' and idupdate='0' order by tglbkt,nobkt"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        'cek yang sama atau pengulangan overwrite atau skip
        OBJ1.Open dsn
        SQL1 = "select kdtrx,notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx = '" & txtkode2 & "' and notrx='" & RST!nobkt & "' and identry='auto'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            str10 = RST1!notrx
            OBJ1.Close
            
            If MsgBox("Transaksi PEMBAYARAN sudah ada di GL." & vbCrLf & _
            "klik YES untuk Overwrite atau klik NO untuk Skip.", vbQuestion + vbYesNo, "Overwrite / Skip") = vbYes Then
                OBJ1.Open dsn
                SQL1 = "delete from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode2 & "' and notrx='" & str10 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            Else
                GoTo yangskip3
            End If
        Else
            OBJ1.Close
        End If
        
        'kas/bank
        OBJ0.Open dsn
        SQL0 = "select * from am_cashsub where nobkt = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        Do While Not RST0.EOF
            If RST0!Typebayar = "TN" Then
                'tunai
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                
                OBJ2.Open dsn
                SQL2 = "select noacc from am_autoaccbank where kodebank = '" & RST!kodecur & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then str15 = RST2!noacc Else str15 = ""
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Bayar Piutang (Tunai)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            ElseIf RST0!Typebayar = "TF" Then
'MsgBox RST!nobkt & " - " & RST0!acbank
                'transfer
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                
                OBJ2.Open dsn
                SQL2 = "select kode from am_bank where acc = '" & RST0!acbank & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    str15 = RST2!kode
                
                    SQL2 = "select noacc from am_autoaccbank where kodebank = '" & str15 & "'"
                    Set RST2 = OBJ2.Execute(SQL2)
                    If Not RST2.EOF Then str15 = RST2!noacc Else str15 = ""
                Else
                    str15 = ""
                End If
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Bayar Piutang (Transfer)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            ElseIf RST0!Typebayar = "G" Then
                'giro
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=7"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str15 = RST1!noacc Else str15 = ""
                OBJ1.Close
                
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Bayar Piutang (Giro)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            End If
            int2 = int2 + 1
            RST0.MoveNext
        Loop
        OBJ0.Close
        
        'discount bayar
        OBJ1.Open dsn
        SQL1 = "select sum(potongan)'discpiutang' from am_cashlin where kodebayar='PM' and nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai2 = RST1!discpiutang Else txtnilai2 = 0
        OBJ1.Close
        
        If txtnilai2 > 0 Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=2"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"  'potongan
            SQL1 = SQL1 + "'Bayar Piutang (Discount Bayar)',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            int2 = int2 + 1
        End If
        'selisihbayar
        OBJ1.Open dsn
        SQL1 = "select sum(selisih)'selpiutang' from am_cashlin where kodebayar='PM' and nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai2 = RST1!selpiutang Else txtnilai2 = 0
        OBJ1.Close
        
        If txtnilai2 > 0 Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=5"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"  'selisih bayar
            SQL1 = SQL1 + "'Bayar Piutang (Selisih Bayar +)',"
            SQL1 = SQL1 + "'K'," 'debet atau kredit
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            int2 = int2 + 1
        ElseIf txtnilai2 < 0 Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=5"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"  'selisih bayar
            SQL1 = SQL1 + "'Bayar Piutang (Selisih Bayar -)',"
            SQL1 = SQL1 + "'D'," 'debet atau kredit
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs * -1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * RST!nilaikurs * -1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            int2 = int2 + 1
        End If
        'selisihkurs
        OBJ1.Open dsn
        SQL1 = "select sum(nilaikurs)'selkpiutang' from am_cashlin where kodebayar='PM' and nobkt = '" & RST!nobkt & "' group by nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai2 = RST1!selkpiutang Else txtnilai2 = 0
        OBJ1.Close
        
        If txtnilai2 > 0 Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=4"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"  'selisih kurs
            SQL1 = SQL1 + "'Bayar Piutang (Selisih Kurs +)',"
            SQL1 = SQL1 + "'D'," 'debet atau kredit
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            int2 = int2 + 1
        ElseIf txtnilai2 < 0 Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=4"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & txtkodecomp & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"  'selisih kurs
            SQL1 = SQL1 + "'Bayar Piutang (Selisih Kurs -)',"
            SQL1 = SQL1 + "'K'," 'debet atau kredit
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * -1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai2 * -1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            int2 = int2 + 1
        End If
        
        'piutang
        txtnilai4 = txtnilai2 'nilai dari selisih kurs
        
        OBJ1.Open dsn
        SQL1 = "select Nobkt, sum((Amount + potongan + PPN + selisih)*nilaikurs) as Total from AM_Aropnfil WHERE nobkt = '" & RST!nobkt & "' and kodecust = '" & RST!kodecust & "' and kodecur = '" & RST!kodecur & "' group by Nobkt"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtnilai2 = ((RST1!total * -1) + txtnilai4) Else txtnilai2 = 0
        
        SQL1 = "select noacc from am_autoaccust where kodecust = '" & RST!kodecust & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
        OBJ1.Close
        
        OBJ1.Open dsn
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "lineitem)"
        
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtkodecomp & "',"
        SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
        SQL1 = SQL1 + "'" & txtkode2 & "',"
        SQL1 = SQL1 + "'" & RST!nobkt & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str3 & "'," 'customer
        
        OBJ2.Open dsn
        SQL2 = "select kodecust,namacust from am_customer where kodecust='" & RST!kodecust & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            SQL1 = SQL1 + "'Bayar Piutang - " & Mid(RST2!namacust + " (" + RST2!kodecust + ")", 1, 40) & "',"
        Else
            SQL1 = SQL1 + "'Bayar Piutang - Customer',"
        End If
        OBJ2.Close
        
        SQL1 = SQL1 + "'K',"
        SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
        SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
yangskip3:
        'selesai
        OBJ1.Open dsn
        SQL1 = "select * from gl_transaksi where kdtrx='" & txtkode2 & "' and notrx='" & RST!nobkt & "' and noactrx=''"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            SQL1 = "delete from gl_transaksi where kdtrx='" & txtkode2 & "' and notrx='" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
        Else
            OBJ1.Close
        
            OBJ1.Open dsn
            SQL1 = "update am_cashhdr set idupdate = '1' where nobkt = '" & RST!nobkt & "' and kodebayar='PM'"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        pro2.Value = pro2.Value + 1
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Proses Complete.", vbInformation, "Information"
    pro2.Visible = False
    pro2.Value = 0
    Unload Me
End Sub

Private Sub cmdproses3_Click()
    If date6 > date7 Then
        MsgBox "Invalid date range, posting aborted.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If date6.Month <> date7.Month Or date6.Year <> date7.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkode3 = "" Or txtkodecomp = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Ganti Tolak berlangsung." & vbCrLf & _
    "Lanjutkan Proses Ganti Tolak Piutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT count(NoBkt)'hitnobkt' FROM AM_cashhdr WHERE kodebayar='GT' and"
    SQL = SQL + " tglbkt>='" & tanggal6 & "' and tglbkt<='" & tanggal7 & "' and posted='9' and idupdate='0'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then pro3.Max = RST!hitnobkt Else pro3.Max = 0
    If pro3.Max = 0 Then pro3.Visible = False Else pro3.Visible = True
    pro3.Value = 0
    OBJ.Close
    
    int2 = 1
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_cashhdr"
    SQL = SQL + " WHERE kodebayar='GT' and tglbkt>='" & tanggal6 & "' and tglbkt<='" & tanggal7 & "'"
    SQL = SQL + " and posted='9' and idupdate='0' order by tglbkt,nobkt"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        'cek yang sama atau pengulangan overwrite atau skip
        OBJ1.Open dsn
        SQL1 = "select kdtrx,notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode3 & "' and notrx='" & RST!nobkt & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            str10 = RST1!notrx
            OBJ1.Close
            
            If MsgBox("Transaksi GANTI TOLAK sudah ada di GL." & vbCrLf & _
            "klik YES untuk Overwrite atau klik NO untuk Skip.", vbQuestion + vbYesNo, "Overwrite / Skip") = vbYes Then
                OBJ1.Open dsn
                SQL1 = "delete from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode3 & "' and notrx='" & str10 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            Else
                GoTo yangskip4
            End If
        Else
            OBJ1.Close
        End If
        
        txtnilai2 = 0
        'kas/bank
        OBJ0.Open dsn
        SQL0 = "select * from am_cashsub where nobkt = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        Do While Not RST0.EOF
            If RST0!Typebayar = "TN" Then
                'tunai
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode3 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                
                OBJ2.Open dsn
                SQL2 = "select noacc from am_autoaccbank where kodebank = '" & RST!kodecur & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then str15 = RST2!noacc Else str15 = ""
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Ganti Tolak (Tunai)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                
                txtnilai2 = txtnilai2 + (RST0!jumlah * RST!nilaikurs)
                
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            ElseIf RST0!Typebayar = "TF" Then
                'transfer
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode3 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                
                OBJ2.Open dsn
                SQL2 = "select kode from am_bank where acc = '" & RST0!acbank & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    str15 = RST2!kode
                
                    SQL2 = "select noacc from am_autoaccbank where kodebank = '" & str15 & "'"
                    Set RST2 = OBJ2.Execute(SQL2)
                    If Not RST2.EOF Then str15 = RST2!noacc Else str15 = ""
                Else
                    str15 = ""
                End If
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Ganti Tolak (Transfer)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                
                txtnilai2 = txtnilai2 + (RST0!jumlah * RST!nilaikurs)
                
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            ElseIf RST0!Typebayar = "G" Then
                'giro
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=7"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str15 = RST1!noacc Else str15 = ""
                OBJ1.Close
                
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
                SQL1 = SQL1 + "'" & txtkode3 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str15 & "',"
                SQL1 = SQL1 + "'Ganti Tolak (Giro)',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * RST!nilaikurs & "'),"
                
                txtnilai2 = txtnilai2 + (RST0!jumlah * RST!nilaikurs)
                
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
            End If
            int2 = int2 + 1
            RST0.MoveNext
        Loop
        OBJ0.Close
        'NSF
        
        OBJ1.Open dsn
        SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=8"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
        OBJ1.Close
        
        OBJ1.Open dsn
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "lineitem)"
        
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtkodecomp & "',"
        SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
        SQL1 = SQL1 + "'" & txtkode3 & "',"
        SQL1 = SQL1 + "'" & RST!nobkt & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str3 & "'," 'nsf
        SQL1 = SQL1 + "'Ganti Tolak (NSF)',"
        SQL1 = SQL1 + "'K',"
        SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
        SQL1 = SQL1 + "convert(money,'" & txtnilai2 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'" & int2 & "'))"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
yangskip4:
        'selesai
        OBJ1.Open dsn
        SQL1 = "select * from gl_transaksi where kdtrx='" & txtkode3 & "' and notrx='" & RST!nobkt & "' and noactrx=''"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            SQL1 = "delete from gl_transaksi where kdtrx='" & txtkode3 & "' and notrx='" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
        Else
            OBJ1.Close
        
            OBJ1.Open dsn
            SQL1 = "update am_cashhdr set idupdate = '1' where nobkt = '" & RST!nobkt & "' and kodebayar='GT'"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        pro3.Value = pro3.Value + 1
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Proses Complete.", vbInformation, "Information"
    pro3.Visible = False
    pro3.Value = 0
    Unload Me
End Sub

Private Sub cmdproses4_Click()
    If date8 > date9 Then
        MsgBox "Invalid date range, posting aborted.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If date8.Month <> date9.Month Or date8.Year <> date9.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkode4 = "" Or txtkodecomp = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Giro berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Giro ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    pro4.Visible = True
    jml = 0
    pro4.Value = 0
    
    'cair
    OBJ.Open dsn
    SQL = "select * from am_cashsub where typebayar='G' and year(tgltolak)=1900 and month(tgltolak)=1 and tglcair>='" & tanggal8 & "' and tglcair<='" & tanggal9 & "' order by tglcair"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
        DoEvents
    Loop
    
    RST.MoveFirst
    pro4.Max = jml
 
    End If
    
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "'"
        SQL1 = SQL1 + " and kdtrx = '" & txtkode4 & "' and notrx like '" & Format(RST!tglcair, "YYMM/") & "%' order by notrx desc"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str7 = Right(RST1!notrx, 3) Else str7 = 0
        OBJ1.Close
        
        str7 = str7 + 1
        If Len(str7) = 1 Then str8 = "000" & str7
        If Len(str7) = 2 Then str8 = "00" & str7
        If Len(str7) = 3 Then str8 = "0" & str7
        If Len(str7) = 4 Then str8 = str7
        
        OBJ1.Open dsn
        SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=7"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then str6 = RST1!noacc Else str6 = ""
        
        SQL1 = "select kodecust,nilaikurs from am_cashhdr where nobkt='" & RST!nobkt & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            txtnilai5 = RST1!nilaikurs
            str4 = RST1!kodecust
            
            SQL1 = "select namacust from am_customer where kodecust='" & str4 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str4 = " (" & str4 & "/" & RST1!namacust & ")"
        Else
            str4 = ""
        End If
        OBJ1.Close
        
        str5 = RST!nogiro & " Cair " & str4
        If Len(str5) > 60 Then str5 = Mid(str5, 1, 60)
        
        OBJ1.Open dsn
        SQL1 = "select kode from am_bank where acc = '" & RST!acbank & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            str3 = RST1!kode
        
            SQL1 = "select noacc from am_autoaccbank where kodebank = '" & str3 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
        Else
            str3 = ""
        End If
        OBJ1.Close
        
        OBJ1.Open dsn
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "lineitem)"
        
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtkodecomp & "',"
        SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tglcair) & "/" & Day(RST!tglcair) & "/" & Year(RST!tglcair) & "'),"
        SQL1 = SQL1 + "'" & txtkode4 & "',"
        SQL1 = SQL1 + "'" & Format(RST!tglcair, "YYMM") & "/" & str8 & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str3 & "',"
        SQL1 = SQL1 + "'" & str5 & "',"
        SQL1 = SQL1 + "'D',"
        SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'" & RST!nogiro & "',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'1'))"
        Set RST1 = OBJ1.Execute(SQL1)
        
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "lineitem)"
        
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtkodecomp & "',"
        SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tglcair) & "/" & Day(RST!tglcair) & "/" & Year(RST!tglcair) & "'),"
        SQL1 = SQL1 + "'" & txtkode4 & "',"
        SQL1 = SQL1 + "'" & Format(RST!tglcair, "YYMM") & "/" & str8 & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str6 & "',"
        SQL1 = SQL1 + "'" & str5 & "',"
        SQL1 = SQL1 + "'K',"
        SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'" & RST!nogiro & "',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'2'))"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        OBJ1.Open dsn
        SQL1 = "update am_cashsub set tgltolak='12/01/1900' where nobkt='" & RST!nobkt & "' and nogiro='" & RST!nogiro & "' and typebayar='G' and year(tgltolak)=1900 and month(tgltolak)=1"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        pro4.Value = pro4.Value + 1
        If pro4.Value = jml Then pro4.Value = 0
        
        RST.MoveNext
        DoEvents
    Loop
    OBJ.Close
    
    jml = 0
    'tolak
            OBJ.Open dsn
            SQL = "select * from am_cashsub where typebayar='G' and year(tglcair)=1900 and month(tglcair)=1 and tgltolak>='" & tanggal8 & "' and tgltolak<='" & tanggal9 & "' order by tgltolak"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
            Do While Not RST.EOF
                jml = jml + 1
                RST.MoveNext
                DoEvents
            Loop
            RST.MoveFirst
            pro4.Max = jml
            End If
            
            Do While Not RST.EOF
                OBJ1.Open dsn
                SQL1 = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "'"
                SQL1 = SQL1 + " and kdtrx = '" & txtkode4 & "' and notrx like '" & Format(RST!tgltolak, "YYMM/") & "%' order by notrx desc"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str7 = Right(RST1!notrx, 3) Else str7 = 0
                OBJ1.Close
                
                str7 = str7 + 1
                If Len(str7) = 1 Then str8 = "000" & str7
                If Len(str7) = 2 Then str8 = "00" & str7
                If Len(str7) = 3 Then str8 = "0" & str7
                If Len(str7) = 4 Then str8 = str7
            
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=7"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str6 = RST1!noacc Else str6 = ""
                
                SQL1 = "select kodecust,nilaikurs from am_cashhdr where nobkt='" & RST!nobkt & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnilai5 = RST1!nilaikurs
                    str4 = RST1!kodecust
                    
                    SQL1 = "select namacust from am_customer where kodecust='" & str4 & "'"
                    Set RST1 = OBJ1.Execute(SQL1)
                    If Not RST1.EOF Then str4 = " (" & str4 & "/" & RST1!namacust & ")"
                Else
                    str4 = ""
                End If
                OBJ1.Close
                
                str5 = RST!nogiro & " Tolak " & str4
                If Len(str5) > 60 Then str5 = Mid(str5, 1, 60)
                
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=8"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                OBJ1.Close
                
                OBJ1.Open dsn
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tgltolak) & "/" & Day(RST!tgltolak) & "/" & Year(RST!tgltolak) & "'),"
                SQL1 = SQL1 + "'" & txtkode4 & "',"
                SQL1 = SQL1 + "'" & Format(RST!tgltolak, "YYMM") & "/" & str8 & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"
                SQL1 = SQL1 + "'" & str5 & "',"
                SQL1 = SQL1 + "'D',"
                SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'" & RST!nogiro & "',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'1'))"
                Set RST1 = OBJ1.Execute(SQL1)
                
                SQL1 = "insert into gl_transaksi "
                SQL1 = SQL1 + "(kdcomp, "
                SQL1 = SQL1 + "tgltrx, "
                SQL1 = SQL1 + "kdtrx, "
                SQL1 = SQL1 + "notrx, "
                SQL1 = SQL1 + "kurs, "
                SQL1 = SQL1 + "noactrx, "
                SQL1 = SQL1 + "desctrx, "
                SQL1 = SQL1 + "dbkrtrx, "
                SQL1 = SQL1 + "amounttrx, "
                SQL1 = SQL1 + "nilaitrx, "
                SQL1 = SQL1 + "currtrx, "
                SQL1 = SQL1 + "flag, "
                SQL1 = SQL1 + "flagprint, "
                SQL1 = SQL1 + "flagadjustment, "
                SQL1 = SQL1 + "cekbg, "
                SQL1 = SQL1 + "identry, "
                SQL1 = SQL1 + "idupdate, "
                SQL1 = SQL1 + "dateentry, "
                SQL1 = SQL1 + "dateupdate, "
                SQL1 = SQL1 + "lineitem)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + "('" & txtkodecomp & "',"
                SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tgltolak) & "/" & Day(RST!tgltolak) & "/" & Year(RST!tgltolak) & "'),"
                SQL1 = SQL1 + "'" & txtkode4 & "',"
                SQL1 = SQL1 + "'" & Format(RST!tgltolak, "YYMM") & "/" & str8 & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str6 & "',"
                SQL1 = SQL1 + "'" & str5 & "',"
                SQL1 = SQL1 + "'K',"
                SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST!jumlah * txtnilai5 & "'),"
                SQL1 = SQL1 + "'" & str1 & "',"
                SQL1 = SQL1 + "'B',"
                SQL1 = SQL1 + "'J',"
                SQL1 = SQL1 + "'0',"
                SQL1 = SQL1 + "'" & RST!nogiro & "',"
                SQL1 = SQL1 + "'auto',"
                SQL1 = SQL1 + "'',"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
                SQL1 = SQL1 + "convert(numeric,'2'))"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
        
                OBJ1.Open dsn
                SQL1 = "update am_cashsub set tglcair='12/01/1900' where nobkt='" & RST!nobkt & "' and nogiro='" & RST!nogiro & "' and typebayar='G' and year(tglcair)=1900 and month(tglcair)=1"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
                
                pro4.Value = pro4.Value + 1
                If pro4.Value = jml Then pro4.Value = 0
                
                RST.MoveNext
            Loop
            OBJ.Close
err_msg:
    MsgBox "Proses Complete.", vbInformation, "Information"
    pro4.Visible = False
    pro4.Value = 0
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecomp = hasil
    lblnamacomp = hasil1
    hasil = ""
    hasil1 = ""
    cariautojurnal
End Sub

Private Sub cmdsearch1_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    lbldesc1 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    lbldesc2 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc3.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc3 = hasil
    lbldesc3 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch5_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc5 = hasil
    lbldesc5 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc6.SetFocus
End Sub

Private Sub cmdsearch6_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch6_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc6 = hasil
    lbldesc6 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtacc1.SetFocus
End Sub

Private Sub cmdverify_Click()
    If txtkodecomp = "" Then Exit Sub
    
    If date4 > date5 Then
        MsgBox "Invalid date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date4.Month <> date5.Month Or date4.Year <> date5.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Verify Account berlangsung." & vbCrLf & _
    "Lanjutkan Proses Verify Account ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select distinct kodecust from am_cashhdr"
    SQL = SQL + " where tglbkt >= '" & tanggal4 & "' and tglbkt <= '" & tanggal5 & "' and kodebayar='PM'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ0.Open dsn
        SQL0 = "select * from am_autoaccust where kodecomp = '" & txtkodecomp & "' and kodecust = '" & RST!kodecust & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        If RST0.EOF Then
            OBJ0.Close
            OBJ.Close
            MsgBox "There is undefine account customer.", vbInformation, "Information"
            Exit Sub
        Else
            If RST0!noacc = "" Then
                OBJ0.Close
                OBJ.Close
                MsgBox "There is undefine account customer.", vbInformation, "Information"
                Exit Sub
            End If
        End If
        OBJ0.Close

        RST.MoveNext
    Loop
    OBJ.Close
    
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Row = 1
    Grid1.TextMatrix(0, 0) = "No Bayar"
    Grid1.TextMatrix(0, 1) = "Tgl Bayar"
    Grid1.TextMatrix(0, 2) = "No Invoice"
    Grid1.ColWidth(0) = 2000
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 2000
    Grid1.RowHeightMin = 300
    
    OBJ.Open dsn
    SQL = "select a.nobkt,a.tglbkt,b.noapply from am_cashhdr a left join am_cashlin b on a.nobkt=b.nobkt and a.kodebayar=b.kodebayar"
    SQL = SQL + " where a.tglbkt>='" & tanggal4 & "' and a.tglbkt<='" & tanggal5 & "' and a.kodebayar='PM' and a.posted='1'"
    SQL = SQL + " order by a.tglbkt,a.nobkt,b.noapply"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        Grid1.TextMatrix(Grid1.Row, 0) = RST!nobkt
        Grid1.TextMatrix(Grid1.Row, 1) = Format(RST!tglbkt, "dd MMM yyyy")
        Grid1.TextMatrix(Grid1.Row, 2) = RST!noapply
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Row = Grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Verify account and currency complete, continue verify Transaksi GL.", vbInformation, "Information"
    
    pro2.Visible = True
    pro2.Value = 0
    
    OBJ.Open dsn
    SQL = "select nobkt from am_cashhdr where kodebayar='PM' and tglbkt>='" & tanggal4 & "' and tglbkt<='" & tanggal5 & "'"
    SQL = SQL + " and posted='0' and idupdate='0'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
    Loop
    
    RST.MoveFirst
    pro2.Max = jml
    End If
    
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode2 & "' and notrx='" & RST!nobkt & "' and flag<>'B'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            OBJ1.Close
            OBJ.Close
            
            pro2.Visible = False
            pro2.Value = 0
            pro2.Max = 0
            
            MsgBox "Transaksi pembayaran sudah ada di GL dan sudah terposting." & vbCrLf & _
            "proses di batalkan, unposting dahulu transaksi GL.", vbInformation, "Proses Batal"
            
            Exit Sub
        End If
        OBJ1.Close
        
        pro2.Value = pro2.Value + 1
        If pro2.Value = jml Then pro2.Value = 0
        RST.MoveNext
    Loop
    OBJ.Close
    
    pro2.Visible = False
    pro2.Value = 0
    pro2.Max = 0
    
    MsgBox "Verify Transaksi GL complete.", vbInformation, "Information"
    
    cmdproses2.Enabled = True
    cmdverify.Enabled = False
    cmdproses2.SetFocus
End Sub

Private Sub cmdverifyGT_Click()
    If txtkodecomp = "" Then Exit Sub
    
    If date6 > date7 Then
        MsgBox "Invalid date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date6.Month <> date7.Month Or date6.Year <> date7.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Verify Transaksi GL berlangsung." & vbCrLf & _
    "Lanjutkan Proses Verify Transaksi GL ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    pro3.Visible = True
    jml = 0
    pro3.Value = 0
    
    OBJ.Open dsn
    SQL = "select nobkt from am_cashhdr where kodebayar='GT' and tglbkt>='" & tanggal6 & "' and tglbkt<='" & tanggal7 & "'"
    SQL = SQL + " and posted='9' and idupdate='0'"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
    Loop
        pro3.Max = jml
        RST.MoveFirst
    End If
    
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select notrx from gl_transaksi where kdcomp='" & txtkodecomp & "' and kdtrx='" & txtkode3 & "' and notrx='" & RST!nobkt & "' and flag<>'B'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            OBJ1.Close
            OBJ.Close
            
            pro3.Visible = False
            pro3.Value = 0
            pro3.Max = 0
            
            MsgBox "Transaksi ganti tolak sudah ada di GL dan sudah terposting." & vbCrLf & _
            "proses di batalkan, unposting dahulu transaksi GL.", vbInformation, "Proses Batal"
            
            Exit Sub
        End If
        OBJ1.Close
        
        pro3.Value = pro3.Value + 1
        If pro3.Value = jml Then pro3.Value = 0
        RST.MoveNext
    Loop
    OBJ.Close
    
    pro3.Visible = False
    pro3.Value = 0
    pro3.Max = 0
    
    MsgBox "Verify Transaksi GL complete.", vbInformation, "Information"
    
    cmdproses3.Enabled = True
    cmdverifyGT.Enabled = False
    cmdproses3.SetFocus
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
Private Sub setgrid2()
    grid2.TextMatrix(0, 0) = "No Bukti"
    grid2.TextMatrix(0, 1) = "Tgl Bukti"
    grid2.TextMatrix(0, 2) = "Kode"
    grid2.TextMatrix(0, 3) = "Keterangan"
    grid2.TextMatrix(0, 4) = "Amount"
    grid2.TextMatrix(0, 5) = "Account"
    
    grid2.ColAlignmentFixed(2) = flexAlignCenterCenter
    grid2.ColAlignmentFixed(4) = flexAlignRightCenter
    grid2.ColAlignment(1) = flexAlignLeftCenter
    grid2.ColAlignment(2) = flexAlignCenterCenter
    grid2.ColAlignment(4) = flexAlignRightCenter
    
    grid2.ColWidth(0) = 1000
    grid2.ColWidth(1) = 1200
    grid2.ColWidth(2) = 800
    grid2.ColWidth(3) = 3500
    grid2.ColWidth(4) = 2000
    grid2.ColWidth(5) = 1200
    grid2.RowHeightMin = 300
End Sub
Private Sub Form_Load()
    
    Grid1.TextMatrix(0, 0) = "No Bayar"
    Grid1.TextMatrix(0, 1) = "Tgl Bayar"
    Grid1.TextMatrix(0, 2) = "No Invoice"
    Grid1.ColWidth(0) = 2000
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 2000
    Grid1.RowHeightMin = 300
    
    setgrid2
    
    date1.Value = Date
    date2.Value = Date
    date4.Value = Date
    date5.Value = Date
    date6.Value = Date
    date7.Value = Date
    date8.Value = Date
    date9.Value = Date
    
    OBJ.Open dsn
    SQL = "select kdkurs from gl_kurs where base='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str1 = RST!kdkurs
    OBJ.Close
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    Select Case grid2.Col
        Case 5:
            If txtacc.Visible = True Then Exit Sub

            txtacc.Width = grid2.ColWidth(grid2.Col) - 40
            txtacc = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtacc.Left = grid2.Left + grid2.CellLeft
            txtacc.Top = grid2.Top + grid2.CellTop + 20
            txtacc.Visible = True
            txtacc.SetFocus
    End Select
End Sub

Private Sub txtacc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtacc = "" Then
            OBJ.Open dsn
            
            OBJ.Close
        Else
        
        End If
    End If
End Sub

Private Sub txtacc_LostFocus()
    txtacc.Visible = False
End Sub

Private Sub txtacc00_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc01.SetFocus
End Sub

Private Sub txtacc00_LostFocus()
    If txtacc00 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc00 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc00 = ""
        lblacc00 = ""
        txtacc00.SetFocus
    Else
        lblacc00 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc01_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc02.SetFocus
End Sub

Private Sub txtacc01_LostFocus()
    If txtacc01 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc01 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc01 = ""
        lblacc01 = ""
        txtacc01.SetFocus
    Else
        lblacc01 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc02_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc03.SetFocus
End Sub

Private Sub txtacc02_LostFocus()
    If txtacc02 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc02 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc02 = ""
        lblacc02 = ""
        txtacc02.SetFocus
    Else
        lblacc02 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc03_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc04.SetFocus
End Sub

Private Sub txtacc03_LostFocus()
    If txtacc03 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc03 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc03 = ""
        lblacc03 = ""
        txtacc03.SetFocus
    Else
        lblacc03 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc04_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc05.SetFocus
End Sub

Private Sub txtacc04_LostFocus()
    If txtacc04 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc04 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc04 = ""
        lblacc04 = ""
        txtacc04.SetFocus
    Else
        lblacc04 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc05_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc06.SetFocus
End Sub

Private Sub txtacc05_LostFocus()
    If txtacc05 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc05 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc05 = ""
        lblacc05 = ""
        txtacc05.SetFocus
    Else
        lblacc05 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc06_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc07.SetFocus
End Sub

Private Sub txtacc06_LostFocus()
    If txtacc06 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc06 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc06 = ""
        lblacc06 = ""
        txtacc06.SetFocus
    Else
        lblacc06 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc07_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdadd.SetFocus
End Sub

Private Sub txtacc07_LostFocus()
    If txtacc07 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc07 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc07 = ""
        lblacc07 = ""
        txtacc07.SetFocus
    Else
        lblacc07 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc2.SetFocus
End Sub

Private Sub txtacc1_LostFocus()
    If txtacc1 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc1 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc1 = ""
        lbldesc1 = ""
        txtacc1.SetFocus
    Else
        lbldesc1 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc3.SetFocus
End Sub

Private Sub txtacc2_LostFocus()
    If txtacc2 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc2 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc2 = ""
        lbldesc2 = ""
        txtacc2.SetFocus
    Else
        lbldesc2 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdadd1.SetFocus
End Sub

Private Sub txtacc3_LostFocus()
    If txtacc3 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc3 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc3 = ""
        lbldesc3 = ""
        txtacc3.SetFocus
    Else
        lbldesc3 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc6.SetFocus
End Sub

Private Sub txtacc5_LostFocus()
    If txtacc5 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc5 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc5 = ""
        lbldesc5 = ""
        txtacc5.SetFocus
    Else
        lbldesc5 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtacc6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc1.SetFocus
End Sub

Private Sub txtacc6_LostFocus()
    If txtacc6 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc6 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc6 = ""
        lbldesc6 = ""
        txtacc6.SetFocus
    Else
        lbldesc6 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdmanual.SetFocus
End Sub

Private Sub txtKode1_LostFocus()
    If txtkode1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdtrx = '" & txtkode1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Kode Transaksi " & txtkode1 & " already exist, continue anyway?", vbQuestion + vbYesNo, "Question") = vbNo Then
            txtkode1 = ""
            txtkode1.SetFocus
        End If
    End If
    OBJ.Close
End Sub

Private Sub txtkode2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdverify.SetFocus
End Sub

Private Sub txtkode2_LostFocus()
    If txtkode2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdtrx = '" & txtkode2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Kode Transaksi " & txtkode1 & " already exist, continue anyway?", vbQuestion + vbYesNo, "Question") = vbNo Then
            txtkode2 = ""
            txtkode2.SetFocus
        End If
    End If
    OBJ.Close
End Sub

Private Sub txtkode3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdverify.SetFocus
End Sub

Private Sub txtkode3_LostFocus()
    If txtkode3 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdtrx = '" & txtkode3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Kode Transaksi " & txtkode2 & " alreday exist.", vbInformation, "Information"
        txtkode3 = ""
        txtkode3.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkodecomp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SSTab1.SetFocus
End Sub

Private Sub txtkodecomp_LostFocus()
    If txtkodecomp = "" Then Exit Sub
    If txtkodecomp.SelLength <> 0 Then Exit Sub
        
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacomp = RST!nmcompscr
    Else
        MsgBox "Company " & txtkodecomp & " Not Found.", vbInformation, "Information"
        txtkodecomp = ""
        txtkodecomp.SetFocus
    End If
    OBJ.Close
    
    cariautojurnal
End Sub

Private Sub cariautojurnal()
    If txtkodecomp = "" Then Exit Sub
    'jurnal piuutang
    OBJ.Open dsn
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=2"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc00 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=3"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc01 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=4"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc02 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=5"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc03 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=6"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc04 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=7"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc05 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=8"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc06 = RST!noacc
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'piutang' and line=9"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc07 = RST!noacc
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc00 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc00 = RST!nmac Else lblacc00 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc01 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc01 = RST!nmac Else lblacc01 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc02 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc02 = RST!nmac Else lblacc02 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc03 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc03 = RST!nmac Else lblacc03 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc04 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc04 = RST!nmac Else lblacc04 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc05 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc05 = RST!nmac Else lblacc05 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc06 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc06 = RST!nmac Else lblacc06 = ""
    
    SQL = "SELECT * FROM gl_masterac where noac = '" & txtacc07 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lblacc07 = RST!nmac Else lblacc07 = ""
    OBJ.Close
    
    'jurnal bayar piutang
    OBJ.Open dsn
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=2"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc1 = RST!noacc Else txtacc1 = ""
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=4"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc2 = RST!noacc Else txtacc2 = ""
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=5"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc3 = RST!noacc Else txtacc3 = ""
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=7"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc5 = RST!noacc Else txtacc5 = ""
    
    SQL = "select * from am_autoo where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'bpiutang' and line=8"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc6 = RST!noacc Else txtacc6 = ""
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc1 = RST!nmac Else lbldesc1 = ""
    
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc2 = RST!nmac Else lbldesc2 = ""

    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc3 = RST!nmac Else lbldesc3 = ""
    
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc5 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc5 = RST!nmac Else lbldesc5 = ""
    
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc6 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc6 = RST!nmac Else lbldesc6 = ""
    OBJ.Close
    
    SSTab1.Tab = 1
    SSTab1.Tab = 0
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggal3()
    tanggal3 = Month(date3) & "/" & Day(date3) & "/" & Year(date3)
End Function

Function tanggal4()
    tanggal4 = Month(date4) & "/" & Day(date4) & "/" & Year(date4)
End Function

Function tanggal5()
    tanggal5 = Month(date5) & "/" & Day(date5) & "/" & Year(date5)
End Function

Function tanggal6()
    tanggal6 = Month(date6) & "/" & Day(date6) & "/" & Year(date6)
End Function

Function tanggal7()
    tanggal7 = Month(date7) & "/" & Day(date7) & "/" & Year(date7)
End Function

Function tanggal8()
    tanggal8 = Month(date8) & "/" & Day(date8) & "/" & Year(date8)
End Function

Function tanggal9()
    tanggal9 = Month(date9) & "/" & Day(date9) & "/" & Year(date9)
End Function

Function tanggal10()
    tanggal10 = Month(date10) & "/" & Day(date10) & "/" & Year(date10)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
