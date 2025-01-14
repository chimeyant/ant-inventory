VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdefinejurnal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Journal"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Jurnal Hutang"
      TabPicture(0)   =   "frmdefinejurnal.frx":2372
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "blank"
      Tab(0).Control(1)=   "check"
      Tab(0).Control(2)=   "uncheck"
      Tab(0).Control(3)=   "grid"
      Tab(0).Control(4)=   "cmdadd"
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(7)=   "Label23"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Jurnal Bayar Hutang"
      TabPicture(1)   =   "frmdefinejurnal.frx":238E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtacc4"
      Tab(1).Control(1)=   "txtacc3"
      Tab(1).Control(2)=   "txtacc1"
      Tab(1).Control(3)=   "txtacc2"
      Tab(1).Control(4)=   "cmdadd1"
      Tab(1).Control(5)=   "cmdsearch1"
      Tab(1).Control(6)=   "cmdsearch2"
      Tab(1).Control(7)=   "cmdsearch3"
      Tab(1).Control(8)=   "cmdsearch4"
      Tab(1).Control(9)=   "lbldesc4"
      Tab(1).Control(10)=   "lbldesc3"
      Tab(1).Control(11)=   "lbldesc2"
      Tab(1).Control(12)=   "lbldesc1"
      Tab(1).Control(13)=   "Label20"
      Tab(1).Control(14)=   "Label19"
      Tab(1).Control(15)=   "Label18"
      Tab(1).Control(16)=   "Label17"
      Tab(1).Control(17)=   "Label16"
      Tab(1).Control(18)=   "Label15"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Proses Jurnal Hutang"
      TabPicture(2)   =   "frmdefinejurnal.frx":23AA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtnilai3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "date3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtnilai2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdmanual"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "grid2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdproses1"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "pro1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "date2"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "date1"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtno1"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtkode1"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtnilai1"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Proses Jurnal Bayar Hutang"
      TabPicture(3)   =   "frmdefinejurnal.frx":23C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtkode2"
      Tab(3).Control(1)=   "pro2"
      Tab(3).Control(2)=   "cmdproses2"
      Tab(3).Control(3)=   "date4"
      Tab(3).Control(4)=   "date5"
      Tab(3).Control(5)=   "grid1"
      Tab(3).Control(6)=   "cmdverify"
      Tab(3).Control(7)=   "txtnilai4"
      Tab(3).Control(8)=   "txtnilai5"
      Tab(3).Control(9)=   "txtnilai6"
      Tab(3).Control(10)=   "txtnilai7"
      Tab(3).Control(11)=   "Label21"
      Tab(3).Control(12)=   "Label9"
      Tab(3).Control(13)=   "Label8"
      Tab(3).Control(14)=   "Label7"
      Tab(3).Control(15)=   "Label6"
      Tab(3).Control(16)=   "Label5"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Proses Jurnal Koreksi Hutang"
      TabPicture(4)   =   "frmdefinejurnal.frx":23E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label10"
      Tab(4).Control(1)=   "Label11"
      Tab(4).Control(2)=   "Label12"
      Tab(4).Control(3)=   "Label14"
      Tab(4).Control(4)=   "Label22"
      Tab(4).Control(5)=   "grid3"
      Tab(4).Control(6)=   "cmdverifykoreksi"
      Tab(4).Control(7)=   "cmdproses3"
      Tab(4).Control(8)=   "pro3"
      Tab(4).Control(9)=   "date7"
      Tab(4).Control(10)=   "date6"
      Tab(4).Control(11)=   "txtno3"
      Tab(4).Control(12)=   "txtkode3"
      Tab(4).Control(13)=   "Option1"
      Tab(4).Control(14)=   "Option2"
      Tab(4).ControlCount=   15
      Begin VB.TextBox txtacc4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtacc3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtacc1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtacc2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Debit Note"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credit Note"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtkode3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   30
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtno3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73560
         MaxLength       =   8
         TabIndex        =   31
         Text            =   "YYMM/999"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtkode2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   22
         Top             =   1440
         Width           =   375
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai1 
         Height          =   255
         Left            =   9840
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":23FE
         Caption         =   "frmdefinejurnal.frx":241E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":248A
         Keys            =   "frmdefinejurnal.frx":24A8
         Spin            =   "frmdefinejurnal.frx":24EA
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
      Begin VB.TextBox txtkode1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtno1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   16
         Text            =   "YYMM/999"
         Top             =   1440
         Width           =   975
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
         Left            =   -75000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   40
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
         Left            =   -75000
         Picture         =   "frmdefinejurnal.frx":2512
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   960
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
         Left            =   -75000
         Picture         =   "frmdefinejurnal.frx":27F4
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7435
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
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin akProgress.akProgressBar pro1 
         Height          =   375
         Left            =   2040
         TabIndex        =   43
         Top             =   4320
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         BackColour      =   16744576
         FontColour      =   4210752
         BarColour       =   16761024
         Horizontal      =   -1  'True
         ReverseGradient =   0   'False
         Max             =   100
         Min             =   0
         GapWidth        =   3
         LineWidth       =   7
         Caption         =   1
         BorderStyle     =   0
         Margin          =   2
         Gradient        =   0
         Alignment       =   2
      End
      Begin Chameleon.chameleonButton cmdproses1 
         Height          =   375
         Left            =   9480
         TabIndex        =   19
         Top             =   4320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Hutang"
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
         MICON           =   "frmdefinejurnal.frx":2B42
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin akProgress.akProgressBar pro2 
         Height          =   375
         Left            =   -74880
         TabIndex        =   46
         Top             =   4320
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   661
         BackColour      =   8421631
         FontColour      =   4210752
         BarColour       =   12632319
         Horizontal      =   -1  'True
         ReverseGradient =   0   'False
         Max             =   100
         Min             =   0
         GapWidth        =   3
         LineWidth       =   7
         Caption         =   1
         BorderStyle     =   0
         Margin          =   2
         Gradient        =   0
         Alignment       =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   9
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
         _Band(0).Cols   =   9
      End
      Begin Chameleon.chameleonButton cmdproses2 
         Height          =   375
         Left            =   -66000
         TabIndex        =   25
         Top             =   4320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Bayar Hutang"
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
         MICON           =   "frmdefinejurnal.frx":2E5C
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
         Left            =   105
         TabIndex        =   18
         Top             =   4305
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Manual Jurnal Hutang"
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
         MICON           =   "frmdefinejurnal.frx":3176
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
         Left            =   9840
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":3490
         Caption         =   "frmdefinejurnal.frx":34B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":351C
         Keys            =   "frmdefinejurnal.frx":353A
         Spin            =   "frmdefinejurnal.frx":357C
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
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   7920
         TabIndex        =   49
         Top             =   600
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdadd 
         Height          =   375
         Left            =   -65760
         TabIndex        =   3
         Top             =   4320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Simpan Jurnal Hutang"
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
         MICON           =   "frmdefinejurnal.frx":35A4
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
         Left            =   9840
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":38BE
         Caption         =   "frmdefinejurnal.frx":38DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":394A
         Keys            =   "frmdefinejurnal.frx":3968
         Spin            =   "frmdefinejurnal.frx":39AA
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
      Begin MSComCtl2.DTPicker date4 
         Height          =   285
         Left            =   -73560
         TabIndex        =   20
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date5 
         Height          =   285
         Left            =   -73560
         TabIndex        =   21
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin Chameleon.chameleonButton cmdadd1 
         Height          =   375
         Left            =   -66120
         TabIndex        =   12
         Top             =   4200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Simpan Jurnal Bayar Hutang"
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
         MICON           =   "frmdefinejurnal.frx":39D2
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
         TabIndex        =   28
         Top             =   1320
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin MSComCtl2.DTPicker date7 
         Height          =   285
         Left            =   -73560
         TabIndex        =   29
         Top             =   1680
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
         Format          =   134807555
         CurrentDate     =   37426
      End
      Begin akProgress.akProgressBar pro3 
         Height          =   375
         Left            =   -74880
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   661
         BackColour      =   16744703
         FontColour      =   4210752
         BarColour       =   16761087
         Horizontal      =   -1  'True
         ReverseGradient =   0   'False
         Max             =   100
         Min             =   0
         GapWidth        =   3
         LineWidth       =   7
         Caption         =   1
         BorderStyle     =   0
         Margin          =   2
         Gradient        =   0
         Alignment       =   2
      End
      Begin Chameleon.chameleonButton cmdproses3 
         Height          =   375
         Left            =   -66000
         TabIndex        =   34
         Top             =   4320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses Jurnal Koreksi Hutang"
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
         MICON           =   "frmdefinejurnal.frx":3CEC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   3495
         Left            =   -70560
         TabIndex        =   24
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   4
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
         _Band(0).Cols   =   4
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   -71280
         TabIndex        =   5
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
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
         MICON           =   "frmdefinejurnal.frx":4006
         PICN            =   "frmdefinejurnal.frx":4022
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
         Height          =   285
         Left            =   -71280
         TabIndex        =   7
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
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
         MICON           =   "frmdefinejurnal.frx":4134
         PICN            =   "frmdefinejurnal.frx":4150
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
         Height          =   285
         Left            =   -71280
         TabIndex        =   9
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
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
         MICON           =   "frmdefinejurnal.frx":4262
         PICN            =   "frmdefinejurnal.frx":427E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch4 
         Height          =   285
         Left            =   -71280
         TabIndex        =   11
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
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
         MICON           =   "frmdefinejurnal.frx":4390
         PICN            =   "frmdefinejurnal.frx":43AC
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
         Height          =   375
         Left            =   -72960
         TabIndex        =   23
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verify Account Supplier >>"
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
         MICON           =   "frmdefinejurnal.frx":44BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai4 
         Height          =   255
         Left            =   -72240
         TabIndex        =   68
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":47D8
         Caption         =   "frmdefinejurnal.frx":47F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":4864
         Keys            =   "frmdefinejurnal.frx":4882
         Spin            =   "frmdefinejurnal.frx":48C4
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
      Begin TDBNumber6Ctl.TDBNumber txtnilai5 
         Height          =   255
         Left            =   -72240
         TabIndex        =   73
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":48EC
         Caption         =   "frmdefinejurnal.frx":490C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":4978
         Keys            =   "frmdefinejurnal.frx":4996
         Spin            =   "frmdefinejurnal.frx":49D8
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
      Begin TDBNumber6Ctl.TDBNumber txtnilai6 
         Height          =   255
         Left            =   -72240
         TabIndex        =   74
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":4A00
         Caption         =   "frmdefinejurnal.frx":4A20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":4A8C
         Keys            =   "frmdefinejurnal.frx":4AAA
         Spin            =   "frmdefinejurnal.frx":4AEC
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
      Begin TDBNumber6Ctl.TDBNumber txtnilai7 
         Height          =   255
         Left            =   -72240
         TabIndex        =   75
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         Calculator      =   "frmdefinejurnal.frx":4B14
         Caption         =   "frmdefinejurnal.frx":4B34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefinejurnal.frx":4BA0
         Keys            =   "frmdefinejurnal.frx":4BBE
         Spin            =   "frmdefinejurnal.frx":4C00
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
      Begin Chameleon.chameleonButton cmdverifykoreksi 
         Height          =   375
         Left            =   -72960
         TabIndex        =   32
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verify Account Supplier >>"
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
         MICON           =   "frmdefinejurnal.frx":4C28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
         Height          =   3495
         Left            =   -70560
         TabIndex        =   33
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   4
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
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label25 
         Caption         =   "Jurnal A = Persediaan"
         Height          =   255
         Left            =   -65760
         TabIndex        =   80
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Jurnal B = PPn Masukan"
         Height          =   255
         Left            =   -65760
         TabIndex        =   79
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label23 
         Caption         =   "Jurnal C = Hutang Dagang"
         Height          =   255
         Left            =   -65760
         TabIndex        =   78
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Define Account Supplier"
         Height          =   255
         Left            =   -67800
         TabIndex        =   77
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Define Account Supplier"
         Height          =   255
         Left            =   -67800
         TabIndex        =   76
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lbldesc4 
         Height          =   255
         Left            =   -70560
         TabIndex        =   72
         Top             =   2310
         Width           =   6615
      End
      Begin VB.Label lbldesc3 
         Height          =   255
         Left            =   -70560
         TabIndex        =   71
         Top             =   1950
         Width           =   6615
      End
      Begin VB.Label lbldesc2 
         Height          =   255
         Left            =   -70560
         TabIndex        =   70
         Top             =   1590
         Width           =   6615
      End
      Begin VB.Label lbldesc1 
         Height          =   255
         Left            =   -70560
         TabIndex        =   69
         Top             =   870
         Width           =   6615
      End
      Begin VB.Label Label20 
         Caption         =   "Biaya Administrasi (Debet)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   2310
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Selisih Bayar (Debet/Kredit)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   1950
         Width           =   2055
      End
      Begin VB.Label Label18 
         Caption         =   "Potongan Bayar (Kredit)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   1590
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Kas/Bank/Cek/Giro (Kredit)   = Define Account Bank/Cash"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   1230
         Width           =   4215
      End
      Begin VB.Label Label16 
         Caption         =   "Selisih Kurs (Debet/Kredit)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   870
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Account Hutang (Debet)      = Define Account Supplier"
         Height          =   255
         Left            =   -74760
         TabIndex        =   62
         Top             =   510
         Width           =   4095
      End
      Begin VB.Label Label14 
         Caption         =   "Kode Transaksi            (2 character)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   60
         Top             =   2190
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "(YY=Year, MM=Month, 999=Counter)"
         Height          =   255
         Left            =   -73560
         TabIndex        =   59
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   56
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   55
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "No Transaksi di GL sesuai dengan No Bukti Pembayaran"
         Height          =   255
         Left            =   -74760
         TabIndex        =   54
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Kode Transaksi            (2 character)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   53
         Top             =   1470
         Width           =   3015
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -67680
         TabIndex        =   52
         Top             =   1470
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   7320
         TabIndex        =   50
         Top             =   1470
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Kode Transaksi   (2 character [JB])"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1470
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "(YY=Year, MM=Month, 999=Counter)"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   1470
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Sampai Tanggal"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Dari Tanggal"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   630
         Width           =   1455
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   10440
      TabIndex        =   35
      Top             =   5400
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
      MICON           =   "frmdefinejurnal.frx":4F42
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
      Caption         =   "frmdefinejurnal.frx":525C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdefinejurnal.frx":52C8
      Key             =   "frmdefinejurnal.frx":52E6
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
      TabIndex        =   44
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
      MICON           =   "frmdefinejurnal.frx":5322
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
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
      TabIndex        =   45
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

Dim SP As New ADODB.Command
Dim vsp(2) As Variant

Dim posrow As String
Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, str12, str13, str14, str15 As String
Dim hitung1, int1, int2, int3, int4, int5, int6 As Integer

Private Sub cmdadd_Click()
    If txtkodecomp = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 1) = "" Or grid.TextMatrix(grid.Row, 2) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
                
        If grid.TextMatrix(grid.Row, 3) = "" Or grid.TextMatrix(grid.Row, 4) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal a'"
    Set RST = OBJ.Execute(SQL)
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal b'"
    Set RST = OBJ.Execute(SQL)
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal c'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        OBJ.Open dsn
        SQL = "insert into am_auto ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "jurnal_, "
        SQL = SQL + "noacc, "
        SQL = SQL + "dk, "
        SQL = SQL + "kdkurs, "
        SQL = SQL + "nanti, "
        SQL = SQL + "line)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        grid.Row = grid.Row + 1
    Loop

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdadd1_Click()
    If txtkodecomp = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If txtacc1 = "" And txtacc2 = "" And txtacc3 = "" And txtacc4 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal d'"
    Set RST = OBJ.Execute(SQL)
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal e'"
    Set RST = OBJ.Execute(SQL)
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal f'"
    Set RST = OBJ.Execute(SQL)
    SQL = "delete from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal g'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    If txtacc1 <> "" Then
        OBJ.Open dsn
        SQL = "insert into am_auto ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "jurnal_, "
        SQL = SQL + "noacc, "
        SQL = SQL + "dk, "
        SQL = SQL + "kdkurs, "
        SQL = SQL + "nanti, "
        SQL = SQL + "line)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'Jurnal d',"
        SQL = SQL + "'" & txtacc1 & "',"
        SQL = SQL + "'A',"
        SQL = SQL + "'one',"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'0'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If

    If txtacc2 <> "" Then
        OBJ.Open dsn
        SQL = "insert into am_auto ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "jurnal_, "
        SQL = SQL + "noacc, "
        SQL = SQL + "dk, "
        SQL = SQL + "kdkurs, "
        SQL = SQL + "nanti, "
        SQL = SQL + "line)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'Jurnal e',"
        SQL = SQL + "'" & txtacc2 & "',"
        SQL = SQL + "'A',"
        SQL = SQL + "'one',"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'0'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    
    If txtacc3 <> "" Then
        OBJ.Open dsn
        SQL = "insert into am_auto ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "jurnal_, "
        SQL = SQL + "noacc, "
        SQL = SQL + "dk, "
        SQL = SQL + "kdkurs, "
        SQL = SQL + "nanti, "
        SQL = SQL + "line)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'Jurnal f',"
        SQL = SQL + "'" & txtacc3 & "',"
        SQL = SQL + "'A',"
        SQL = SQL + "'one',"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'0'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    
    If txtacc4 <> "" Then
        OBJ.Open dsn
        SQL = "insert into am_auto ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "jurnal_, "
        SQL = SQL + "noacc, "
        SQL = SQL + "dk, "
        SQL = SQL + "kdkurs, "
        SQL = SQL + "nanti, "
        SQL = SQL + "line)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'Jurnal g',"
        SQL = SQL + "'" & txtacc4 & "',"
        SQL = SQL + "'A',"
        SQL = SQL + "'one',"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'0'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdmanual_Click()
    If txtkodecomp = "" Or txtkode1 = "" Then Exit Sub
    
    If date1 > date2 Then
        MsgBox "Invalid date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date1.Month <> date2.Month Or date1.Year <> date2.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses List Manual Jurnal Hutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses List Manual Jurnal Hutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    grid2.Clear
    grid2.Rows = 2
    
    grid2.TextMatrix(0, 0) = "NoTrx"
    grid2.TextMatrix(0, 1) = "Curr"
    grid2.TextMatrix(0, 2) = "Nilai"
    grid2.TextMatrix(0, 3) = "D/K"
    grid2.TextMatrix(0, 4) = "NoAccount"
    grid2.TextMatrix(0, 5) = "Description"
    grid2.TextMatrix(0, 6) = "TglTrx"
    grid2.TextMatrix(0, 7) = "Kurs"
    
    grid2.ColWidth(0) = 1000
    grid2.ColWidth(1) = 500
    grid2.ColWidth(2) = 1500
    grid2.ColWidth(3) = 500
    grid2.ColWidth(4) = 1200
    grid2.ColWidth(5) = 3500
    grid2.ColWidth(6) = 1200
    grid2.ColWidth(7) = 1000
    grid2.ColWidth(8) = 0
    
    grid2.RowHeightMin = 300
    
    OBJ.Open dsn
    'cek data confirm
    SQL = "select count(nobeli)'hit' from am_beliapp where tglbeli >= '" & tanggal1 & "' and tglbeli <= '" & tanggal2 & "' and flag2 = '1'"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then hitung1 = RST!hit Else hitung1 = 0
    OBJ.Close
    
    
    
    If hitung1 = 0 Then
        MsgBox "Tidak ada data penerimaan yang telah terconfirm untuk di posting.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    pro1.Visible = True
    pro1.Max = 3
    pro1.Value = 0
    
    OBJ.Open dsn
    SQL = "select count(kdkurs)'hitmanual' from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then int1 = RST!hitmanual Else int1 = 0
    OBJ.Close
    
    pro1.Value = pro1.Value + 1
    
    If int1 = 0 Then
        myArray(1, 1) = ""
        myArray(2, 1) = ""
        myArray(3, 1) = ""
        myArray(4, 1) = ""
        myArray(5, 1) = ""
    Else
        Select Case int1
        Case 1
            OBJ.Open dsn
            SQL = "select kdkurs from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                myArray(1, 1) = RST!kdkurs
                myArray(2, 1) = RST!kdkurs
                myArray(3, 1) = RST!kdkurs
                myArray(4, 1) = RST!kdkurs
                myArray(5, 1) = RST!kdkurs
            End If
            OBJ.Close
        Case 2
            x1 = 1
            OBJ.Open dsn
            SQL = "select kdkurs from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                myArray(x1, 1) = RST!kdkurs
                
                myArray(3, 1) = RST!kdkurs
                myArray(4, 1) = RST!kdkurs
                myArray(5, 1) = RST!kdkurs
                
                x1 = x1 + 1
                RST.MoveNext
            Loop
            OBJ.Close
        Case 3
            x1 = 1
            OBJ.Open dsn
            SQL = "select kdkurs from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                myArray(x1, 1) = RST!kdkurs
                
                myArray(4, 1) = RST!kdkurs
                myArray(5, 1) = RST!kdkurs
                
                x1 = x1 + 1
                RST.MoveNext
            Loop
            OBJ.Close
        Case 4
            x1 = 1
            OBJ.Open dsn
            SQL = "select kdkurs from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                myArray(x1, 1) = RST!kdkurs
                
                myArray(5, 1) = RST!kdkurs
                
                x1 = x1 + 1
                RST.MoveNext
            Loop
            OBJ.Close
        Case 5
            x1 = 1
            OBJ.Open dsn
            SQL = "select kdkurs from am_auto where kodecomp = '" & txtkodecomp & "' and nanti = 'x'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                myArray(x1, 1) = RST!kdkurs
                
                x1 = x1 + 1
                RST.MoveNext
            Loop
            OBJ.Close
        End Select
    End If
    
    pro1.Value = pro1.Value + 1
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_postingpenerimaan1"
    vsp(0) = Format(date1, "yyyyMMdd")
    vsp(1) = Format(date2, "yyyyMMdd")
    vsp(2) = txtkodecomp
    SP.Execute , vsp
    Set SP = Nothing
    
    pro1.Value = pro1.Value + 1
        
    OBJ.Open dsn
    SQL = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode1 & "' and notrx like '" & Format(date1, "YYMM/") & "%' order by notrx desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str7 = Right(RST!notrx, 3) Else str7 = 0
    
    pro1.Value = 0
    If int1 <> 0 Then SQL = "select count(nobeli)'hitnobeli' from am_manualposterima where noacc = '' or kodebarang like '" & myArray(1, 1) & "%' or kodebarang like '" & myArray(2, 1) & "%' or kodebarang like '" & myArray(3, 1) & "%' or kodebarang like '" & myArray(4, 1) & "%' or kodebarang like '" & myArray(5, 1) & "%'"
    If int1 = 0 Then SQL = "select count(nobeli)'hitnobeli' from am_manualposterima where noacc = ''"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then pro1.Max = RST!hitnobeli Else pro1.Max = 0
    OBJ.Close
    
    If pro1.Max = 0 Then pro1.Visible = False
    
    grid2.Row = 1
    str9 = ""
    
    OBJ.Open dsn
    If int1 <> 0 Then SQL = "select * from am_manualposterima where noacc = '' or kodebarang like '" & myArray(1, 1) & "%' or kodebarang like '" & myArray(2, 1) & "%' or kodebarang like '" & myArray(3, 1) & "%' or kodebarang like '" & myArray(4, 1) & "%' or kodebarang like '" & myArray(5, 1) & "%' order by tglbeli,nobeli"
    If int1 = 0 Then SQL = "select * from am_manualposterima where noacc = '' order by tglbeli,nobeli"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If str9 <> RST!nobeli Then
            If str9 <> "" Then
                '-----------------jurnal b
                OBJ1.Open dsn
                SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str9 & "' group by nobeli"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai1 = RST1!nilai
                
                SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str9 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai2 = RST1!ppn
                OBJ1.Close
                
                If txtnilai2 <> 0 Then
                    txtnilai1 = (txtnilai1 * txtnilai2) / 100
                
                    grid2.TextMatrix(grid2.Row, 0) = Format(date3, "yyMM") & "/" & str8
                    grid2.TextMatrix(grid2.Row, 1) = str10
                    grid2.TextMatrix(grid2.Row, 2) = Format(txtnilai1, "###,###,##0.00")
                    grid2.TextMatrix(grid2.Row, 3) = "D"
                    grid2.TextMatrix(grid2.Row, 6) = Format(date3, "dd/MMM/yyyy")
                    grid2.TextMatrix(grid2.Row, 7) = Format(txtnilai3, "###,###,##0.00")
                    
                    OBJ1.Open dsn
                    SQL1 = "select top 1 ref2 from am_beliapp where nobeli = '" & str9 & "'"
                    Set RST1 = OBJ1.Execute(SQL1)
                    If Not RST1.EOF Then
                        If Len(RST1!ref2) > 50 Then
                            grid2.TextMatrix(grid2.Row, 5) = "PPn (" & Mid(RST1!ref2, 1, 50) & ")"
                        Else
                            grid2.TextMatrix(grid2.Row, 5) = "PPn (" & RST1!ref2 & ")"
                        End If
                    End If
                    OBJ1.Close
                    
                    grid2.Rows = grid2.Rows + 1
                    grid2.Row = grid2.Row + 1
                End If
                
                '------------jurnal c
                OBJ1.Open dsn
                SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str9 & "' group by nobeli"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai1 = RST1!nilai
                
                SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str9 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai2 = RST1!ppn
                OBJ1.Close
                
                If txtnilai2 <> 0 Then txtnilai1 = txtnilai1 * 1.1 Else txtnilai1 = txtnilai1 * 1
                
                grid2.TextMatrix(grid2.Row, 0) = Format(date3, "yyMM") & "/" & str8
                grid2.TextMatrix(grid2.Row, 1) = str10
                grid2.TextMatrix(grid2.Row, 2) = Format(txtnilai1, "###,###,##0.00")
                grid2.TextMatrix(grid2.Row, 3) = "K"
                grid2.TextMatrix(grid2.Row, 6) = Format(date3, "dd/MMM/yyyy")
                grid2.TextMatrix(grid2.Row, 7) = Format(txtnilai3, "###,###,##0.00")
                                
                int3 = Len(str2)
                
                OBJ1.Open dsn
                SQL1 = "select * from am_supplier where kodesupp = '" & str11 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then grid2.TextMatrix(grid2.Row, 5) = RST1!namasupp
                
                If Len(grid2.TextMatrix(grid2.Row, 5)) > (60 - (int3 + 2)) Then grid2.TextMatrix(grid2.Row, 5) = Mid(grid2.TextMatrix(grid2.Row, 5), 1, (60 - (int3 + 2)))
                
                grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row, 5) & "(" & str2 & ")"
                OBJ1.Close
                
                grid2.Rows = grid2.Rows + 1
                grid2.Row = grid2.Row + 1
            End If
            
            str9 = RST!nobeli
            
            str7 = str7 + 1
            If Len(str7) = 1 Then str8 = "00" & str7
            If Len(str7) = 2 Then str8 = "0" & str7
            If Len(str7) = 3 Then str8 = str7
        End If
        '---------jurnal a
        grid2.TextMatrix(grid2.Row, 0) = Format(RST!tglbeli, "yyMM") & "/" & str8
        grid2.TextMatrix(grid2.Row, 1) = RST!kodecur
        grid2.TextMatrix(grid2.Row, 2) = Format(RST!nilaibeli, "###,###,##0.00")
        grid2.TextMatrix(grid2.Row, 3) = "D"
        grid2.TextMatrix(grid2.Row, 6) = Format(RST!tglbeli, "dd/MMM/yyyy")
        grid2.TextMatrix(grid2.Row, 7) = Format(RST!nilaikurs, "###,###,##0.00")
        grid2.TextMatrix(grid2.Row, 8) = RST!nobeli
        
        OBJ1.Open dsn
        SQL1 = "select top 1 ref2 from am_beliapp where nobeli = '" & str9 & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then int3 = Len(RST1!ref2) Else int3 = 0
                
        SQL1 = "select * from am_apitemmst where kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid2.TextMatrix(grid2.Row, 5) = RST1!namabarang
        
        If Len(grid2.TextMatrix(grid2.Row, 5)) > (60 - (int3 + 2)) Then grid2.TextMatrix(grid2.Row, 5) = Mid(grid2.TextMatrix(grid2.Row, 5), 1, (60 - (int3 + 2)))
        
        SQL1 = "select top 1 ref2 from am_beliapp where nobeli = '" & str9 & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row, 5) & "(" & RST1!ref2 & ")"
            str2 = RST1!ref2
        Else
            str2 = ""
        End If
        OBJ1.Close
        
        date3 = RST!tglbeli
        str10 = RST!kodecur
        str11 = RST!kodesupp
        str12 = RST!nobeli
        txtnilai3 = RST!nilaikurs
        
        pro1.Value = pro1.Value + 1
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        RST.MoveNext
        
        If RST.EOF Then
            '------------------jurnal b
            OBJ1.Open dsn
            SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str12 & "' group by nobeli"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai1 = RST1!nilai
            
            SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str12 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai2 = RST1!ppn
            OBJ1.Close
            
            If txtnilai2 <> 0 Then
                txtnilai1 = (txtnilai1 * txtnilai2) / 100
            
                grid2.TextMatrix(grid2.Row, 0) = Format(date3, "yyMM") & "/" & str8
                grid2.TextMatrix(grid2.Row, 1) = str10
                grid2.TextMatrix(grid2.Row, 2) = Format(txtnilai1, "###,###,##0.00")
                grid2.TextMatrix(grid2.Row, 3) = "D"
                grid2.TextMatrix(grid2.Row, 6) = Format(date3, "dd/MMM/yyyy")
                grid2.TextMatrix(grid2.Row, 7) = Format(txtnilai3, "###,###,##0.00")
                                
                OBJ1.Open dsn
                SQL1 = "select top 1 ref2 from am_beliapp where nobeli = '" & str12 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    If Len(RST1!ref2) > 50 Then
                        grid2.TextMatrix(grid2.Row, 5) = "PPn (" & Mid(RST1!ref2, 1, 50) & ")"
                    Else
                        grid2.TextMatrix(grid2.Row, 5) = "PPn (" & RST1!ref2 & ")"
                    End If
                End If
                OBJ1.Close
                
                grid2.Rows = grid2.Rows + 1
                grid2.Row = grid2.Row + 1
            End If
            '------------jurnal c
            OBJ1.Open dsn
            SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str12 & "' group by nobeli"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai1 = RST1!nilai
            
            SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str12 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai2 = RST1!ppn
            OBJ1.Close
            
            If txtnilai2 <> 0 Then txtnilai1 = txtnilai1 * 1.1 Else txtnilai1 = txtnilai1 * 1
            
            grid2.TextMatrix(grid2.Row, 0) = Format(date3, "yyMM") & "/" & str8
            grid2.TextMatrix(grid2.Row, 1) = str10
            grid2.TextMatrix(grid2.Row, 2) = Format(txtnilai1, "###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 3) = "K"
            grid2.TextMatrix(grid2.Row, 6) = Format(date3, "dd/MMM/yyyy")
            grid2.TextMatrix(grid2.Row, 7) = Format(txtnilai3, "###,###,##0.00")
                        
            int3 = Len(str2)
            
            OBJ1.Open dsn
            SQL1 = "select * from am_supplier where kodesupp = '" & str11 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid2.TextMatrix(grid2.Row, 5) = RST1!namasupp
            
            If Len(grid2.TextMatrix(grid2.Row, 5)) > (60 - (int3 + 2)) Then grid2.TextMatrix(grid2.Row, 5) = Mid(grid2.TextMatrix(grid2.Row, 5), 1, (60 - (int3 + 2)))
            
            grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row, 5) & "(" & str2 & ")"
            OBJ1.Close
            
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
        End If
    Loop
    OBJ.Close
    
    MsgBox "Manual Jurnal Complete.", vbInformation, "Information"
    cmdproses1.Enabled = True
    pro1.Visible = False
    pro1.Value = 0
End Sub

Private Sub cmdproses1_Click()
    If date1 > date2 Then
        MsgBox "Invalid date range, posting aborted.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If txtkode1 = "" Or txtkodecomp = "" Then
        MsgBox "Data entry not complete.????", vbExclamation, "Warning"
        Exit Sub
    End If
    
    int5 = 0
    OBJ.Open dsn
    
    'Filtering Kode Barang
    SQL = "select distinct substring(kodebarang,1,3) 'a' from am_beliapp order by a"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        int5 = int5 + 1
        RST.MoveNext
    Loop
    
    SQL = "select count(kdkurs)'qq' from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal a'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then int6 = RST!qq Else int6 = 0
    OBJ.Close
    
    If int5 <> int6 Then
        MsgBox "Please recheck Jurnal a, proses aborted.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Hutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Hutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    int4 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        If grid2.Row <> 1 Then
            If grid2.TextMatrix(grid2.Row - 1, 0) = grid2.TextMatrix(grid2.Row, 0) Then
                If (grid2.TextMatrix(grid2.Row - 1, 4) = "" And grid2.TextMatrix(grid2.Row, 4) <> "") Or (grid2.TextMatrix(grid2.Row - 1, 4) <> "" And grid2.TextMatrix(grid2.Row, 4) = "") Then
                    MsgBox "Data entry not complete on " & grid2.TextMatrix(grid2.Row, 5), vbExclamation, "Warning"
                    Exit Sub
                End If
            End If
        End If
        
        If grid2.TextMatrix(grid2.Row, 4) = "" Then int4 = int4 + 1
        
        grid2.Row = grid2.Row + 1
    Loop
    
    If int4 = grid2.Rows - 2 Then
        If MsgBox("Manual jurnal tidak ditemukan. " & Chr(13) & "Apakah anda akan melanjutkan proses jurnal hutang", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    End If
    
    '====================================================================================
    pro1.Visible = True
    pro1.Max = grid2.Rows - 2
    pro1.Value = 0
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        
        If grid2.Row = 1 Then
            int2 = 1
        Else
            If grid2.TextMatrix(grid2.Row, 0) <> grid2.TextMatrix(grid2.Row - 1, 0) Then int2 = 1
        End If
        
        If grid2.TextMatrix(grid2.Row, 4) <> "" Then
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
            SQL1 = SQL1 + "convert(datetime,'" & Month(grid2.TextMatrix(grid2.Row, 6)) & "/" & Day(grid2.TextMatrix(grid2.Row, 6)) & "/" & Year(grid2.TextMatrix(grid2.Row, 6)) & "'),"
            SQL1 = SQL1 + "'" & txtkode1 & "',"
            SQL1 = SQL1 + "'" & grid2.TextMatrix(grid2.Row, 0) & "',"
            SQL1 = SQL1 + "convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 7), "general number") & "'),"
            SQL1 = SQL1 + "'" & grid2.TextMatrix(grid2.Row, 4) & "',"
            SQL1 = SQL1 + "'" & grid2.TextMatrix(grid2.Row, 5) & "',"
            SQL1 = SQL1 + "'" & grid2.TextMatrix(grid2.Row, 3) & "',"
            SQL1 = SQL1 + "convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") * Format(grid2.TextMatrix(grid2.Row, 7), "general number") & "'),"
            SQL1 = SQL1 + "convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") & "'),"
            SQL1 = SQL1 + "'" & grid2.TextMatrix(grid2.Row, 1) & "',"
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
            
            'nanti masa percobaan
            SQL1 = "update am_beliapp set flag2 = '2' where nobeli = '" & grid2.TextMatrix(grid2.Row, 8) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        
        pro1.Value = pro1.Value + 1
        int2 = int2 + 1
        grid2.Row = grid2.Row + 1
    Loop
    '==================================================================================================
    
    OBJ.Open dsn
    SQL = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode1 & "' and notrx like '" & Format(date1, "YYMM/") & "%' order by notrx desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str7 = Right(RST!notrx, 3) Else str7 = 0
    OBJ.Close
    
    str9 = ""
    pro1.Value = 0
    
    OBJ.Open dsn
    If int1 <> 0 Then SQL = "select count(nobeli)'hitnobeli' from am_manualposterima where noacc <> '' and kodebarang not like '" & myArray(1, 1) & "%' and kodebarang not like '" & myArray(2, 1) & "%' and kodebarang not like '" & myArray(3, 1) & "%' and kodebarang not like '" & myArray(4, 1) & "%' and kodebarang not like '" & myArray(5, 1) & "%'"
    If int1 = 0 Then SQL = "select count(nobeli)'hitnobeli' from am_manualposterima where noacc <> ''"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then pro1.Max = RST!hitnobeli Else pro1.Max = 0
    If pro1.Max = 0 Then pro1.Visible = False
    
    If int1 <> 0 Then SQL = "select * from am_manualposterima where noacc <> '' and kodebarang not like '" & myArray(1, 1) & "%' and kodebarang not like '" & myArray(2, 1) & "%' and kodebarang not like '" & myArray(3, 1) & "%' and kodebarang not like '" & myArray(4, 1) & "%' and kodebarang not like '" & myArray(5, 1) & "%' order by nobeli"
    If int1 = 0 Then SQL = "select * from am_manualposterima where noacc <> '' order by nobeli"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If str9 = "" Then int2 = 1
        
        If str9 <> RST!nobeli Then
            If str9 <> "" Then
                OBJ1.Open dsn
                SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str9 & "' group by nobeli"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai1 = RST1!nilai
                
                SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str9 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai2 = RST1!ppn
                OBJ1.Close
                
                If txtnilai2 <> 0 Then
                    txtnilai1 = (txtnilai1 * txtnilai2) / 100
                    
                    'jurnal b
                    OBJ1.Open dsn
                    SQL1 = "select noacc,dk from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal b'"
                    Set RST1 = OBJ1.Execute(SQL1)
                    If Not RST1.EOF Then
                        str3 = RST1!noacc
                        str5 = RST1!dk
                    End If
                    
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
                    SQL1 = SQL1 + "'" & Format(RST!tglbeli, "YYMM") & "/" & str8 & "',"
                    SQL1 = SQL1 + "convert(money,'" & txtnilai3 & "'),"
                    SQL1 = SQL1 + "'" & str3 & "',"
                    
                    OBJ2.Open dsn
                    SQL2 = "select top 1 ref2 from am_beliapp where nobeli = '" & str9 & "'"
                    Set RST2 = OBJ2.Execute(SQL2)
                    If Not RST2.EOF Then
                        If Len(RST2!ref2) > 50 Then
                            SQL1 = SQL1 + "'PPn (" & Mid(RST2!ref2, 1, 50) & ")',"
                        Else
                            SQL1 = SQL1 + "'PPn (" & RST2!ref2 & ")',"
                        End If
                    End If
                    OBJ2.Close
                    
                    SQL1 = SQL1 + "'" & str5 & "',"
                    SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
                    SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
                    SQL1 = SQL1 + "'" & str10 & "',"
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
                '------------
                OBJ1.Open dsn
                SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str9 & "' group by nobeli"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai1 = RST1!nilai
                
                SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str9 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then txtnilai2 = RST1!ppn
                OBJ1.Close
                
                If txtnilai2 <> 0 Then txtnilai1 = txtnilai1 * 1.1 Else txtnilai1 = txtnilai1 * 1
                
                'jurnal c
                OBJ1.Open dsn
                SQL1 = "select noacc,dk from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal c'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    str3 = RST1!noacc
                    str5 = RST1!dk
                End If
                
                SQL1 = "select noacc from am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & str11 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc
                
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
                SQL1 = SQL1 + "'" & Format(RST!tglbeli, "YYMM") & "/" & str8 & "',"
                SQL1 = SQL1 + "convert(money,'" & txtnilai3 & "'),"
                SQL1 = SQL1 + "'" & str3 & "',"
                                
                int3 = Len(str2)
                
                OBJ2.Open dsn
                SQL2 = "select * from am_supplier where kodesupp = '" & str11 & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then str13 = RST2!namasupp
                
                If Len(str13) > (60 - (int3 + 2)) Then str13 = Mid(str13, 1, (60 - (int3 + 2)))
                
                str13 = str13 & "(" & str2 & ")"
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str13 & "',"
                SQL1 = SQL1 + "'" & str5 & "',"
                SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
                SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
                SQL1 = SQL1 + "'" & str10 & "',"
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
                
                '############ update confirman
                OBJ1.Open dsn
                SQL1 = "update am_beliapp set flag2 = '2' where nobeli = '" & str9 & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
                
                int2 = 1
            End If
            
            str9 = RST!nobeli
            
            str7 = str7 + 1
            If Len(str7) = 1 Then str8 = "00" & str7
            If Len(str7) = 2 Then str8 = "0" & str7
            If Len(str7) = 3 Then str8 = str7
        End If
        
        'jurnal a
        OBJ1.Open dsn
        SQL1 = "select noacc,dk from am_auto where kodecomp = '" & txtkodecomp & "' and kdkurs = '" & Left(RST!kodebarang, 3) & "' and jurnal_ = 'Jurnal a'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            str3 = RST1!noacc
            str5 = RST1!dk
        End If
        
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
        SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tglbeli) & "/" & Day(RST!tglbeli) & "/" & Year(RST!tglbeli) & "'),"
        SQL1 = SQL1 + "'" & txtkode1 & "',"
        SQL1 = SQL1 + "'" & Format(RST!tglbeli, "YYMM") & "/" & str8 & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaikurs & "'),"
        SQL1 = SQL1 + "'" & str3 & "',"
        
        OBJ2.Open dsn
        SQL2 = "select top 1 ref2 from am_beliapp where nobeli = '" & RST!nobeli & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then int3 = Len(RST2!ref2) Else int3 = 0
        
        SQL2 = "select * from am_apitemmst where kodebarang = '" & RST!kodebarang & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then str13 = RST2!namabarang
        
        If Len(str13) > (60 - (int3 + 2)) Then str13 = Mid(str13, 1, (60 - (int3 + 2)))
        
        SQL2 = "select top 1 ref2 from am_beliapp where nobeli = '" & RST!nobeli & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            str13 = str13 & "(" & RST2!ref2 & ")"
            str2 = RST2!ref2
        Else
            str2 = ""
        End If
        OBJ2.Close
        
        SQL1 = SQL1 + "'" & str13 & "',"
        SQL1 = SQL1 + "'" & str5 & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaibeli * RST!nilaikurs & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaibeli & "'),"
        SQL1 = SQL1 + "'" & RST!kodecur & "',"
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
        
        date3 = RST!tglbeli
        str10 = RST!kodecur
        str11 = RST!kodesupp
        str12 = RST!nobeli
        txtnilai3 = RST!nilaikurs
        
        pro1.Value = pro1.Value + 1
        int2 = int2 + 1
        RST.MoveNext
        
        If RST.EOF Then
            OBJ1.Open dsn
            SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str12 & "' group by nobeli"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai1 = RST1!nilai
            
            SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str12 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai2 = RST1!ppn
            OBJ1.Close
            
            If txtnilai2 <> 0 Then
                txtnilai1 = (txtnilai1 * txtnilai2) / 100
                
                'jurnal b
                OBJ1.Open dsn
                SQL1 = "select noacc,dk from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal b'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    str3 = RST1!noacc
                    str5 = RST1!dk
                End If
                
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
                SQL1 = SQL1 + "convert(money,'" & txtnilai3 & "'),"
                SQL1 = SQL1 + "'" & str3 & "',"
                
                OBJ2.Open dsn
                SQL2 = "select top 1 ref2 from am_beliapp where nobeli = '" & str12 & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    If Len(RST2!ref2) > 50 Then
                        SQL1 = SQL1 + "'PPn (" & Mid(RST2!ref2, 1, 50) & ")',"
                    Else
                        SQL1 = SQL1 + "'PPn (" & RST2!ref2 & ")',"
                    End If
                End If
                OBJ2.Close
                
                SQL1 = SQL1 + "'" & str5 & "',"
                SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
                SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
                SQL1 = SQL1 + "'" & str10 & "',"
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
            '------------
            OBJ1.Open dsn
            SQL1 = "select sum(nilaibeli)'nilai',nobeli from am_manualposterima where nobeli = '" & str12 & "' group by nobeli"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai1 = RST1!nilai
            
            SQL1 = "select top 1 ppn from am_manualposterima where nobeli = '" & str12 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai2 = RST1!ppn
            OBJ1.Close
            
            If txtnilai2 <> 0 Then txtnilai1 = txtnilai1 * 1.1 Else txtnilai1 = txtnilai1 * 1
            
            'jurnal c
            OBJ1.Open dsn
            SQL1 = "select noacc,dk from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal c'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                str3 = RST1!noacc
                str5 = RST1!dk
            End If
            
            SQL1 = "select noacc from am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & str11 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc
            
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
            SQL1 = SQL1 + "convert(money,'" & txtnilai3 & "'),"
            SQL1 = SQL1 + "'" & str3 & "',"
                        
            int3 = Len(str2)
            
            OBJ2.Open dsn
            SQL2 = "select * from am_supplier where kodesupp = '" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then str13 = RST2!namasupp
            
            If Len(str13) > (60 - (int3 + 2)) Then str13 = Mid(str13, 1, (60 - (int3 + 2)))
            
            str13 = str13 & "(" & str2 & ")"
            OBJ2.Close
            
            SQL1 = SQL1 + "'" & str13 & "',"
            SQL1 = SQL1 + "'" & str5 & "',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 * txtnilai3 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "'" & str10 & "',"
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
            
            '#################################### update confirm #########################################
            
            OBJ1.Open dsn
            SQL1 = "update am_beliapp set flag2 = '2' where nobeli = '" & str12 & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
    Loop
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
    
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Do
        
        If Grid1.TextMatrix(Grid1.Row, 2) = "" Then
            MsgBox "Data entry not complete on grid.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        Grid1.Row = Grid1.Row + 1
    Loop
    
    int5 = 0
    OBJ.Open dsn
    SQL = "select distinct substring(kodebarang,1,3)'a' from am_beliapp order by a"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        int5 = int5 + 1
        RST.MoveNext
    Loop

    SQL = "select count(kdkurs)'qq' from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal a'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then int6 = RST!qq Else int6 = 0
    OBJ.Close

    If int5 <> int6 Then
        MsgBox "Please recheck Jurnal a, proses aborted.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Bayar Hutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Bayar Hutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    'ngecek status giro
    OBJ0.Open dsn
    SQL0 = "select year(tanggalcair)'cair',year(tanggaltolak)'tolak',nogiro from am_apcashsub where tglbukti>='" & tanggal4 & "' and tglbukti<='" & tanggal5 & "' and type ='G'"
    
    Set RST0 = OBJ0.Execute(SQL0)
    
    Do While Not RST0.EOF
        If RST0!cair = 1900 And RST0!tolak = 1900 Then
            MsgBox "Status Giro " & RST0!nogiro & " belum terdefinisi, proses di batalkan.", vbExclamation, "Warning"
            OBJ0.Close
            Exit Sub
        End If
        
        RST0.MoveNext
    Loop
    OBJ0.Close
    
    
    'udah
    OBJ.Open dsn
    SQL = "SELECT count(NoBkt)'hitnobkt' FROM AM_apcashhdr WHERE kodebayar='PM' and"
    SQL = SQL + " tglbkt>='" & tanggal4 & "' and tglbkt<='" & tanggal5 & "' and posted='0'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then pro2.Max = RST!hitnobkt Else pro2.Max = 0
    If pro2.Max = 0 Then pro2.Visible = False Else pro2.Visible = True
    pro2.Value = 0
    OBJ.Close
    
    '===========================================================================================================
    OBJ.Open dsn
    SQL = "SELECT a.Kodesupp, a.NoBkt, a.TglBkt,b.TglJT,b.type, a.Amount, a.kodecur, a.nilaikurs, c.base"
    SQL = SQL + " FROM AM_apcashhdr a left join gl_kurs c"
    SQL = SQL + " ON a.kodecur=c.kdkurs"
    SQL = SQL + " left join am_apcashsub b on a.NoBkt =b.nobukti"
    SQL = SQL + " WHERE a.kodebayar='PM' and a.tglbkt>='" & tanggal4 & "' and a.tglbkt<='" & tanggal5 & "'"
    SQL = SQL + " and a.posted='0' order by a.tglbkt,a.nobkt"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str14 = ""
        OBJ0.Open dsn
        SQL0 = "select noapply from am_apcashlin where nobkt = '" & RST!nobkt & "' and kodebayar = 'PM'"
        Set RST0 = OBJ0.Execute(SQL0)
        Do While Not RST0.EOF
            str14 = str14 + RST0!noapply + "; "
            RST0.MoveNext
        Loop
        OBJ0.Close
        
        OBJ2.Open dsn
        txtnilai2 = 0
      
        SQL2 = "select sum(selisih * -1)'selisi' from am_apcashlin where nobkt = '" & RST!nobkt & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then txtnilai2 = RST2!selisi
        
        txtnilai5 = 0
        str6 = ""
        SQL2 = "select noapply from am_apopnfil where nobeli = '" & RST!nobkt & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then str6 = RST2!noapply
        
        SQL2 = "select * from am_apopnfil where noapply = '" & str6 & "' and transtype='I'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then
            SQL2 = "select sum(amount * nilaikurs)'utang' from am_apopnfil where noapply = '" & str6 & "' and transtype='I'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai5 = RST2!utang
            
            SQL2 = "select nilaikurs from am_apopnfil where noapply = '" & str6 & "' and transtype='I'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai7 = RST2!nilaikurs
        Else
            SQL2 = "select sum(amount * nilaikurs)'utang' from am_apopnfil where noapply = '" & str6 & "' and (transtype='CN' or transtype='DN')"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai5 = Format(RST2!utang, "###,###,##0.00")
            
            SQL2 = "select nilaikurs from am_apopnfil where noapply = '" & str6 & "' and (transtype='CN' or transtype='DN')"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai7 = RST2!nilaikurs
        End If
        OBJ2.Close
    'Test
    If RST!Type = "TN" Then
        date3 = RST!tglbkt
    Else
        date3 = RST!tgljt
    End If
    
        txtnilai3 = RST!nilaikurs
        str10 = RST!kodecur
        str11 = RST!kodesupp
        If str10 <> str1 Then txtnilai1 = txtnilai5 Else txtnilai1 = RST!amount * RST!nilaikurs
        
        'ngecek pembayaran pake giro apa bukan
        OBJ0.Open dsn
        SQL0 = "select count(Type)'q' from am_apcashsub where nobukti = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        If Not RST0.EOF Then
            If RST0!q = 1 Then
                SQL0 = "select nogiro,type,tanggalcair,tanggaltolak,year(tanggalcair)'cair',year(tanggaltolak)'tolak' from am_apcashsub where nobukti = '" & RST!nobkt & "'"
                Set RST0 = OBJ0.Execute(SQL0)
                If Not RST0.EOF Then
                    If RST0!Type = "G" Then
                        If RST0!cair <> 1900 Then
                            date3 = RST0!tanggalcair
                            str4 = RST0!nogiro
                        End If
                        If RST0!tolak <> 1900 Then
                            OBJ0.Close
                            GoTo lompat
                        End If
                    Else
                        str4 = ""
                    End If
                End If
            End If
        End If
        OBJ0.Close
        'udah
        
        'Jurnal hutang supplier start
        OBJ1.Open dsn
        SQL1 = "select * from am_autoaccsupp where kodesupp = '" & str11 & "' and kodecomp = '" & txtkodecomp & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            If RST1!noacc = "" Then
                Grid1.Row = 1
                Do While True
                    If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Do
                    
                    If Grid1.TextMatrix(Grid1.Row, 0) = str11 Then
                        str3 = Grid1.TextMatrix(Grid1.Row, 2)
                        Exit Do
                    End If
                    
                    Grid1.Row = Grid1.Row + 1
                Loop
            Else
                str3 = RST1!noacc
            End If
        Else
            Grid1.Row = 1
            Do While True
                If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Do
                
                If Grid1.TextMatrix(Grid1.Row, 0) = str11 Then
                    str3 = Grid1.TextMatrix(Grid1.Row, 2)
                    Exit Do
                End If
                
                Grid1.Row = Grid1.Row + 1
            Loop
        End If
        
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
        SQL1 = SQL1 + "'" & txtkode2 & "',"
        SQL1 = SQL1 + "'" & RST!nobkt & "',"
        SQL1 = SQL1 + "convert(money,'1'),"
        SQL1 = SQL1 + "'" & str3 & "',"
        
        If str4 <> "" Then str14 = "(" & str4 & ") " & str14
        If Len(str14) > 60 Then str14 = Mid(str14, 1, 60)
        
        str15 = ""
        OBJ2.Open dsn
        SQL2 = "select namasupp from am_supplier where kodesupp = '" & str11 & "'"
        Set RST2 = OBJ2.Execute(SQL2)
        If Not RST2.EOF Then str15 = RST2!namasupp
        OBJ2.Close
        str15 = str14 + str15
        
        If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
        
        SQL1 = SQL1 + "'" & str15 & "',"
        SQL1 = SQL1 + "'D',"
        SQL1 = SQL1 + "convert(money,'" & txtnilai1 + txtnilai2 & "'),"
        SQL1 = SQL1 + "convert(money,'" & txtnilai1 + txtnilai2 & "'),"
        SQL1 = SQL1 + "'" & str1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "'auto',"
        SQL1 = SQL1 + "'',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(numeric,'1'))"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        'Jurnal hutang supplier end ok
        int2 = 2
        
        'jurnal g (biaya admin) start
        OBJ0.Open dsn
        SQL0 = "select isnull(sum(byadmin),0)'biaya' from am_apcashsub where nobukti = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        If Not RST0.EOF Then txtnilai4 = RST0!biaya
        OBJ0.Close
        If txtnilai4 > 0 Then
            OBJ0.Open dsn
            SQL0 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal g'"
            Set RST0 = OBJ0.Execute(SQL0)
            If Not RST0.EOF Then str3 = RST0!noacc Else str3 = ""
            OBJ0.Close
            
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
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            str15 = "Biaya Administrasi"
            str15 = str14 + str15
        
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
            
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai4 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai4 & "'),"
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
        'jurnal g (biaya admin) end ok
        
        OBJ0.Open dsn
        SQL0 = "select selisihkurs,potongan,selisih,noapply from am_apcashlin where nobkt = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        Do While Not RST0.EOF
            'jurnal d (selisih kurs) start
            If RST0!selisihkurs > 0 Then
                OBJ1.Open dsn
                SQL1 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal d'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                
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
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"  'selisih kurs
                SQL1 = SQL1 + "'Selisih Kurs " & RST0!noapply & "',"
                SQL1 = SQL1 + "'K'," 'debet atau kredit
                SQL1 = SQL1 + "convert(money,'" & RST0!selisihkurs & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!selisihkurs & "'),"
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
            ElseIf RST0!selisihkurs < 0 Then
                OBJ1.Open dsn
                SQL1 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal d'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                
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
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"  'selisih kurs
                SQL1 = SQL1 + "'Selisih Kurs " & RST0!noapply & "',"
                SQL1 = SQL1 + "'D'," 'debet atau kredit
                SQL1 = SQL1 + "convert(money,'" & RST0!selisihkurs * -1 & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!selisihkurs * -1 & "'),"
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
            'jurnal d (selisih kurs) end ok
            
            'jurnal e (potongan) start
            If RST0!potongan > 0 Then
                OBJ1.Open dsn
                SQL1 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal e'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                
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
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"  'potongan
                SQL1 = SQL1 + "'Potongan Bayar " & RST0!noapply & "',"
                SQL1 = SQL1 + "'K'," 'debet atau kredit
                SQL1 = SQL1 + "convert(money,'" & RST0!potongan & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!potongan & "'),"
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
            'jurnal e (potongan) end ok
            
            'jurnal f (selisih bayar) start
            If RST0!selisih > 0 Then
                OBJ1.Open dsn
                SQL1 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal f'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                
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
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"  'selisih bayar
                SQL1 = SQL1 + "'Selisih Bayar " & RST0!noapply & "',"
                SQL1 = SQL1 + "'D'," 'debet atau kredit
                SQL1 = SQL1 + "convert(money,'" & RST0!selisih & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!selisih & "'),"
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
                
            ElseIf RST0!selisih < 0 Then
                
                OBJ1.Open dsn
                SQL1 = "select noacc from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal f'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                
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
                SQL1 = SQL1 + "'" & txtkode2 & "',"
                SQL1 = SQL1 + "'" & RST!nobkt & "',"
                SQL1 = SQL1 + "convert(money,'1'),"
                SQL1 = SQL1 + "'" & str3 & "',"  'selisih bayar
                SQL1 = SQL1 + "'Selisih bayar " & RST0!noapply & "',"
                SQL1 = SQL1 + "'K'," 'debet atau kredit
                SQL1 = SQL1 + "convert(money,'" & RST0!selisih * -1 & "'),"
                SQL1 = SQL1 + "convert(money,'" & RST0!selisih * -1 & "'),"
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
            'jurnal f (selisih bayar) end ok
            
            RST0.MoveNext
        Loop
        OBJ0.Close
        
        'jurnal kas/bank start
        OBJ0.Open dsn
        SQL0 = "select Type,Bank,Byadmin,Jumlah from am_apcashsub where nobukti = '" & RST!nobkt & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        Do While Not RST0.EOF
            If RST0!Type = "TN" Then
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoaccbank where kodecomp = '" & txtkodecomp & "' and kodebank = '" & RST!kodecur & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                OBJ1.Close
            Else 'TF/C/G/TN
                OBJ1.Open dsn
                SQL1 = "select noacc from am_autoaccbank where kodecomp = '" & txtkodecomp & "' and kodebank = '" & RST0!bank & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
                OBJ1.Close
            End If
                
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
            SQL1 = SQL1 + "'" & txtkode2 & "',"
            SQL1 = SQL1 + "'" & RST!nobkt & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            If str4 <> "" Then str15 = "(" & str4 & ") " & "Kas/Bank; " Else str15 = "Kas/Bank; "
            str15 = str14 + str15
        
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
            
            OBJ2.Open dsn
            SQL2 = "select namasupp from am_supplier where kodesupp = '" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then str15 = str15 + RST2!namasupp
            OBJ2.Close
            
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
            
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * txtnilai3 & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST0!jumlah * txtnilai3 & "'),"
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
            RST0.MoveNext
        Loop
        OBJ0.Close
        'jurnal kas/bank end ok
        
        If Asc(Right(RST!nobkt, 1)) >= 65 And Asc(Right(RST!nobkt, 1)) <= 90 Then
            'update yg pertama jika cicil
            OBJ1.Open dsn
            SQL1 = "select sum(Jumlah)'jum' from am_apcashsub where nobukti = '" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then txtnilai6 = RST1!jum Else txtnilai6 = 0
            
            SQL1 = "update gl_transaksi set amounttrx = " & (txtnilai6 * txtnilai7) & " , nilaitrx = " & txtnilai6 & " where kdtrx = '" & txtkode2 & "' and notrx = '" & Format(date3, "YYMM") & "/" & RST!nobkt & "' and lineitem=1"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        
        
        'NANTI DI AKTIVKAN
        
        OBJ1.Open dsn
        SQL1 = "update am_apcashhdr set posted = '1' where nobkt = '" & RST!nobkt & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
lompat:
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
    
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "User harus memilih proses debit note atau credit note.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid3.Row = 1
    Do While True
        If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
        
        If grid3.TextMatrix(grid3.Row, 2) = "" Then
            MsgBox "Data entry not complete on grid.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid3.Row = grid3.Row + 1
    Loop
    
    int5 = 0
    OBJ.Open dsn
    SQL = "select distinct substring(kodebarang,1,3)'a' from am_beliapp order by a"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        int5 = int5 + 1
        
        RST.MoveNext
    Loop
    
    SQL = "select count(kdkurs)'qq' from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal a'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then int6 = RST!qq Else int6 = 0
    OBJ.Close
    
    If int5 <> int6 Then
        MsgBox "Please recheck Jurnal a, proses aborted.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Jurnal Koreksi Hutang berlangsung." & vbCrLf & _
    "Lanjutkan Proses Jurnal Koreksi Hutang ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    If Option1.Value Then
        SQL = "SELECT count(NoBkt)'hitnobkt' FROM AM_apcashhdr WHERE kodebayar='CN' and"
        SQL = SQL + " tglbkt>='" & tanggal6 & "' and tglbkt<='" & tanggal7 & "' and posted='0'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then pro3.Max = RST!hitnobkt Else pro3.Max = 0
        If pro3.Max = 0 Then pro3.Visible = False Else pro3.Visible = True
    
        SQL = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode3 & "' and notrx like '" & Format(date6, "YYMM/") & "%' order by notrx desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then str7 = Right(RST!notrx, 3) Else str7 = 0
    ElseIf Option2.Value Then
        SQL = "SELECT count(NoBkt)'hitnobkt' FROM AM_apcashhdr WHERE kodebayar='DN' and"
        SQL = SQL + " tglbkt>='" & tanggal6 & "' and tglbkt<='" & tanggal7 & "' and posted='0'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then pro3.Max = RST!hitnobkt Else pro3.Max = 0
        If pro3.Max = 0 Then pro3.Visible = False Else pro3.Visible = True
    
        SQL = "select top 1 notrx from gl_transaksi where kdcomp = '" & txtkodecomp & "' and kdtrx = '" & txtkode3 & "' and notrx like '" & Format(date6, "YYMM/") & "%' order by notrx desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then str7 = Right(RST!notrx, 3) Else str7 = 0
    End If
    OBJ.Close
    
    pro3.Value = 0
    str7 = str7 + 1
    If Len(str7) = 1 Then str8 = "00" & str7
    If Len(str7) = 2 Then str8 = "0" & str7
    If Len(str7) = 3 Then str8 = str7
    
    '====================================================================================
    If Option1.Value Then
        OBJ.Open dsn
        SQL = "SELECT a.Kodesupp, a.NoBkt, a.TglBkt, a.Amount, a.kodecur, a.nilaikurs, a.idupdate, c.base"
        SQL = SQL + " FROM AM_apcashhdr a left join gl_kurs c"
        SQL = SQL + " ON a.kodecur=c.kdkurs"
        SQL = SQL + " WHERE a.kodebayar='CN' and a.tglbkt>='" & tanggal6 & "' and a.tglbkt<='" & tanggal7 & "' and a.posted='0' order by a.tglbkt,a.nobkt"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            str14 = ""
            OBJ0.Open dsn
            SQL0 = "select noapply from am_apcashlin where nobkt = '" & RST!nobkt & "' and kodebayar = 'CN'"
            Set RST0 = OBJ0.Execute(SQL0)
            If Not RST0.EOF Then str14 = RST0!noapply
            OBJ0.Close
            
            OBJ2.Open dsn
            txtnilai2 = 0
            SQL2 = "select jumlah from am_apcashlin where nobkt = '" & RST!nobkt & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai2 = RST2!jumlah
            
            txtnilai5 = 0
            SQL2 = "select koreksippn from am_apcashlinppn where nobkt = '" & RST!nobkt & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai5 = RST2!koreksippn
            OBJ2.Close
            
            date3 = RST!tglbkt
            txtnilai3 = RST!nilaikurs
            str10 = RST!kodecur
            str11 = RST!kodesupp
            
            If str10 <> str1 Then txtnilai1 = txtnilai2 + txtnilai5 Else txtnilai1 = (txtnilai2 + txtnilai5) * RST!nilaikurs
                
            OBJ1.Open dsn
            SQL1 = "select * from am_autoaccsupp where kodesupp = '" & str11 & "' and kodecomp = '" & txtkodecomp & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                If RST1!noacc = "" Then
                    grid3.Row = 1
                    Do While True
                        If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
                        
                        If grid3.TextMatrix(grid3.Row, 0) = str11 Then
                            str3 = grid3.TextMatrix(grid3.Row, 2)
                            Exit Do
                        End If
                        
                        grid3.Row = grid3.Row + 1
                    Loop
                Else
                    str3 = RST1!noacc
                End If
            Else
                grid3.Row = 1
                Do While True
                    If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
                    
                    If grid3.TextMatrix(grid3.Row, 0) = str11 Then
                        str3 = grid3.TextMatrix(grid3.Row, 2)
                        Exit Do
                    End If
                    
                    grid3.Row = grid3.Row + 1
                Loop
            End If
            'debet
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
            SQL1 = SQL1 + "'" & txtkode3 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & RST!idupdate & "',"
            
            If Len(str14) > 60 Then str14 = Mid(str14, 1, 60)
        
            str15 = ""
            OBJ2.Open dsn
            SQL2 = "select namasupp from am_supplier where kodesupp = '" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then str15 = RST2!namasupp
            OBJ2.Close
            str15 = str14 + " " + str15
            
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
                
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            'kredit
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
            SQL1 = SQL1 + "'" & txtkode3 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            str15 = " Koreksi"
            str15 = str14 + str15
        
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
            
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "update am_apcashhdr set posted = '1' where nobkt = '" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
                        
            str7 = str7 + 1
            If Len(str7) = 1 Then str8 = "00" & str7
            If Len(str7) = 2 Then str8 = "0" & str7
            If Len(str7) = 3 Then str8 = str7
            pro3.Value = pro3.Value + 1
            
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    
    If Option2.Value Then
        OBJ.Open dsn
        SQL = "SELECT a.Kodesupp, a.NoBkt, a.TglBkt, a.Amount, a.kodecur, a.nilaikurs, a.idupdate, c.base"
        SQL = SQL + " FROM AM_apcashhdr a left join gl_kurs c"
        SQL = SQL + " ON a.kodecur=c.kdkurs"
        SQL = SQL + " WHERE a.kodebayar='DN' and a.tglbkt>='" & tanggal6 & "' and a.tglbkt<='" & tanggal7 & "' and a.posted='0' order by a.tglbkt,a.nobkt"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            str14 = ""
            OBJ0.Open dsn
            SQL0 = "select noapply from am_apcashlin where nobkt = '" & RST!nobkt & "' and kodebayar = 'DN'"
            Set RST0 = OBJ0.Execute(SQL0)
            If Not RST0.EOF Then str14 = RST0!noapply
            OBJ0.Close
            
            OBJ2.Open dsn
            txtnilai2 = 0
            SQL2 = "select jumlah from am_apcashlin where nobkt = '" & RST!nobkt & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai2 = RST2!jumlah
            
            txtnilai5 = 0
            SQL2 = "select koreksippn from am_apcashlinppn where nobkt = '" & RST!nobkt & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then txtnilai5 = RST2!koreksippn
            OBJ2.Close
            
            date3 = RST!tglbkt
            txtnilai3 = RST!nilaikurs
            str10 = RST!kodecur
            str11 = RST!kodesupp
            
            If str10 <> str1 Then txtnilai1 = txtnilai2 + txtnilai5 Else txtnilai1 = (txtnilai2 + txtnilai5) * RST!nilaikurs
                
            OBJ1.Open dsn
            SQL1 = "select * from am_autoaccsupp where kodesupp = '" & str11 & "' and kodecomp = '" & txtkodecomp & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                If RST1!noacc = "" Then
                    grid3.Row = 1
                    Do While True
                        If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
                        
                        If grid3.TextMatrix(grid3.Row, 0) = str11 Then
                            str3 = grid3.TextMatrix(grid3.Row, 2)
                            Exit Do
                        End If
                        
                        grid3.Row = grid3.Row + 1
                    Loop
                Else
                    str3 = RST1!noacc
                End If
            Else
                grid3.Row = 1
                Do While True
                    If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
                    
                    If grid3.TextMatrix(grid3.Row, 0) = str11 Then
                        str3 = grid3.TextMatrix(grid3.Row, 2)
                        Exit Do
                    End If
                    
                    grid3.Row = grid3.Row + 1
                Loop
            End If
            
            'debet
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
            SQL1 = SQL1 + "'" & txtkode3 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            
            If Len(str14) > 60 Then str14 = Mid(str14, 1, 60)
        
            str15 = ""
            OBJ2.Open dsn
            SQL2 = "select namasupp from am_supplier where kodesupp = '" & str11 & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If Not RST2.EOF Then str15 = RST2!namasupp
            OBJ2.Close
            str15 = str14 + " " + str15
            
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
                
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            'kredit
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
            SQL1 = SQL1 + "'" & txtkode3 & "',"
            SQL1 = SQL1 + "'" & Format(date3, "YYMM") & "/" & str8 & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & RST!idupdate & "',"
            
            str15 = " Koreksi"
            str15 = str14 + str15
        
            If Len(str15) > 60 Then str15 = Mid(str15, 1, 60)
            
            SQL1 = SQL1 + "'" & str15 & "',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "convert(money,'" & txtnilai1 & "'),"
            SQL1 = SQL1 + "'" & str1 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "update am_apcashhdr set posted = '1' where nobkt = '" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
                        
            str7 = str7 + 1
            If Len(str7) = 1 Then str8 = "00" & str7
            If Len(str7) = 2 Then str8 = "0" & str7
            If Len(str7) = 3 Then str8 = str7
            pro3.Value = pro3.Value + 1
            
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    
    MsgBox "Proses Complete.", vbInformation, "Information"
    pro3.Visible = False
    pro3.Value = 0
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
    txtacc4.SetFocus
End Sub

Private Sub cmdsearch4_Click()
    If txtkodecomp = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
    namatabel = "Company Account"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc4 = hasil
    lbldesc4 = hasil1
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    cmdadd1.SetFocus
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
    
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Row = 1
    Grid1.TextMatrix(0, 0) = "KodeSupp"
    Grid1.TextMatrix(0, 1) = "NamaSupp"
    Grid1.TextMatrix(0, 2) = "NoAcc"
    Grid1.TextMatrix(0, 3) = "Desc"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 2000
    Grid1.RowHeightMin = 300
    
    OBJ.Open dsn
    SQL = "select distinct kodesupp from am_apcashhdr"
    SQL = SQL + " where tglbkt >= '" & tanggal4 & "' and tglbkt <= '" & tanggal5 & "' and posted = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ0.Open dsn
        SQL0 = "select * from am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & RST!kodesupp & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        If RST0.EOF Then
            Grid1.TextMatrix(Grid1.Row, 0) = RST!kodesupp
            
            SQL0 = "select namasupp from am_supplier where kodesupp = '" & RST!kodesupp & "'"
            Set RST0 = OBJ0.Execute(SQL0)
            Grid1.TextMatrix(Grid1.Row, 1) = RST0!namasupp
            
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Row = Grid1.Row + 1
        Else
            If RST0!noacc = "" Then
                Grid1.TextMatrix(Grid1.Row, 0) = RST!kodesupp
                
                SQL0 = "select namasupp from am_supplier where kodesupp = '" & RST!kodesupp & "'"
                Set RST0 = OBJ0.Execute(SQL0)
                Grid1.TextMatrix(Grid1.Row, 1) = RST0!namasupp
                
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Row = Grid1.Row + 1
            End If
        End If
        OBJ0.Close

        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Verify Account Complete.", vbInformation, "Information"
    cmdproses2.Enabled = True
End Sub

Private Sub cmdverifykoreksi_Click()
    If txtkodecomp = "" Then Exit Sub
    
    If date6 > date7 Then
        MsgBox "Invalid date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date6.Month <> date7.Month Or date6.Year <> date7.Year Then
        MsgBox "Bulan/Tahun antara batasan tanggal harus sama.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "User harus memilih proses debit note atau credit note.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Diharapkan user tidak melakukan proses yang lain selama proses Verify Account berlangsung." & vbCrLf & _
    "Lanjutkan Proses Verify Account ?", vbYesNo + vbInformation, "Information") = vbNo Then Exit Sub
    
    grid3.Clear
    grid3.Rows = 2
    grid3.Row = 1
    grid3.TextMatrix(0, 0) = "KodeSupp"
    grid3.TextMatrix(0, 1) = "NamaSupp"
    grid3.TextMatrix(0, 2) = "NoAcc"
    grid3.TextMatrix(0, 3) = "Desc"
    grid3.ColWidth(0) = 1000
    grid3.ColWidth(1) = 2000
    grid3.ColWidth(2) = 1000
    grid3.ColWidth(3) = 2000
    grid3.RowHeightMin = 300
    
    OBJ.Open dsn
    SQL = "select distinct kodesupp from am_apcashhdr"
    SQL = SQL + " where tglbkt >= '" & tanggal6 & "' and tglbkt <= '" & tanggal7 & "' and posted = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ0.Open dsn
        SQL0 = "select * from am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & RST!kodesupp & "'"
        Set RST0 = OBJ0.Execute(SQL0)
        If RST0.EOF Then
            grid3.TextMatrix(grid3.Row, 0) = RST!kodesupp
            
            SQL0 = "select namasupp from am_supplier where kodesupp = '" & RST!kodesupp & "'"
            Set RST0 = OBJ0.Execute(SQL0)
            grid3.TextMatrix(grid3.Row, 1) = RST0!namasupp
            
            grid3.Rows = grid3.Rows + 1
            grid3.Row = grid3.Row + 1
        Else
            If RST0!noacc = "" Then
                grid3.TextMatrix(grid3.Row, 0) = RST!kodesupp
                
                SQL0 = "select namasupp from am_supplier where kodesupp = '" & RST!kodesupp & "'"
                Set RST0 = OBJ0.Execute(SQL0)
                grid3.TextMatrix(grid3.Row, 1) = RST0!namasupp
                
                grid3.Rows = grid3.Rows + 1
                grid3.Row = grid3.Row + 1
            End If
        End If
        OBJ0.Close

        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Verify Account Complete.", vbInformation, "Information"
    cmdproses3.Enabled = True
End Sub

Private Sub date4_Change()
    If cmdproses2.Enabled = True Then cmdproses2.Enabled = False
End Sub

Private Sub date5_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    If cmdproses2.Enabled = True Then cmdproses2.Enabled = False
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
   
    grid.TextMatrix(0, 1) = "Jurnal"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    grid.TextMatrix(0, 4) = "D/K"
    grid.TextMatrix(0, 5) = "X/Y/Z"
    grid.TextMatrix(0, 6) = "Manual"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 3000
    grid.ColWidth(4) = 500
    grid.ColWidth(5) = 800
    grid.ColWidth(6) = 800
    grid.RowHeightMin = 300
    
    Grid1.TextMatrix(0, 0) = "KodeSupp"
    Grid1.TextMatrix(0, 1) = "NamaSupp"
    Grid1.TextMatrix(0, 2) = "NoAcc"
    Grid1.TextMatrix(0, 3) = "Desc"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 2000
    Grid1.RowHeightMin = 300
    
    grid2.TextMatrix(0, 0) = "NoTrx"
    grid2.TextMatrix(0, 1) = "Curr"
    grid2.TextMatrix(0, 2) = "Nilai"
    grid2.TextMatrix(0, 3) = "D/K"
    grid2.TextMatrix(0, 4) = "NoAccount"
    grid2.TextMatrix(0, 5) = "Description"
    grid2.TextMatrix(0, 6) = "TglTrx"
    grid2.TextMatrix(0, 7) = "Kurs"
    grid2.ColWidth(0) = 1000
    grid2.ColWidth(1) = 500
    grid2.ColWidth(2) = 1500
    grid2.ColWidth(3) = 500
    grid2.ColWidth(4) = 1200
    grid2.ColWidth(5) = 3500
    grid2.ColWidth(6) = 1200
    grid2.ColWidth(7) = 1000
    grid2.ColWidth(8) = 0
    grid2.RowHeightMin = 300
    
    grid3.TextMatrix(0, 0) = "KodeSupp"
    grid3.TextMatrix(0, 1) = "NamaSupp"
    grid3.TextMatrix(0, 2) = "NoAcc"
    grid3.TextMatrix(0, 3) = "Desc"
    grid3.ColWidth(0) = 1000
    grid3.ColWidth(1) = 2000
    grid3.ColWidth(2) = 1000
    grid3.ColWidth(3) = 2000
    grid3.RowHeightMin = 300
    
    date1.Value = Date
    date2.Value = Date
    date4.Value = Date
    date5.Value = Date
    date6.Value = Date
    date7.Value = Date
    
    OBJ.Open dsn
    SQL = "select kdkurs from gl_kurs where base='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str1 = RST!kdkurs
    OBJ.Close
    
    txtkode1 = "JB"
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
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
            frmdefinejurnalshow.Frame1.Visible = True
            frmdefinejurnalshow.Frame2.Visible = False
            frmdefinejurnalshow.Frame3.Visible = False
            
            frmdefinejurnalshow.Show vbModal
        Case 2
            If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecomp = "" Then Exit Sub
            
            If grid.Rows - 1 = 100 Then
                MsgBox "Maximum line 100", vbExclamation, "Warning"
                Exit Sub
            End If
            
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
        Case 4
            If grid.TextMatrix(grid.Row, 2) = "" Or txtkodecomp = "" Then Exit Sub
            
            frmdefinejurnalshow.Frame1.Visible = False
            frmdefinejurnalshow.Frame2.Visible = True
            frmdefinejurnalshow.Frame3.Visible = False
            frmdefinejurnalshow.Option9.Enabled = False
            
            frmdefinejurnalshow.Show vbModal
        Case 5
            If grid.TextMatrix(grid.Row, 2) = "" Or txtkodecomp = "" Then Exit Sub
            
            frmdefinejurnalshow.Frame1.Visible = False
            frmdefinejurnalshow.Frame2.Visible = False
            frmdefinejurnalshow.Frame3.Visible = True
            
            frmdefinejurnalshow.Show vbModal
        Case 6
            If grid.TextMatrix(grid.Row, 1) <> "Jurnal a" Or grid.TextMatrix(grid.Row, 2) = "" Or txtkodecomp = "" Then Exit Sub
            
            If grid.TextMatrix(grid.Row, 6) = "" Then grid.TextMatrix(grid.Row, 6) = "x" Else grid.TextMatrix(grid.Row, 6) = ""
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 1
            grid.Row = posrow
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
        Case 2
            grid.Row = posrow
            grid.Col = 2
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 2) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
                        
            OBJ.Open dsn
            SQL = "select * from gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 3) = RST!nmac
            Else
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
            End If
            OBJ.Close
                            
            SetRow grid.Row, True
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Case 4
            grid.Row = posrow
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
        Case 5
            grid.Row = posrow
            grid.Col = 5
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
    End Select
    SSTab1.Tab = 1
    SSTab1.Tab = 0
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    
    posrow = Grid1.Row
    Select Case Grid1.Col
        Case 2
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Or txtkodecomp = "" Then Exit Sub
            
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case Grid1.Col
        Case 2
            Grid1.Row = posrow
            Grid1.Col = 2
            Grid1.CellAlignment = 1
            Grid1.TextMatrix(Grid1.Row, 2) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
                        
            OBJ.Open dsn
            SQL = "select * from gl_masterac where noac = '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                Grid1.TextMatrix(Grid1.Row, 3) = RST!nmac
            Else
                Grid1.TextMatrix(Grid1.Row, 2) = ""
                Grid1.TextMatrix(Grid1.Row, 3) = ""
            End If
            OBJ.Close
    End Select
    SSTab1.Tab = 2
    SSTab1.Tab = 3
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    
    posrow = grid2.Row
    
    If grid2.TextMatrix(grid2.Row, 4) <> "" Then
        OBJ.Open dsn
        SQL = "select * from gl_masterac where noac = '" & grid2.TextMatrix(grid2.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then Label4 = RST!nmac Else Label4 = ""
        OBJ.Close
    End If
            
    Select Case grid2.Col
        Case 4
            If txtkodecomp = "" Then Exit Sub
            
            If grid2.TextMatrix(grid2.Row, 4) <> "" Then
                If MsgBox("Cancel this account ?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    grid2.TextMatrix(grid2.Row, 4) = ""
                    Exit Sub
                End If
            End If
            
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid2_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid2.Col
        Case 4
            grid2.Row = posrow
            grid2.Col = 4
            grid2.CellAlignment = 1
            grid2.TextMatrix(grid2.Row, 4) = hasil
            Label4 = hasil1
            
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            SSTab1.Tab = 1
            SSTab1.Tab = 2
    End Select
End Sub

Private Sub grid3_Click()
    If grid3.MouseRow = 0 Then Exit Sub
    
    posrow = grid3.Row
    Select Case grid3.Col
        Case 2
            If grid3.TextMatrix(grid3.Row, 1) = "" Or txtkodecomp = "" Then Exit Sub
            
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid3_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid3.Col
        Case 2
            grid3.Row = posrow
            grid3.Col = 2
            grid3.CellAlignment = 1
            grid3.TextMatrix(grid3.Row, 2) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
                        
            OBJ.Open dsn
            SQL = "select * from gl_masterac where noac = '" & grid3.TextMatrix(grid3.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid3.TextMatrix(grid3.Row, 3) = RST!nmac
            Else
                grid3.TextMatrix(grid3.Row, 2) = ""
                grid3.TextMatrix(grid3.Row, 3) = ""
            End If
            OBJ.Close
    End Select
    SSTab1.Tab = 2
    SSTab1.Tab = 4
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
    If KeyAscii = 13 Then txtacc4.SetFocus
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

Private Sub txtacc4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdadd1.SetFocus
End Sub

Private Sub txtacc4_LostFocus()
    If txtacc4 = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc4 & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If RST2.EOF Then
        txtacc4 = ""
        lbldesc4 = ""
        txtacc4.SetFocus
    Else
        lbldesc4 = RST2!nmac
    End If
    OBJ2.Close
End Sub

Private Sub txtkode2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdverify.SetFocus
End Sub

Private Sub txtkode3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdverifykoreksi.SetFocus
End Sub

Private Sub txtkodecomp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecomp_LostFocus
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

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
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
End Sub

Private Sub cariautojurnal()
    If txtkodecomp = "" Then Exit Sub
    
    grid.Clear
    grid.Rows = 2
    grid.TextMatrix(0, 1) = "Jurnal"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    grid.TextMatrix(0, 4) = "D/K"
    grid.TextMatrix(0, 5) = "X/Y/Z"
    grid.TextMatrix(0, 6) = "Manual"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 3000
    grid.ColWidth(4) = 500
    grid.ColWidth(5) = 800
    grid.ColWidth(6) = 800
    
    grid.RowHeightMin = 300
    grid.Row = 1
    
    OBJ.Open dsn
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal a' order by line"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!jurnal_
        grid.Col = 2
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 2) = RST!noacc
        grid.Col = 4
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 4) = RST!dk
        grid.Col = 5
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 5) = RST!kdkurs
        grid.TextMatrix(grid.Row, 6) = RST!nanti
        
        OBJ1.Open dsn
        SQL1 = "SELECT * FROM gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 3) = RST1!nmac
        OBJ1.Close
        
        SetRow grid.Row, True
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal b' order by line"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!jurnal_
        grid.Col = 2
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 2) = RST!noacc
        grid.Col = 4
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 4) = RST!dk
        grid.Col = 5
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 5) = RST!kdkurs
        grid.TextMatrix(grid.Row, 6) = RST!nanti
        
        OBJ1.Open dsn
        SQL1 = "SELECT * FROM gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 3) = RST1!nmac
        OBJ1.Close
        
        SetRow grid.Row, True
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal c' order by line"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!jurnal_
        grid.Col = 2
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 2) = RST!noacc
        grid.Col = 4
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 4) = RST!dk
        grid.Col = 5
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 5) = RST!kdkurs
        grid.TextMatrix(grid.Row, 6) = RST!nanti
        
        OBJ1.Open dsn
        SQL1 = "SELECT * FROM gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 3) = RST1!nmac
        OBJ1.Close
        
        SetRow grid.Row, True
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
        
    OBJ.Open dsn
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal d'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc1 = RST!noacc Else txtacc1 = ""
    
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc1 = RST!nmac Else lbldesc1 = ""
    
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal e'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc2 = RST!noacc Else txtacc2 = ""
    
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc2 = RST!nmac Else lbldesc2 = ""
    
    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal f'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc3 = RST!noacc Else txtacc3 = ""

    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc3 = RST!nmac Else lbldesc3 = ""

    SQL = "select * from am_auto where kodecomp = '" & txtkodecomp & "' and jurnal_ = 'Jurnal g'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then txtacc4 = RST!noacc Else txtacc4 = ""

    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "' and a.noac = '" & txtacc4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then lbldesc4 = RST!nmac Else lbldesc4 = ""
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

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
