VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmbrowseacc 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Account"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmbrowseacc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Saldo Account"
      TabPicture(0)   =   "frmbrowseacc.frx":2372
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblperiode1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblperiode2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblperiode3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblperiode4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblperiode5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label19"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label21"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label22"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label23"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label25"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label27"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblperiode6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblperiode7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblperiode8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblperiode9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblperiode10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblperiode11"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblperiode12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblperiode13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Budget"
      TabPicture(1)   =   "frmbrowseacc.frx":238E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label31"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label32"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label33"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label34"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label35"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label36"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label38"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label39"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label40"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label41"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label42"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtotal"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtbudget12"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtbudget11"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtbudget10"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtbudget9"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtbudget8"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtbudget7"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtbudget13"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtbudget6"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtbudget5"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtbudget4"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtbudget3"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtbudget2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtbudget1"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      Begin TDBNumber6Ctl.TDBNumber txtbudget1 
         Height          =   285
         Left            =   -73590
         TabIndex        =   3
         Top             =   150
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":23AA
         Caption         =   "frmbrowseacc.frx":23CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2436
         Keys            =   "frmbrowseacc.frx":2454
         Spin            =   "frmbrowseacc.frx":2496
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget2 
         Height          =   285
         Left            =   -73590
         TabIndex        =   4
         Top             =   510
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":24BE
         Caption         =   "frmbrowseacc.frx":24DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":254A
         Keys            =   "frmbrowseacc.frx":2568
         Spin            =   "frmbrowseacc.frx":25AA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget3 
         Height          =   285
         Left            =   -73590
         TabIndex        =   5
         Top             =   870
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":25D2
         Caption         =   "frmbrowseacc.frx":25F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":265E
         Keys            =   "frmbrowseacc.frx":267C
         Spin            =   "frmbrowseacc.frx":26BE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget4 
         Height          =   285
         Left            =   -73590
         TabIndex        =   6
         Top             =   1230
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":26E6
         Caption         =   "frmbrowseacc.frx":2706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2772
         Keys            =   "frmbrowseacc.frx":2790
         Spin            =   "frmbrowseacc.frx":27D2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget5 
         Height          =   285
         Left            =   -73590
         TabIndex        =   7
         Top             =   1590
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":27FA
         Caption         =   "frmbrowseacc.frx":281A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2886
         Keys            =   "frmbrowseacc.frx":28A4
         Spin            =   "frmbrowseacc.frx":28E6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget6 
         Height          =   285
         Left            =   -73590
         TabIndex        =   8
         Top             =   1950
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":290E
         Caption         =   "frmbrowseacc.frx":292E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":299A
         Keys            =   "frmbrowseacc.frx":29B8
         Spin            =   "frmbrowseacc.frx":29FA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget13 
         Height          =   285
         Left            =   -70350
         TabIndex        =   15
         Top             =   1950
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2A22
         Caption         =   "frmbrowseacc.frx":2A42
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2AAE
         Keys            =   "frmbrowseacc.frx":2ACC
         Spin            =   "frmbrowseacc.frx":2B0E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget7 
         Height          =   285
         Left            =   -73590
         TabIndex        =   9
         Top             =   2310
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2B36
         Caption         =   "frmbrowseacc.frx":2B56
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2BC2
         Keys            =   "frmbrowseacc.frx":2BE0
         Spin            =   "frmbrowseacc.frx":2C22
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget8 
         Height          =   285
         Left            =   -70350
         TabIndex        =   10
         Top             =   150
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2C4A
         Caption         =   "frmbrowseacc.frx":2C6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2CD6
         Keys            =   "frmbrowseacc.frx":2CF4
         Spin            =   "frmbrowseacc.frx":2D36
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget9 
         Height          =   285
         Left            =   -70350
         TabIndex        =   11
         Top             =   510
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2D5E
         Caption         =   "frmbrowseacc.frx":2D7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2DEA
         Keys            =   "frmbrowseacc.frx":2E08
         Spin            =   "frmbrowseacc.frx":2E4A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget10 
         Height          =   285
         Left            =   -70350
         TabIndex        =   12
         Top             =   870
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2E72
         Caption         =   "frmbrowseacc.frx":2E92
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":2EFE
         Keys            =   "frmbrowseacc.frx":2F1C
         Spin            =   "frmbrowseacc.frx":2F5E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget11 
         Height          =   285
         Left            =   -70350
         TabIndex        =   13
         Top             =   1230
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":2F86
         Caption         =   "frmbrowseacc.frx":2FA6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":3012
         Keys            =   "frmbrowseacc.frx":3030
         Spin            =   "frmbrowseacc.frx":3072
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtbudget12 
         Height          =   285
         Left            =   -70350
         TabIndex        =   14
         Top             =   1590
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":309A
         Caption         =   "frmbrowseacc.frx":30BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":3126
         Keys            =   "frmbrowseacc.frx":3144
         Spin            =   "frmbrowseacc.frx":3186
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtotal 
         Height          =   285
         Left            =   -70350
         TabIndex        =   65
         Top             =   2310
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         Calculator      =   "frmbrowseacc.frx":31AE
         Caption         =   "frmbrowseacc.frx":31CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmbrowseacc.frx":323A
         Keys            =   "frmbrowseacc.frx":3258
         Spin            =   "frmbrowseacc.frx":329A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,###,##0.00;(##,###,###,##0.00);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,###,##0.00;(##,###,###,##0.00)"
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
         ValueVT         =   114229253
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label lblperiode13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   46
         Top             =   2310
         Width           =   2175
      End
      Begin VB.Label lblperiode12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   44
         Top             =   1950
         Width           =   2175
      End
      Begin VB.Label lblperiode11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   42
         Top             =   1590
         Width           =   2175
      End
      Begin VB.Label lblperiode10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   36
         Top             =   1230
         Width           =   2175
      End
      Begin VB.Label lblperiode9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   35
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label lblperiode8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   34
         Top             =   510
         Width           =   2175
      End
      Begin VB.Label lblperiode7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   4650
         TabIndex        =   33
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label lblperiode6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   32
         Top             =   1950
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         Left            =   -71310
         TabIndex        =   66
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "Budget 13"
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
         Left            =   -71190
         TabIndex        =   60
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Budget 12"
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
         Left            =   -71190
         TabIndex        =   59
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Budget 11"
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
         Left            =   -71190
         TabIndex        =   58
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "Budget 10"
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
         Left            =   -71190
         TabIndex        =   57
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Budget 09"
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
         Left            =   -71190
         TabIndex        =   56
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Budget 08"
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
         Left            =   -71190
         TabIndex        =   55
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Budget 07"
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
         Left            =   -74430
         TabIndex        =   54
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Budget 06"
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
         Left            =   -74430
         TabIndex        =   53
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Budget 05"
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
         Left            =   -74430
         TabIndex        =   52
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "Budget 04"
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
         Left            =   -74430
         TabIndex        =   51
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Budget 03"
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
         Left            =   -74430
         TabIndex        =   50
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Budget 02"
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
         Left            =   -74430
         TabIndex        =   49
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Budget 01"
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
         Left            =   -74430
         TabIndex        =   48
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Periode 13"
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
         Left            =   3810
         TabIndex        =   47
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Periode 12"
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
         Left            =   3810
         TabIndex        =   45
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Periode 11"
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
         Left            =   3810
         TabIndex        =   43
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Periode 10"
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
         Left            =   3810
         TabIndex        =   41
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Periode 09"
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
         Left            =   3810
         TabIndex        =   40
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Periode 08"
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
         Left            =   3810
         TabIndex        =   39
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Periode 07"
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
         Left            =   3810
         TabIndex        =   38
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Periode 06"
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
         Left            =   570
         TabIndex        =   37
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label lblperiode5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   31
         Top             =   1590
         Width           =   2175
      End
      Begin VB.Label lblperiode4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   30
         Top             =   1230
         Width           =   2175
      End
      Begin VB.Label lblperiode3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   29
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label lblperiode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   28
         Top             =   510
         Width           =   2175
      End
      Begin VB.Label lblperiode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Left            =   1410
         TabIndex        =   27
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Periode 05"
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
         Left            =   570
         TabIndex        =   26
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Periode 04"
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
         Left            =   570
         TabIndex        =   25
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Periode 03"
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
         Left            =   570
         TabIndex        =   24
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Periode 02"
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
         Left            =   570
         TabIndex        =   23
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Periode 01"
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
         Left            =   570
         TabIndex        =   22
         Top             =   150
         Width           =   1095
      End
   End
   Begin TDBText6Ctl.TDBText txtnoac 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmbrowseacc.frx":32C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmbrowseacc.frx":332E
      Key             =   "frmbrowseacc.frx":334C
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   5280
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmbrowseacc.frx":3388
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
      Left            =   5520
      TabIndex        =   18
      Top             =   5280
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmbrowseacc.frx":36A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmbrowseacc.frx":39BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtkode 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmbrowseacc.frx":3CD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmbrowseacc.frx":3D42
      Key             =   "frmbrowseacc.frx":3D60
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   68
      Top             =   1200
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
      MICON           =   "frmbrowseacc.frx":3D9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   360
      TabIndex        =   69
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Account"
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
      MICON           =   "frmbrowseacc.frx":40B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdelete 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmbrowseacc.frx":43D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblnamatype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   73
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   71
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   70
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblnamacomp 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   67
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lblnamacc 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      Caption         =   "Saldo Awal"
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
      Left            =   360
      TabIndex        =   64
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label lblsawal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   285
      Left            =   1680
      TabIndex        =   63
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblsakhir 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   285
      Left            =   4920
      TabIndex        =   62
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo Akhir"
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
      Left            =   3840
      TabIndex        =   61
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmbrowseacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2, str3, str4 As String

Private Sub cmdelete_Click()
    If txtKode = "" Or txtnoac = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_accrl where rl_ptd = '" & txtnoac & "' or rl_ytd = '" & txtnoac & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_transaksi WHERE noactrx = '" & txtnoac & "' and kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_dforms WHERE acc_no1 = '" & txtnoac & "' or acc_no2 = '" & txtnoac & "' or acc_no3 = '" & txtnoac & "' or acc_no4 = '" & txtnoac & "' or acc_no5 = '" & txtnoac & "' or acc_no6 = '" & txtnoac & "' or acc_no7 = '" & txtnoac & "' or acc_no8 = '" & txtnoac & "' or acc_no9 = '" & txtnoac & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_aktiva WHERE ac_aktiva = '" & txtnoac & "' or ac_susut = '" & txtnoac & "' or ac_biaya = '" & txtnoac & "' or ac_lawan = '" & txtnoac & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "DELETE FROM gl_chacct WHERE noac = '" & txtnoac & "' and kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
        
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
    Exit Sub
    
jump:
    MsgBox "Can Not Delete, Record Still In Use.", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdSave_Click()
    If txtnoac = "" Or txtKode = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtnoac)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtnoac = Trim(txtnoac)
    
    OBJ.Open dsn
    SQL = "select * from gl_chacct where noac = '" & txtnoac & "' and kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE gl_chacct SET "
        SQL = SQL + "budget01 = convert(money,'" & txtbudget1 & "'),"
        SQL = SQL + "budget02 = convert(money,'" & txtbudget2 & "'),"
        SQL = SQL + "budget03 = convert(money,'" & txtbudget3 & "'),"
        SQL = SQL + "budget04 = convert(money,'" & txtbudget4 & "'),"
        SQL = SQL + "budget05 = convert(money,'" & txtbudget5 & "'),"
        SQL = SQL + "budget06 = convert(money,'" & txtbudget6 & "'),"
        SQL = SQL + "budget07 = convert(money,'" & txtbudget7 & "'),"
        SQL = SQL + "budget08 = convert(money,'" & txtbudget8 & "'),"
        SQL = SQL + "budget09 = convert(money,'" & txtbudget9 & "'),"
        SQL = SQL + "budget10 = convert(money,'" & txtbudget10 & "'),"
        SQL = SQL + "budget11 = convert(money,'" & txtbudget11 & "'),"
        SQL = SQL + "budget12 = convert(money,'" & txtbudget12 & "'),"
        SQL = SQL + "budget13 = convert(money,'" & txtbudget13 & "')"
        SQL = SQL + "WHERE noac = '" & txtnoac & "' and kdcomp = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    Else
        SQL = "insert into gl_chacct"
        SQL = SQL + "(kdcomp"
        SQL = SQL + ",noac"
        SQL = SQL + ",typeac"
        SQL = SQL + ",balancedb"
        SQL = SQL + ",balancecr"
        SQL = SQL + ",begindb"
        SQL = SQL + ",begincr"
        SQL = SQL + ",periode01"
        SQL = SQL + ",periode02"
        SQL = SQL + ",periode03"
        SQL = SQL + ",periode04"
        SQL = SQL + ",periode05"
        SQL = SQL + ",periode06"
        SQL = SQL + ",periode07"
        SQL = SQL + ",periode08"
        SQL = SQL + ",periode09"
        SQL = SQL + ",periode10"
        SQL = SQL + ",periode11"
        SQL = SQL + ",periode12"
        SQL = SQL + ",periode13"
        SQL = SQL + ",last01"
        SQL = SQL + ",last02"
        SQL = SQL + ",last03"
        SQL = SQL + ",last04"
        SQL = SQL + ",last05"
        SQL = SQL + ",last06"
        SQL = SQL + ",last07"
        SQL = SQL + ",last08"
        SQL = SQL + ",last09"
        SQL = SQL + ",last10"
        SQL = SQL + ",last11"
        SQL = SQL + ",last12"
        SQL = SQL + ",last13"
        SQL = SQL + ",temp01"
        SQL = SQL + ",temp02"
        SQL = SQL + ",temp03"
        SQL = SQL + ",temp04"
        SQL = SQL + ",temp05"
        SQL = SQL + ",temp06"
        SQL = SQL + ",temp07"
        SQL = SQL + ",temp08"
        SQL = SQL + ",temp09"
        SQL = SQL + ",temp10"
        SQL = SQL + ",temp11"
        SQL = SQL + ",temp12"
        SQL = SQL + ",temp13"
        SQL = SQL + ",budget01"
        SQL = SQL + ",budget02"
        SQL = SQL + ",budget03"
        SQL = SQL + ",budget04"
        SQL = SQL + ",budget05"
        SQL = SQL + ",budget06"
        SQL = SQL + ",budget07"
        SQL = SQL + ",budget08"
        SQL = SQL + ",budget09"
        SQL = SQL + ",budget10"
        SQL = SQL + ",budget11"
        SQL = SQL + ",budget12"
        SQL = SQL + ",budget13)"
        
        SQL = SQL + "VALUES"
        SQL = SQL + "('" & txtKode & "'"
        SQL = SQL + ", '" & txtnoac & "'"
        SQL = SQL + ", '" & lbltype & "'"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'" & txtbudget1 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget2 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget3 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget4 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget5 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget6 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget7 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget8 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget9 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget10 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget11 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget12 & "')"
        SQL = SQL + ", convert(money,'" & txtbudget13 & "'))"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtKode = ""
    lblnamacomp = ""
    txtnoac = ""
    lblnamacc = ""
    lbltype = ""
    lblnamatype = ""
    hapusemua
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    txtKode_LostFocus
    hasil = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtKode & "'"
    namatabel = "Company Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtnoac = hasil
    lblnamacc = hasil1
    sstab.SetFocus
    hasil = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  
    
    str1 = 0
    str2 = 0
    str3 = 0
    str4 = 0
End Sub

Private Sub txtbudget1_LostFocus()
    txtotal = txtbudget1 + txtbudget2 + txtbudget3 + txtbudget4 + _
    txtbudget5 + txtbudget6 + txtbudget7 + txtbudget8 + txtbudget9 + _
    txtbudget10 + txtbudget11 + txtbudget12 + txtbudget13
End Sub

Private Sub txtbudget2_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget3_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget4_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget5_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget6_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget7_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget8_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget9_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget10_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget11_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget12_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtbudget13_LostFocus()
    txtbudget1_LostFocus
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnoac.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    If txtKode = "" Then Exit Sub
    hapusemua
    lblnamacomp = ""
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacomp = RST!nmcompscr
        format_coa = RST!formatac
    Else
        MsgBox "Company " & txtKode & " Not Found.", vbInformation, "Information"
        txtKode = ""
        lblnamacomp = ""
        txtKode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    If txtnoac <> "" Then
        SQL = "select * from gl_chacct where noac = '" & x_original(txtnoac) & "' and kdcomp = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblperiode1 = Format(RST!periode01, "###,###,###,##0.00")
            lblperiode2 = Format(RST!periode02, "###,###,###,##0.00")
            lblperiode3 = Format(RST!periode03, "###,###,###,##0.00")
            lblperiode4 = Format(RST!periode04, "###,###,###,##0.00")
            lblperiode5 = Format(RST!periode05, "###,###,###,##0.00")
            lblperiode6 = Format(RST!periode06, "###,###,###,##0.00")
            lblperiode7 = Format(RST!periode07, "###,###,###,##0.00")
            lblperiode8 = Format(RST!periode08, "###,###,###,##0.00")
            lblperiode9 = Format(RST!periode09, "###,###,###,##0.00")
            lblperiode10 = Format(RST!periode10, "###,###,###,##0.00")
            lblperiode11 = Format(RST!periode11, "###,###,###,##0.00")
            lblperiode12 = Format(RST!periode12, "###,###,###,##0.00")
            lblperiode13 = Format(RST!periode13, "###,###,###,##0.00")
            txtbudget1 = RST!budget01
            txtbudget2 = RST!budget02
            txtbudget3 = RST!budget03
            txtbudget4 = RST!budget04
            txtbudget5 = RST!budget05
            txtbudget6 = RST!budget06
            txtbudget7 = RST!budget07
            txtbudget8 = RST!budget08
            txtbudget9 = RST!budget09
            txtbudget10 = RST!budget10
            txtbudget11 = RST!budget11
            txtbudget12 = RST!budget12
            txtbudget13 = RST!budget13
            lblsawal = Format((RST!begindb + RST!begincr), "###,###,###,##0.00")
            lblsakhir = Format((RST!begindb + RST!begincr + RST!balancedb + RST!balancecr), "###,###,###,##0.00")
            str1 = RST!begindb
            str2 = RST!begincr
            str3 = RST!balancedb
            str4 = RST!balancecr
        End If
    End If
    OBJ.Close
End Sub

Private Sub txtnoac_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sstab.SetFocus
End Sub

Private Sub txtnoac_LostFocus()
    If txtnoac = "" Then Exit Sub
    hapusemua
    lblnamacc = ""
    lbltype = ""
    lblnamatype = ""
    OBJ.Open dsn
    SQL = "select * from gl_chacct where noac = '" & x_original(txtnoac) & "' and kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "select * from gl_masterac where noac = '" & x_original(txtnoac) & "'"
        Set RST = OBJ.Execute(SQL)
        txtnoac = original(RST!noac)
        lblnamacc = RST!nmac
        lbltype = RST!typeac
        
        Select Case lbltype
        Case "AS"
            lblnamatype = "Type Account : Assets"
        Case "LI"
            lblnamatype = "Type Account : Liability"
        Case "CA"
            lblnamatype = "Type Account : Capital"
        Case "IN"
            lblnamatype = "Type Account : Income"
        Case "EX"
            lblnamatype = "Type Account : Expenses"
        Case "IS"
            lblnamatype = "Type Account : Income Summary"
        End Select
    Else
        MsgBox "Account " & original(txtnoac) & " Not Found.", vbInformation, "Information"
        txtnoac = ""
        lblnamacc = ""
        lbltype = ""
        lblnamatype = ""
        txtnoac.SetFocus
        
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from gl_chacct where noac = '" & x_original(txtnoac) & "' and kdcomp = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblperiode1 = Format(RST!periode01, "###,###,###,##0.00")
        lblperiode2 = Format(RST!periode02, "###,###,###,##0.00")
        lblperiode3 = Format(RST!periode03, "###,###,###,##0.00")
        lblperiode4 = Format(RST!periode04, "###,###,###,##0.00")
        lblperiode5 = Format(RST!periode05, "###,###,###,##0.00")
        lblperiode6 = Format(RST!periode06, "###,###,###,##0.00")
        lblperiode7 = Format(RST!periode07, "###,###,###,##0.00")
        lblperiode8 = Format(RST!periode08, "###,###,###,##0.00")
        lblperiode9 = Format(RST!periode09, "###,###,###,##0.00")
        lblperiode10 = Format(RST!periode10, "###,###,###,##0.00")
        lblperiode11 = Format(RST!periode11, "###,###,###,##0.00")
        lblperiode12 = Format(RST!periode12, "###,###,###,##0.00")
        lblperiode13 = Format(RST!periode13, "###,###,###,##0.00")
        txtbudget1 = RST!budget01
        txtbudget2 = RST!budget02
        txtbudget3 = RST!budget03
        txtbudget4 = RST!budget04
        txtbudget5 = RST!budget05
        txtbudget6 = RST!budget06
        txtbudget7 = RST!budget07
        txtbudget8 = RST!budget08
        txtbudget9 = RST!budget09
        txtbudget10 = RST!budget10
        txtbudget11 = RST!budget11
        txtbudget12 = RST!budget12
        txtbudget13 = RST!budget13
        lblsawal = Format((RST!begindb + RST!begincr), "###,###,###,##0.00")
        lblsakhir = Format((RST!begindb + RST!begincr + RST!balancedb + RST!balancecr), "###,###,###,##0.00")
        str1 = RST!begindb
        str2 = RST!begincr
        str3 = RST!balancedb
        str4 = RST!balancecr
    End If
    OBJ.Close
    sstab.SetFocus
End Sub

Private Sub hapusemua()
    sstab.Tab = 0
    lblsawal = "0.00"
    lblsakhir = "0.00"
    lblperiode1 = "0.00"
    lblperiode2 = "0.00"
    lblperiode3 = "0.00"
    lblperiode4 = "0.00"
    lblperiode5 = "0.00"
    lblperiode6 = "0.00"
    lblperiode7 = "0.00"
    lblperiode8 = "0.00"
    lblperiode9 = "0.00"
    lblperiode10 = "0.00"
    lblperiode11 = "0.00"
    lblperiode12 = "0.00"
    lblperiode13 = "0.00"
    txtbudget1 = 0
    txtbudget2 = 0
    txtbudget3 = 0
    txtbudget4 = 0
    txtbudget5 = 0
    txtbudget6 = 0
    txtbudget7 = 0
    txtbudget8 = 0
    txtbudget9 = 0
    txtbudget10 = 0
    txtbudget11 = 0
    txtbudget12 = 0
    txtbudget13 = 0
End Sub
