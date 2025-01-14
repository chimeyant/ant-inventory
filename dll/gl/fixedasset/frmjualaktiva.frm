VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmjualaktiva 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   5715
   ClientTop       =   5235
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmjualaktiva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5435
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   10335
      Begin VB.VScrollBar VScroll 
         Height          =   5370
         Left            =   7140
         TabIndex        =   31
         Top             =   15
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10695
         Left            =   0
         ScaleHeight     =   10695
         ScaleWidth      =   7515
         TabIndex        =   30
         Top             =   225
         Width           =   7515
         Begin XtremeSuiteControls.CheckBox cbdisposal 
            Height          =   390
            Left            =   3645
            TabIndex        =   73
            Top             =   7995
            Width           =   1860
            _Version        =   851970
            _ExtentX        =   3281
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Disposal"
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
         Begin TDBText6Ctl.TDBText txtkodefa 
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   555
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":2372
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":23DE
            Key             =   "frmjualaktiva.frx":23FC
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   10
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtcom 
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   195
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":2440
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":24AC
            Key             =   "frmjualaktiva.frx":24CA
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   4
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
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
            Left            =   240
            TabIndex        =   32
            Top             =   195
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
            MICON           =   "frmjualaktiva.frx":250E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker date2 
            Height          =   255
            Left            =   3960
            TabIndex        =   37
            Top             =   2400
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   134414337
            CurrentDate     =   37747
         End
         Begin TDBNumber6Ctl.TDBNumber txtbeli 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   3480
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2828
            Caption         =   "frmjualaktiva.frx":2848
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":28B4
            Keys            =   "frmjualaktiva.frx":28D2
            Spin            =   "frmjualaktiva.frx":291C
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
            ValueVT         =   18153477
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
            Height          =   285
            Left            =   1560
            TabIndex        =   8
            Top             =   3120
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2944
            Caption         =   "frmjualaktiva.frx":2964
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":29D0
            Keys            =   "frmjualaktiva.frx":29EE
            Spin            =   "frmjualaktiva.frx":2A30
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483628
            BorderStyle     =   1
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
         Begin TDBNumber6Ctl.TDBNumber txtsisa 
            Height          =   285
            Left            =   4800
            TabIndex        =   15
            Top             =   5280
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2A58
            Caption         =   "frmjualaktiva.frx":2A78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":2AE4
            Keys            =   "frmjualaktiva.frx":2B02
            Spin            =   "frmjualaktiva.frx":2B4C
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
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   0
            ValueVT         =   1995898885
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin TDBNumber6Ctl.TDBNumber txtumur 
            Height          =   285
            Left            =   1560
            TabIndex        =   14
            Top             =   5280
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2B74
            Caption         =   "frmjualaktiva.frx":2B94
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":2C00
            Keys            =   "frmjualaktiva.frx":2C1E
            Spin            =   "frmjualaktiva.frx":2C68
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0;;Null"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "####0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   -99999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   0
            ValueVT         =   2085486597
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin TDBNumber6Ctl.TDBNumber txtjual 
            Height          =   285
            Left            =   1560
            TabIndex        =   20
            Top             =   7650
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2C90
            Caption         =   "frmjualaktiva.frx":2CB0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":2D1C
            Keys            =   "frmjualaktiva.frx":2D3A
            Spin            =   "frmjualaktiva.frx":2D84
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
            ValueVT         =   18153477
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSComCtl2.DTPicker date1 
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   6570
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
            Format          =   134414339
            CurrentDate     =   37694
         End
         Begin TDBText6Ctl.TDBText txtkodecur 
            Height          =   285
            Left            =   1560
            TabIndex        =   18
            Top             =   6930
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":2DAC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":2E18
            Key             =   "frmjualaktiva.frx":2E36
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
         Begin TDBNumber6Ctl.TDBNumber txtkursjual 
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Top             =   7290
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":2E72
            Caption         =   "frmjualaktiva.frx":2E92
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":2EFE
            Keys            =   "frmjualaktiva.frx":2F1C
            Spin            =   "frmjualaktiva.frx":2F5E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483628
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
            EditMode        =   1
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
            ValueVT         =   -65531
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin Chameleon.chameleonButton cmdsearch3 
            Height          =   285
            Left            =   240
            TabIndex        =   51
            Top             =   6930
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Kode Currency"
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
            MICON           =   "frmjualaktiva.frx":2F86
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
            Left            =   240
            TabIndex        =   33
            Top             =   555
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Kode Aktiva"
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
            MICON           =   "frmjualaktiva.frx":32A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TDBText6Ctl.TDBText txtkodebank 
            Height          =   285
            Left            =   1560
            TabIndex        =   57
            Top             =   8790
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":35BA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3626
            Key             =   "frmjualaktiva.frx":3644
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   10
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin Chameleon.chameleonButton cmdbank 
            Height          =   285
            Left            =   150
            TabIndex        =   58
            Top             =   8790
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Acc Bank"
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
            MICON           =   "frmjualaktiva.frx":3688
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TDBText6Ctl.TDBText txtBM 
            Height          =   285
            Left            =   1560
            TabIndex        =   61
            Top             =   8415
            Width           =   480
            _Version        =   65536
            _ExtentX        =   847
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":39A2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3A0E
            Key             =   "frmjualaktiva.frx":3A2C
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   10
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtnotran 
            Height          =   285
            Left            =   2115
            TabIndex        =   62
            Top             =   8415
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":3A70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3ADC
            Key             =   "frmjualaktiva.frx":3AFA
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
         Begin TDBNumber6Ctl.TDBNumber txtRL 
            Height          =   285
            Left            =   1560
            TabIndex        =   64
            Top             =   9930
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmjualaktiva.frx":3B36
            Caption         =   "frmjualaktiva.frx":3B56
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3BC2
            Keys            =   "frmjualaktiva.frx":3BE0
            Spin            =   "frmjualaktiva.frx":3C2A
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
            EditMode        =   1
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.00;(##,###,###,##0.00)"
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
         Begin TDBText6Ctl.TDBText txtnoaccRL 
            Height          =   285
            Left            =   1560
            TabIndex        =   65
            Top             =   10290
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":3C52
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3CBE
            Key             =   "frmjualaktiva.frx":3CDC
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
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
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   10
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtcekbg 
            Height          =   285
            Left            =   1560
            TabIndex        =   71
            Top             =   9165
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   503
            Caption         =   "frmjualaktiva.frx":3D20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmjualaktiva.frx":3D8C
            Key             =   "frmjualaktiva.frx":3DAA
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
            MaxLength       =   20
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
         Begin VB.Label R_L_BEP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4560
            TabIndex        =   74
            Top             =   9960
            Width           =   450
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            Caption         =   "No. Cheque/BG"
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
            Left            =   120
            TabIndex        =   72
            Top             =   9195
            Width           =   1335
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            Caption         =   "zz=<kode bank>"
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
            Height          =   180
            Left            =   5895
            TabIndex        =   70
            Top             =   8295
            Width           =   1335
         End
         Begin MSForms.ComboBox cmbdaerah 
            Height          =   330
            Left            =   1560
            TabIndex        =   69
            Top             =   8010
            Width           =   840
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1482;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            Caption         =   "Kode Daerah"
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
            Left            =   240
            TabIndex        =   68
            Top             =   8100
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "R/L Penjualan Fixed Assets"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   270
            TabIndex        =   67
            Top             =   9930
            Width           =   1170
         End
         Begin VB.Label lblakunRL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
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
            Left            =   2880
            TabIndex        =   66
            Top             =   10290
            Width           =   3990
         End
         Begin VB.Line Line2 
            X1              =   195
            X2              =   6870
            Y1              =   9765
            Y2              =   9765
         End
         Begin VB.Line Line1 
            X1              =   180
            X2              =   6840
            Y1              =   6255
            Y2              =   6255
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            Caption         =   "(manual) BM=Bank Masuk (YYMM/zz/XXXXX)"
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
            Left            =   3450
            TabIndex        =   63
            Top             =   8445
            Width           =   3255
         End
         Begin VB.Label lblnamaacc 
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
            Left            =   2880
            TabIndex        =   60
            Top             =   8790
            Width           =   3990
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            Caption         =   "Kode/No Transaksi"
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
            Left            =   165
            TabIndex        =   59
            Top             =   8445
            Width           =   1350
         End
         Begin VB.Label lblcom 
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
            Left            =   2400
            TabIndex        =   34
            Top             =   195
            Width           =   4455
         End
         Begin VB.Label lblnamacur 
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
            Left            =   2400
            TabIndex        =   56
            Top             =   6930
            Width           =   4455
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            Caption         =   "Nilai Kurs Jual"
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
            Left            =   240
            TabIndex        =   55
            Top             =   7320
            Width           =   1335
         End
         Begin VB.Label lblbase 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            TabIndex        =   54
            Top             =   7290
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            Caption         =   "Tanggal Jual"
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
            Left            =   240
            TabIndex        =   53
            Top             =   6600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            Caption         =   "Harga Jual"
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
            Left            =   240
            TabIndex        =   52
            Top             =   7680
            Width           =   1575
         End
         Begin VB.Label lblawan 
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
            Left            =   1560
            TabIndex        =   11
            Top             =   4200
            Width           =   5295
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            Caption         =   "Acc Biaya"
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
            Left            =   240
            TabIndex        =   50
            Top             =   4950
            Width           =   975
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            Caption         =   "Acc Penyusutan"
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
            Left            =   240
            TabIndex        =   49
            Top             =   4590
            Width           =   1215
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            Caption         =   "Acc Aktiva"
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
            Left            =   240
            TabIndex        =   48
            Top             =   3870
            Width           =   855
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            Caption         =   "Acc Lawan Aktiva"
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
            Left            =   240
            TabIndex        =   47
            Top             =   4230
            Width           =   1335
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            Caption         =   "Jurnal Penyusut"
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
            Left            =   240
            TabIndex        =   46
            Top             =   5670
            Width           =   1215
         End
         Begin VB.Label lbljurnal 
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
            Left            =   1560
            TabIndex        =   16
            Top             =   5640
            Width           =   5295
         End
         Begin VB.Label lblbiaya 
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
            Left            =   1560
            TabIndex        =   13
            Top             =   4920
            Width           =   5295
         End
         Begin VB.Label lblsusut 
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
            Left            =   1560
            TabIndex        =   12
            Top             =   4560
            Width           =   5295
         End
         Begin VB.Label lblaktiva 
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
            Left            =   1560
            TabIndex        =   10
            Top             =   3840
            Width           =   5295
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Nilai Sisa"
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
            Left            =   3720
            TabIndex        =   45
            Top             =   5310
            Width           =   975
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            Caption         =   "Umur                                          Bulan"
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
            Left            =   240
            TabIndex        =   44
            Top             =   5310
            Width           =   2775
         End
         Begin VB.Label lblcur 
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
            Left            =   1560
            TabIndex        =   7
            Top             =   2760
            Width           =   5295
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            Caption         =   "Nilai Kurs Beli"
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
            Left            =   240
            TabIndex        =   43
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            Caption         =   "Kode Currency "
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
            Left            =   240
            TabIndex        =   42
            Top             =   2790
            Width           =   1215
         End
         Begin VB.Label lbltglbeli 
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
            Left            =   1560
            TabIndex        =   6
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label lbldept 
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
            Left            =   1560
            TabIndex        =   5
            Top             =   2040
            Width           =   5295
         End
         Begin VB.Label lblokasi 
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
            Left            =   1560
            TabIndex        =   4
            Top             =   1680
            Width           =   5295
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            Caption         =   "Harga Beli"
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
            Left            =   240
            TabIndex        =   41
            Top             =   3510
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            Caption         =   "Tanggal Beli"
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
            Left            =   240
            TabIndex        =   40
            Top             =   2430
            Width           =   975
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            Caption         =   "Departement"
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
            Left            =   240
            TabIndex        =   39
            Top             =   2070
            Width           =   1095
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Lokasi"
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
            Left            =   240
            TabIndex        =   38
            Top             =   1710
            Width           =   975
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            Caption         =   "Jenis Aktiva"
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
            Left            =   240
            TabIndex        =   36
            Top             =   1305
            Width           =   975
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            Caption         =   "Nama Aktiva"
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
            Left            =   240
            TabIndex        =   35
            Top             =   945
            Width           =   975
         End
         Begin VB.Label lblnamafa 
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
            Left            =   1560
            TabIndex        =   2
            Top             =   915
            Width           =   5295
         End
         Begin VB.Label lbljenis 
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
            Left            =   1560
            TabIndex        =   3
            Top             =   1275
            Width           =   5295
         End
      End
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   6765
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmjualaktiva.frx":3DE6
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
      TabIndex        =   22
      Top             =   6765
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmjualaktiva.frx":4100
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   6765
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmjualaktiva.frx":441A
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
      Left            =   6480
      TabIndex        =   24
      Top             =   6765
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmjualaktiva.frx":4734
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets"
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
      TabIndex        =   27
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penjualan"
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
      TabIndex        =   26
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label posted 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POSTED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5760
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmjualaktiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL                 As String
Dim susut               As Double
Dim jual                As Double
Dim beli                As Double
Private BEP             As Boolean
Private DISPOSAL        As Boolean

Private Sub cbdisposal_Click()
    If cbdisposal.Value = xtpChecked Then
        txtBM.Enabled = False
        txtnotran.Enabled = False
        cmdbank.Enabled = False
        txtkodebank.Enabled = False
        If txtsisa.Value = "0" Or IsNull(txtsisa.Value) Then DISPOSAL = True
    ElseIf cbdisposal.Value = xtpUnchecked Then
        txtBM.Enabled = True
        txtnotran.Enabled = True
        cmdbank.Enabled = True
        txtkodebank.Enabled = True
    End If
End Sub

Private Sub cmdbank_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac"
    'carisql1 = carisql1 + " where a.kdcomp = '" & txtcom & "' and b.noac >= '11102001' and b.noac <= '11102014'"
    carisql1 = carisql1 + " where a.kdcomp = '01' and b.noac in('11102001','11102002','11102003','11102004','11102005','11102006','11102006','11102007','11102008','11102009','11102010','11102011','11102012','11102013','11102014','11102015','22001001','22001002','22001003','22001004','22001005','22001006','22001007','22001008','22001009','22001010','22001011','22001012','27011001','27011002')"
    'carisql1 = carisql1 + " order by b.noac asc"
    namatabel = "Company Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdbank_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodebank = hasil
    lblnamaacc = hasil1
    hasil = "": hasil1 = ""
    OBJ.Open dsn
    SQL = "Select * From gl_masterac Where noac = '86000000'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnoaccRL = RST!noac
        lblakunRL = RST!nmac
    End If
    OBJ.Close
End Sub

Private Sub cmdclear_Click()
    hapusemua
    txtkodefa = ""
    txtcom = ""
    lblcom = ""
    VScroll.Value = 0
    txtcom.SetFocus
End Sub

Private Sub hapusemua()
    posted.Visible = False
    lblnamafa = ""
    lblokasi = ""
    lbldept = ""
    date1.Value = Date
    txtbeli = 0
    lblaktiva = ""
    lbltglbeli = ""
    lblawan = ""
    lblsusut = ""
    lblcur = ""
    lblbiaya = ""
    txtumur = 0
    lbljenis = ""
    txtsisa = 0
    txtjual = 0
    lbljurnal = ""
    txtkodecur = ""
    txtkursjual = 0
    txtnilaikurs = 0
    lblnamacur = ""
    lblnamaacc = ""
    txtkodebank = ""
    txtBM = ""
    txtnotran = ""
    txtRL = 0
    txtnoaccRL = ""
    lblakunRL = ""
    cmbdaerah = ""
    DISPOSAL = False
    BEP = False
    R_L_BEP = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Private Sub cmdelete_Click()
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Access denied" & vbCrLf & _
            "Computer name : " & RST!comp_id & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Unload Me
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If txtcom = "" Or txtkodefa = "" Or txtjual = 0 Or txtkodecur = "" Or txtkursjual = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If posted.Visible = True Then
        MsgBox "Can Not Delete, Record Still Posted.", vbExclamation, "Warning"
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE gl_aktiva SET "
        SQL = SQL + "tgljual = convert(datetime,' '),"
        SQL = SQL + "hargajual = convert(money,'0'),"
        SQL = SQL + "nilaisisa = convert(money,'0'),"
        SQL = SQL + "kurs1 = convert(money,'0'),"
        SQL = SQL + "nilaijual = convert(money,'0'),"
        SQL = SQL + "curr1 = ' ',"
        SQL = SQL + "dateupdate = convert(datetime,'" & tanggalsekarang & "'),"
        SQL = SQL + "idupdate = '" & kuser & "'"
        SQL = SQL + "WHERE kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    End If
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdSave_Click()
    
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Access denied" & vbCrLf & _
            "Computer name : " & RST!comp_id & " Username : " & UserOnline & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Unload Me
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If txtcom = "" Or txtkodefa = "" Or txtkodecur = "" Or txtkursjual = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    If cbdisposal.Value = xtpUnchecked Then
        If txtjual = 0 Then
            MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
        End If
    End If
    
    txtcom = Trim(txtcom)
    txtkodefa = Trim(txtkodefa)
    
    If posted.Visible = True Then
        MsgBox "Can Not Update, Record Still Posted.", vbExclamation, "Warning"
        cmdclear_Click
        Exit Sub
    End If
    
    If date2 > date1 Then
        MsgBox "Sale Date Can Not Smaller Then Buy Date.", vbExclamation, "Warning"
        Exit Sub
    End If
        
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE gl_aktiva SET "
        SQL = SQL + "tgljual = convert(datetime,'" & tanggal1 & "'),"
        SQL = SQL + "hargajual = convert(money,'" & txtjual & "'),"
        SQL = SQL + "nilaisisa = convert(money,'" & txtsisa & "'),"
        SQL = SQL + "kurs1 = convert(money,'" & txtkursjual & "'),"
        SQL = SQL + "nilaijual = convert(money,'" & (txtkursjual * txtjual) & "'),"
        SQL = SQL + "curr1 = '" & txtkodecur & "',"
        SQL = SQL + "dateupdate = convert(datetime,'" & tanggalsekarang & "'),"
        SQL = SQL + "idupdate = '" & kuser & "'"
        SQL = SQL + "WHERE kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    End If
    OBJ.Close
    
    saverugilaba
    cmdclear_Click
End Sub

Private Sub saverugilaba()
    OBJ.Open dsn
    SQL = "Select * From gl_transaksi Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    If R_L_BEP = "BEP" Then GoTo bankin:
    If DISPOSAL = True Then GoTo dispos:
'RUGI/LABA  (86000000)
    With RST
        .AddNew
        !KDCOMP = txtcom
        !TGLTRX = Format(date1, "dd/MM/yyy") 'tanggal1
        !KDTRX = "JJ"
        !notrx = txtkodefa
        !kurs = txtkursjual
        !noactrx = txtnoaccRL
        !DESCTRX = "Penjualan " & lblnamafa
        If R_L_BEP = "R" Then
            !DBKRTRX = "D"
        ElseIf R_L_BEP = "L" Then
            !DBKRTRX = "K"
        End If
        !AMOUNTTRX = txtRL
        !nilaitrx = txtRL
        !CURRTRX = txtkodecur
        !flag = "P"
        !FLAGPRINT = "I"
        !flagadjustment = cmbdaerah
        !LINEITEM = "1"
        !IDENTRY = ""
        !IDUPDATE = ""
        !DATEENTRY = Format(date1, "dd/MM/yyy")
        !DATEUPDATE = Format(date1, "dd/MM/yyy")
        !cekbg = txtcekbg
        !RECONSIL = ""
        .Update
    End With
bankin:
'BANK IN
    With RST
        .AddNew
        !KDCOMP = txtcom
        !TGLTRX = Format(date1, "dd/MM/yyy")
        !KDTRX = txtBM
        !notrx = txtnotran
        !kurs = txtkursjual
        !noactrx = txtkodebank
        !DESCTRX = "Penjualan Aktiva " & txtkodefa
        !DBKRTRX = "D"
        !AMOUNTTRX = txtjual * txtkursjual
        !nilaitrx = txtjual
        !CURRTRX = txtkodecur
        !flag = "P"
        !FLAGPRINT = "I"
        !flagadjustment = cmbdaerah
        !LINEITEM = "2"
        !IDENTRY = ""
        !IDUPDATE = ""
        !DATEENTRY = Format(date1, "dd/MM/yyy")
        !DATEUPDATE = Format(date1, "dd/MM/yyy")
        !cekbg = txtcekbg
        !RECONSIL = ""
        .Update
    End With
'AKUMULASI PENYUSUTAN
dispos:
    With RST
        .AddNew
        !KDCOMP = txtcom
        !TGLTRX = Format(date1, "dd/MM/yyy")
        !KDTRX = "JJ"
        !notrx = txtkodefa
        !kurs = txtkursjual
        !noactrx = Left(lblsusut, 8)
        If DISPOSAL = True Then
            !DESCTRX = "Disposal " & lblnamafa
            !AMOUNTTRX = susut
            !nilaitrx = susut
        Else
            !DESCTRX = "Penjualan " & lblnamafa
            !AMOUNTTRX = txtjual * txtkursjual
            !nilaitrx = txtjual
        End If
        !DBKRTRX = "D"
        '!AMOUNTTRX = susut
        '!nilaitrx = susut
        !CURRTRX = txtkodecur
        !flag = "P"
        !FLAGPRINT = "I"
        !flagadjustment = cmbdaerah
        !LINEITEM = "3"
        !IDENTRY = ""
        !IDUPDATE = ""
        !DATEENTRY = Format(date1, "dd/MM/yyy")
        !DATEUPDATE = Format(date1, "dd/MM/yyy")
        !cekbg = txtcekbg
        !RECONSIL = ""
        .Update
    End With
    
    'With RST
        '.AddNew
        '!KDCOMP = txtcom
        '!TGLTRX = tanggal1
        '!KDTRX = "JJ"
        '!notrx = txtkodefa
        '!kurs = txtkursjual
        '!noactrx = Left(lblaktiva, 8)
        '!DESCTRX = "Penjualan " & lblnamafa
        '!DBKRTRX = "K"
        '!AMOUNTTRX = txtbeli
        '!nilaitrx = txtbeli
        '!CURRTRX = txtkodecur
        '!flag = "P"
        '!FLAGPRINT = "I"
        '!flagadjustment = cmbdaerah
        '!LINEITEM = "4"
        '!IDENTRY = ""
        '!IDUPDATE = ""
        '!DATEENTRY = tanggal1
        '!DATEUPDATE = tanggal1
        '!cekbg = txtcekbg
        '!RECONSIL = ""
        '.Update
    'End With
    
    MsgBox "Data Is Successfuly saved, Click OK To Continue ...", vbInformation, "Information"
    OBJ.Close
End Sub
Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    hapusemua
    txtkodefa = ""
    txtcom = hasil
    lblcom = hasil1
    txtcom_LostFocus
    hasil = ""
End Sub

Private Sub cmdsearch2_Click()
    setup6 = txtcom
    carisql1 = "select kdaktiva, nmaktiva from gl_aktiva"
    namatabel = " Fixed  Assets"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    hapusemua
    txtkodefa = hasil
    txtkodefa_LostFocus
    hasil = ""
End Sub

Private Sub cmbdaerah_LostFocus()
    If Not (cmbdaerah >= 1 And cmbdaerah <= 4) Then
        cmbdaerah = ""
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        cmbdaerah.SetFocus
    Else
        cari_in
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    date1 = Date
    cmbdaerah.Clear
    cmbdaerah.ColumnCount = 2
    cmbdaerah.ListWidth = "4 cm"
    cmbdaerah.ColumnWidths = "1 cm; 3 cm"
    cmbdaerah.AddItem "1"
    cmbdaerah.AddItem "2"
    cmbdaerah.AddItem "3"
    cmbdaerah.AddItem "4"
    cmbdaerah.List(0, 1) = "Pabrik"
    cmbdaerah.List(1, 1) = "Jakarta"
    cmbdaerah.List(2, 1) = "Surabaya"
    cmbdaerah.List(3, 1) = "Semarang"
End Sub

Private Sub Form_Resize()
    VScroll.Max = Picture1.Height - 3700
    VScroll.LargeChange = CLng(VScroll.Max / 5)
    VScroll.SmallChange = CLng(VScroll.Max / 50)
End Sub

Private Sub txtBM_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnotran.SetFocus
End Sub

Private Sub txtcom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodefa.SetFocus
End Sub

Private Sub txtcom_LostFocus()
    If txtcom = "" Then Exit Sub
    hapusemua
    txtkodefa = ""
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblcom = RST!nmcompscr
        format_coa = RST!formatac
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Company " & txtcom & " Not Found.", vbInformation, "Information"
    txtcom = ""
    txtcom.SetFocus
End Sub

Private Sub txtjual_Change()
    txtRL.Value = txtjual - txtsisa
    If txtsisa.Value = txtjual.Value Then
        R_L_BEP = "BEP"
    ElseIf txtjual.Value < txtsisa.Value Then
        R_L_BEP = "R"
    ElseIf txtjual.Value > txtsisa.Value Then
        R_L_BEP = "L"
    End If
End Sub

Private Sub txtkodefa_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtkodefa_LostFocus()
    If txtkodefa = "" Then Exit Sub
    hapusemua
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom & "' and kdaktiva = '" & txtkodefa & "' and (flag = 'P' or flag = 'J')"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = "J" Then
            posted.Visible = True
        Else
            posted.Visible = False
        End If
        lblnamafa = RST!nmaktiva
        lblokasi = RST!lokasi
        lbldept = RST!dept
        lbltglbeli = Format(RST!tglbeli, "dd MMMM yyyy")
        date2 = RST!tglbeli
        txtbeli = RST!hargabeli
        beli = RST!nilaibeli
        lblaktiva = original(RST!ac_aktiva)
        lblawan = original(RST!ac_lawan)
        lblsusut = original(RST!ac_susut)
        lblbiaya = original(RST!ac_biaya)
        txtumur = RST!umur
        lbljenis = RST!jenisfa

        If Trim(RST!curr1) <> "" Then date1 = RST!tgljual
        jual = RST!hargajual
        If RST!jurnal = "F" Then
            lbljurnal = "Awal Bulan"
        Else
            lbljurnal = "Akhir Bulan"
        End If
        txtkodecur = RST!curr1
        txtkursjual = RST!kurs1
        lblcur = RST!curr
        txtnilaikurs = RST!kurs
        
        'If posted.Visible = True Then
            txtsisa = RST!nilaisisa
            
            'CEK STATUS AKTIVA
            'PERIKSA AKUN AKTIVA IF KND% THEN 17410000, IF INV% THEN 17510000
            SQL = "Select * From gl_transaksi Where kdtrx = 'JJ' and notrx = '" & txtkodefa & "' and noactrx='17410000' and left(desctrx,8)='Disposal'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then cmbdaerah = RST!flagadjustment: cbdisposal.Value = xtpChecked: GoTo nextview:

            SQL = "Select * From gl_transaksi Where kdtrx = 'BM' and desctrx = 'Penjualan Aktiva ' + '" & txtkodefa & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                txtBM = "BM"
                txtnotran = RST!notrx
                txtkodebank = RST!noactrx
                txtcekbg = RST!cekbg
                cmbdaerah = RST!flagadjustment
            
                SQL = "Select nmac From gl_masterac Where noac = '" & RST!noactrx & "'"
                Set RST = OBJ.Execute(SQL)
                lblnamaacc = RST!nmac
            End If
            
            SQL = "Select * From gl_transaksi Where kdtrx = 'JJ' and notrx = '" & txtkodefa & "' and noactrx = '86000000'"
            Set RST = OBJ.Execute(SQL)
            
            If RST.EOF Then
                txtRL = "0.00"
            Else
                txtRL = RST!nilaitrx
            End If
            
            SQL = "Select * From gl_masterac Where noac = '86000000'"
            Set RST = OBJ.Execute(SQL)
            txtnoaccRL = RST!noac
            lblakunRL = RST!nmac

        'End If
nextview:
        txtjual = jual
        If txtkursjual <> 0 Then
            SQL = "select * from gl_kurs where kdkurs = '" & txtkodecur & "'"
            Set RST = OBJ.Execute(SQL)
            lblnamacur = RST!nmkurs
            If RST!base = 1 Then
                lblbase = "1"
            Else
                lblbase = "0"
            End If
        End If
        
        SQL = "select * from gl_kurs where kdkurs = '" & lblcur & "'"
        Set RST = OBJ.Execute(SQL)
        lblcur = lblcur & " - " & RST!nmkurs
        
        SQL = "select * from gl_masterac where noac = '" & x_original(lblaktiva) & "'"
        Set RST = OBJ.Execute(SQL)
        lblaktiva = lblaktiva & " - " & RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(lblawan) & "'"
        Set RST = OBJ.Execute(SQL)
        lblawan = lblawan & " - " & RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(lblsusut) & "'"
        Set RST = OBJ.Execute(SQL)
        lblsusut = lblsusut & " - " & RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(lblbiaya) & "'"
        Set RST = OBJ.Execute(SQL)
        lblbiaya = lblbiaya & " - " & RST!nmac
        
        SQL = "select * from gl_jenis where kdjenis = '" & lbljenis & "'"
        Set RST = OBJ.Execute(SQL)
        lbljenis = lbljenis & " - " & RST!nmjenis
        
        SQL = "SELECT notrx,sum(amounttrx)'susut' From gl_transaksi"
        'SQL = SQL + " WHERE kdtrx = 'JS' and tgltrx < '" & tanggalsekarang & "' and dbkrtrx='D' and notrx = '" & txtkodefa & "' Group by notrx"
        SQL = SQL + " WHERE kdtrx = 'JS' and tgltrx < '" & tanggal1 & "' and dbkrtrx='D' and notrx = '" & txtkodefa & "' Group by notrx"
        Set RST = OBJ.Execute(SQL)
        
        If posted.Visible = False Then
            susut = Format(RST!susut, "#,##0.0000")
            beli = Format(beli, "#,###,###,###,##0.0000")
            txtsisa = beli - susut
        End If

        date1.SetFocus
        OBJ.Close
        Exit Sub
    End If
    MsgBox "Aktiva " & txtkodefa & " Not Found.", vbInformation, "Information"
    txtkodefa = ""
    txtkodefa.SetFocus
    OBJ.Close
End Sub

Private Sub txtkodecur_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecur_LostFocus
End Sub

Private Sub txtkodecur_LostFocus()
    carikurs
End Sub

Private Sub carikurs()
    If txtkodecur = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_kurs where kdkurs = '" & txtkodecur & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacur = RST!nmkurs
        If RST!base = 1 Then
            lblbase = "1"
        Else
            lblbase = "0"
        End If
        Select Case Month(date1)
        Case 1
            txtkursjual = RST!kurs1
        Case 2
            txtkursjual = RST!kurs2
        Case 3
            txtkursjual = RST!kurs3
        Case 4
            txtkursjual = RST!kurs4
        Case 5
            txtkursjual = RST!kurs5
        Case 6
            txtkursjual = RST!kurs6
        Case 7
            txtkursjual = RST!kurs7
        Case 8
            txtkursjual = RST!kurs8
        Case 9
            txtkursjual = RST!kurs9
        Case 10
            txtkursjual = RST!kurs10
        Case 11
            txtkursjual = RST!kurs11
        Case 12
            txtkursjual = RST!kurs12
        End Select
        txtkursjual.SetFocus
    Else
        MsgBox "Currency " & txtkodecur & " Not Found.", vbInformation, "Information"
        txtkodecur = ""
        txtkodecur.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkursjual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtjual.SetFocus
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecur = hasil
    carikurs
    hasil = ""
End Sub

Private Sub txtnotran_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtBM <> "BM" Then KeyCode = 0
End Sub

Private Sub txtnotran_KeyPress(KeyAscii As Integer)
    If txtBM <> "BM" Then KeyAscii = 0 Else KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtBM <> "BM" Then txtRL.SetFocus
End Sub

Private Sub txtnotran_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtBM = "BM" Then
        'hapusemua
        cari_in
    End If
End Sub

Private Sub cari_in()
    If txtcom = "" Or txtBM = "" Or txtnotran = "" Or cmbdaerah = "" Then Exit Sub
    If txtBM = "BM" Then
        If Len(txtnotran) = 8 Then
            If Not (Left(txtnotran, 2) >= "08" And Left(txtnotran, 2) < "99") Then
                MsgBox "Format digit pertama dan kedua salah, format yang dipakai adalah format tahun, YY", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotran, 3, 2) >= "01" And Mid(txtnotran, 3, 2) <= "12") Then
                MsgBox "Format digit ketiga dan keempat salah, format yang dipakai adalah format bulan, MM", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not Mid(txtnotran, 5, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not Right(txtnotran, 1) = "/" Then
                MsgBox "Karakter pemisah, memakai garis miring, /.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            If Not (Mid(txtnotran, 6, 2) >= "01" And Mid(txtnotran, 6, 2) <= "09") And Not Mid(txtnotran, 6, 2) <= "11" Then
                MsgBox "Format digit keenam dan ketujuh salah, tekan F2 untuk melihat list.", vbInformation, "Information"
                txtnotran = ""
                txtnotran.SetFocus
                Exit Sub
            End If
            
            OBJ.Open dsn
            SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '" & txtcom & "' and kdtrx = '" & txtBM & "' and notrx like '" & txtnotran & "%' and flagprint='I' order by notrx desc"
            Set RST = OBJ.Execute(SQL)
          
            If Not RST.EOF Then
                If Len(RST!notrx + 1) = 5 Then
                    txtnotran = txtnotran & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 4 Then
                    txtnotran = txtnotran & "0" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 3 Then
                    txtnotran = txtnotran & "00" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 2 Then
                    txtnotran = txtnotran & "000" & RST!notrx + 1
                ElseIf Len(RST!notrx + 1) = 1 Then
                    txtnotran = txtnotran & "0000" & RST!notrx + 1
                End If
            Else
                txtnotran = txtnotran & "00001"
            End If
            OBJ.Close
            
'            txtRL.SetFocus
        End If
    Else
        OBJ.Open dsn
        SQL = "select top 1 right(notrx,5)'notrx' from gl_transaksi where kdcomp = '" & txtcom & "' and kdtrx = '" & txtBM & "' and notrx like '" & Format(date1, "YYMM") & "/" & cmbdaerah & "/%' and flagprint='I' order by notrx desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If Len(RST!notrx + 1) = 5 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 4 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 3 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 2 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/000" & RST!notrx + 1
            ElseIf Len(RST!notrx + 1) = 1 Then
                txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/0000" & RST!notrx + 1
            End If
        Else
            txtnotran = Format(date1, "YYMM") & "/" & cmbdaerah & "/00001"
        End If
        OBJ.Close
    End If
End Sub


Private Sub VScroll_Change()
    Picture1.Top = -VScroll.Value
End Sub

Private Sub VScroll_GotFocus()
    Picture1.SetFocus
End Sub

Private Sub VScroll_Scroll()
    Picture1.Top = -VScroll.Value
End Sub
