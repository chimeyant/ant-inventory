VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmbeliaktiva 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   5715
   ClientTop       =   5235
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmbeliaktiva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3435
      Left            =   120
      TabIndex        =   26
      Top             =   1065
      Width           =   7935
      Begin VB.VScrollBar VScroll 
         Height          =   3360
         Left            =   7305
         TabIndex        =   28
         Top             =   60
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   7515
         TabIndex        =   27
         Top             =   0
         Width           =   7515
         Begin MSComCtl2.DTPicker date1 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   3480
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
            Format          =   176226307
            CurrentDate     =   37694
         End
         Begin VB.TextBox txtsusut 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   15
            TabIndex        =   12
            Top             =   4560
            Width           =   1215
         End
         Begin VB.TextBox txtbiaya 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   15
            TabIndex        =   13
            Top             =   4920
            Width           =   1215
         End
         Begin VB.TextBox txtlawan 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   15
            TabIndex        =   11
            Top             =   4200
            Width           =   1215
         End
         Begin VB.TextBox txtaktiva 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   15
            TabIndex        =   10
            Top             =   3840
            Width           =   1215
         End
         Begin VB.OptionButton opsakhir 
            Caption         =   "Akhir bulan"
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
            Left            =   5400
            TabIndex        =   17
            Top             =   5475
            Width           =   1215
         End
         Begin VB.OptionButton opsawal 
            Caption         =   "Awal Bulan"
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
            Left            =   4080
            TabIndex        =   16
            Top             =   5475
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtdept 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   30
            TabIndex        =   5
            Top             =   2040
            Width           =   5340
         End
         Begin VB.TextBox txtlokasi 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1680
            Width           =   5340
         End
         Begin VB.TextBox txtnamafa 
            Appearance      =   0  'Flat
            DataField       =   "NamaArea"
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
            MaxLength       =   50
            TabIndex        =   2
            Top             =   915
            Width           =   5340
         End
         Begin VB.TextBox txtjenis 
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
            MaxLength       =   2
            TabIndex        =   3
            Top             =   1275
            Width           =   735
         End
         Begin TDBText6Ctl.TDBText txtcom 
            Height          =   285
            Left            =   1575
            TabIndex        =   0
            Top             =   195
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Caption         =   "frmbeliaktiva.frx":2372
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":23DE
            Key             =   "frmbeliaktiva.frx":23FC
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
         Begin TDBText6Ctl.TDBText txtkodefa 
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   555
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "frmbeliaktiva.frx":2440
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":24AC
            Key             =   "frmbeliaktiva.frx":24CA
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
            MaxLength       =   8
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
            TabIndex        =   29
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
            MICON           =   "frmbeliaktiva.frx":250E
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
            TabIndex        =   30
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
            MICON           =   "frmbeliaktiva.frx":2828
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Chameleon.chameleonButton cmdsearch7 
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   1275
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Jenis Aktiva"
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
            MICON           =   "frmbeliaktiva.frx":2B42
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TDBNumber6Ctl.TDBNumber txtbeli 
            Height          =   285
            Left            =   1560
            TabIndex        =   8
            Top             =   3120
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmbeliaktiva.frx":2E5C
            Caption         =   "frmbeliaktiva.frx":2E7C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":2EE8
            Keys            =   "frmbeliaktiva.frx":2F06
            Spin            =   "frmbeliaktiva.frx":2F50
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBText6Ctl.TDBText txtkodecur 
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Top             =   2400
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Caption         =   "frmbeliaktiva.frx":2F78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":2FE4
            Key             =   "frmbeliaktiva.frx":3002
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
         Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   2760
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmbeliaktiva.frx":303E
            Caption         =   "frmbeliaktiva.frx":305E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":30CA
            Keys            =   "frmbeliaktiva.frx":30E8
            Spin            =   "frmbeliaktiva.frx":312A
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin Chameleon.chameleonButton cmdsearch8 
            Height          =   285
            Left            =   240
            TabIndex        =   35
            Top             =   2400
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
            MICON           =   "frmbeliaktiva.frx":3152
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
            Left            =   120
            TabIndex        =   43
            Top             =   3840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Acc Aktiva"
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
            MICON           =   "frmbeliaktiva.frx":346C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Chameleon.chameleonButton cmdsearch4 
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   4200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Acc Lawan Aktiva"
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
            MICON           =   "frmbeliaktiva.frx":3786
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Chameleon.chameleonButton cmdsearch5 
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Acc Penyusutan"
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
            MICON           =   "frmbeliaktiva.frx":3AA0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TDBNumber6Ctl.TDBNumber txtsisa 
            Height          =   285
            Left            =   1560
            TabIndex        =   15
            Top             =   5640
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            Calculator      =   "frmbeliaktiva.frx":3DBA
            Caption         =   "frmbeliaktiva.frx":3DDA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":3E46
            Keys            =   "frmbeliaktiva.frx":3E64
            Spin            =   "frmbeliaktiva.frx":3EAE
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
            ValueVT         =   1245189
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
            Calculator      =   "frmbeliaktiva.frx":3ED6
            Caption         =   "frmbeliaktiva.frx":3EF6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmbeliaktiva.frx":3F62
            Keys            =   "frmbeliaktiva.frx":3F80
            Spin            =   "frmbeliaktiva.frx":3FCA
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
            HighlightText   =   1
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   0
            ValueVT         =   2085486597
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin Chameleon.chameleonButton cmdsearch6 
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Top             =   4920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Acc Biaya"
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
            MICON           =   "frmbeliaktiva.frx":3FF2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            Left            =   2880
            TabIndex        =   53
            Top             =   4920
            Width           =   3975
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
            Left            =   2880
            TabIndex        =   52
            Top             =   4560
            Width           =   3975
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
            Left            =   2880
            TabIndex        =   51
            Top             =   4200
            Width           =   3975
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
            Left            =   2880
            TabIndex        =   50
            Top             =   3840
            Width           =   3975
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            Caption         =   "Jurnal Penyusutan"
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
            Left            =   4080
            TabIndex        =   49
            Top             =   5280
            Width           =   1935
         End
         Begin VB.Label Label12 
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
            Left            =   240
            TabIndex        =   48
            Top             =   5670
            Width           =   1935
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
            TabIndex        =   47
            Top             =   5310
            Width           =   3015
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
            Left            =   3720
            TabIndex        =   42
            Top             =   3120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            Caption         =   "Nilai Kurs"
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
            Top             =   2790
            Width           =   1335
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
            TabIndex        =   40
            Top             =   2400
            Width           =   4500
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
            TabIndex        =   39
            Top             =   3150
            Width           =   1575
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
            TabIndex        =   38
            Top             =   3510
            Width           =   1575
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
            TabIndex        =   37
            Top             =   2070
            Width           =   1575
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
            TabIndex        =   36
            Top             =   1710
            Width           =   1335
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
            Left            =   255
            TabIndex        =   34
            Top             =   945
            Width           =   1335
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
            TabIndex        =   33
            Top             =   195
            Width           =   4500
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
            Left            =   2400
            TabIndex        =   32
            Top             =   1275
            Width           =   4500
         End
      End
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   4680
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
      MICON           =   "frmbeliaktiva.frx":430C
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
      TabIndex        =   19
      Top             =   4680
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
      MICON           =   "frmbeliaktiva.frx":4626
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
      TabIndex        =   20
      Top             =   4680
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
      MICON           =   "frmbeliaktiva.frx":4940
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
      TabIndex        =   21
      Top             =   4680
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
      MICON           =   "frmbeliaktiva.frx":4C5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pembelian"
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
      TabIndex        =   23
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label18 
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
      TabIndex        =   22
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmbeliaktiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String

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
    txtnamafa = ""
    txtlokasi = ""
    txtdept = ""
    date1.Value = Date
    txtbeli = 0
    txtaktiva = ""
    lblaktiva = ""
    txtlawan = ""
    lblawan = ""
    txtsusut = ""
    lblsusut = ""
    txtbiaya = ""
    lblbiaya = ""
    txtumur = 0
    txtjenis = ""
    lbljenis = ""
    txtsisa = 0
    txtkodecur = ""
    lblnamacur = ""
    txtnilaikurs = 0
    opsawal.Value = True
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
    
    If txtcom = "" Or txtkodefa = "" Or txtnamafa = "" Or txtlokasi = "" Or txtdept = "" Or txtbeli = 0 Or txtumur = 0 Or txtaktiva = "" Or txtsusut = "" Or txtbiaya = "" Or txtlawan = "" Or txtjenis = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If posted.Visible = True Then
        MsgBox "Can Not Delete, Record Still Posted.", vbExclamation, "Warning"
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete gl_aktiva where kdcomp = '" & txtcom & "' and kdaktiva = '" & txtkodefa & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdSave_Click()
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Access denied" & vbCrLf & _
            "Computer name : " & RST!comp_id & "Username : " & UserOnline & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Unload Me
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If txtkodecur = "" Or txtnilaikurs = 0 Or txtcom = "" Or txtkodefa = "" Or txtnamafa = "" Or txtlokasi = "" Or txtdept = "" Or txtbeli = 0 Or txtumur = 0 Or txtaktiva = "" Or txtsusut = "" Or txtbiaya = "" Or txtlawan = "" Or txtjenis = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtcom)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtkodefa)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtcom = Trim(txtcom)
    txtkodefa = Trim(txtkodefa)
    
    If posted.Visible = True Then
        MsgBox "Can Not Update, Record Still Posted.", vbExclamation, "Warning"
        cmdclear_Click
        Exit Sub
    End If
    
    If opsawal.Value = True Then
        str1 = "F"
    Else
        str1 = "L"
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE gl_aktiva SET "
        SQL = SQL + "nmaktiva = '" & txtnamafa & "',"
        SQL = SQL + "lokasi = '" & txtlokasi & "',"
        SQL = SQL + "dept = '" & txtdept & "',"
        SQL = SQL + "tglbeli = convert(datetime,'" & tanggal1 & "'),"
        SQL = SQL + "hargabeli = convert(money,'" & txtbeli & "'),"
        SQL = SQL + "ac_aktiva = '" & x_original(txtaktiva) & "',"
        SQL = SQL + "ac_susut = '" & x_original(txtsusut) & "',"
        SQL = SQL + "ac_lawan = '" & x_original(txtlawan) & "',"
        SQL = SQL + "ac_biaya = '" & x_original(txtbiaya) & "',"
        SQL = SQL + "umur = '" & txtumur & "',"
        SQL = SQL + "jenisfa = '" & txtjenis & "',"
        SQL = SQL + "jurnal = '" & str1 & "',"
        SQL = SQL + "nilaisisa = convert(money,'" & txtsisa & "'),"
        SQL = SQL + "curr = '" & txtkodecur & "',"
        SQL = SQL + "kurs = convert(money,'" & txtnilaikurs & "'),"
        SQL = SQL + "nilaibeli = convert(money,'" & (txtbeli * txtnilaikurs) & "'),"
        SQL = SQL + "dateupdate = convert(datetime,'" & tanggalsekarang & "'),"
        SQL = SQL + "idupdate = '" & nmuser & "'"
        SQL = SQL + "WHERE kdaktiva = '" & txtkodefa & "' and kdcomp = '" & txtcom & "'"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    Else
        SQL = "insert into gl_aktiva"
        SQL = SQL + "(kdcomp"
        SQL = SQL + ",kdaktiva"
        SQL = SQL + ",nmaktiva"
        SQL = SQL + ",lokasi"
        SQL = SQL + ",dept"
        SQL = SQL + ",tglbeli"
        SQL = SQL + ",tgljual"
        SQL = SQL + ",hargabeli"
        SQL = SQL + ",hargajual"
        SQL = SQL + ",ac_aktiva"
        SQL = SQL + ",ac_susut"
        SQL = SQL + ",ac_biaya"
        SQL = SQL + ",ac_lawan"
        SQL = SQL + ",umur"
        SQL = SQL + ",jenisfa"
        SQL = SQL + ",nilaisisa"
        SQL = SQL + ",flag"
        SQL = SQL + ",jurnal"
        SQL = SQL + ",curr"
        SQL = SQL + ",kurs"
        SQL = SQL + ",nilaibeli"
        SQL = SQL + ",curr1"
        SQL = SQL + ",kurs1"
        SQL = SQL + ",nilaijual"
        SQL = SQL + ",identry"
        SQL = SQL + ",idupdate"
        SQL = SQL + ",dateentry"
        SQL = SQL + ",dateupdate)"
        
        SQL = SQL + "VALUES"
        SQL = SQL + "('" & txtcom & "'"
        SQL = SQL + ", '" & txtkodefa & "'"
        SQL = SQL + ", '" & txtnamafa & "'"
        SQL = SQL + ", '" & txtlokasi & "'"
        SQL = SQL + ", '" & txtdept & "'"
        SQL = SQL + ", convert(datetime,'" & tanggal1 & "')"
        SQL = SQL + ", convert(datetime,' ')"
        SQL = SQL + ", convert(money,'" & txtbeli & "')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", '" & x_original(txtaktiva) & "'"
        SQL = SQL + ", '" & x_original(txtsusut) & "'"
        SQL = SQL + ", '" & x_original(txtbiaya) & "'"
        SQL = SQL + ", '" & x_original(txtlawan) & "'"
        SQL = SQL + ", '" & txtumur & "'"
        SQL = SQL + ", '" & txtjenis & "'"
        SQL = SQL + ", convert(money,'" & txtsisa & "')"
        SQL = SQL + ", 'N'"
        SQL = SQL + ", '" & str1 & "'"
        SQL = SQL + ", '" & txtkodecur & "'"
        SQL = SQL + ", convert(money,'" & txtnilaikurs & "')"
        SQL = SQL + ", convert(money,'" & (txtnilaikurs * txtbeli) & "')"
        SQL = SQL + ", ''"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", convert(money,'0')"
        SQL = SQL + ", '" & nmuser & "'"
        SQL = SQL + ", ''"
        SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "')"
        SQL = SQL + ", convert(datetime,' '))"
        Set RST = OBJ.Execute(SQL)
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    OBJ.Close
    cmdclear_Click
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
    namatabel = "Fixed Assets"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    hapusemua
    txtkodefa = hasil
    txtkodefa_LostFocus
    hasil = ""
End Sub

Private Sub cmdsearch3_Click()
    setup5 = "AS"
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "'"
    namatabel = "Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(hasil) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!flag = 0 Then
        txtaktiva = hasil
        lblaktiva = hasil1
    End If
    OBJ.Close
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch4_Click()
    setup5 = "AS"
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "'"
    namatabel = "Account  "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(hasil) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!flag = 0 Then
        txtlawan = hasil
        lblawan = hasil1
    End If
    OBJ.Close
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch5_Click()
    setup5 = "AS"
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "'"
    namatabel = "Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(hasil) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!flag = 0 Then
        txtsusut = hasil
        lblsusut = hasil1
    End If
    OBJ.Close
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch6_Click()
    setup5 = "EX"
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "'"
    namatabel = "Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch6_GotFocus()
    If hasil = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(hasil) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!flag = 0 Then
        txtbiaya = hasil
        lblbiaya = hasil1
    End If
    OBJ.Close
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch7_Click()
    carisql1 = "select kdjenis, nmjenis from gl_jenis"
    namatabel = "Jenis Fixed Assets"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch7_GotFocus()
    If hasil = "" Then Exit Sub
    txtjenis = hasil
    lbljenis = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
End Sub

Private Sub Form_Resize()
    VScroll.Max = Picture1.Height - 3600
    VScroll.LargeChange = CLng(VScroll.Max / 5)
    VScroll.SmallChange = CLng(VScroll.Max / 50)
End Sub

Private Sub txtaktiva_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtlawan.SetFocus
End Sub

Private Sub txtaktiva_LostFocus()
    If txtaktiva = "" Then Exit Sub
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtaktiva) & "' and typeac = 'AS'"
    SQL = "select b.noac, b.nmac, b.typeac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "' and a.noac = '" & x_original(txtaktiva) & "' and b.typeac = 'AS'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = 1 Then
            MsgBox "Account " & txtaktiva & " is header.", vbInformation, "Information"
            txtaktiva = ""
            lblaktiva = ""
            txtaktiva.SetFocus
            
            OBJ.Close
            Exit Sub
        End If
    End If
    If RST.EOF Then
        MsgBox "Account " & txtaktiva & " Not Found.", vbInformation, "Information"
        txtaktiva = ""
        txtaktiva.SetFocus
        
        OBJ.Close
        Exit Sub
    End If
    lblaktiva = RST!nmac
    OBJ.Close
End Sub

Private Sub txtbeli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtaktiva.SetFocus
End Sub

Private Sub txtbiaya_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtumur.SetFocus
End Sub

Private Sub txtbiaya_LostFocus()
    If txtbiaya = "" Then Exit Sub
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtbiaya) & "' and typeac = 'EX'"
    SQL = "select b.noac, b.nmac, b.typeac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "' and a.noac = '" & x_original(txtbiaya) & "' and b.typeac = 'EX'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = 1 Then
            MsgBox "Account " & txtbiaya & " is header.", vbInformation, "Information"
            txtbiaya = ""
            lblbiaya = ""
            txtbiaya.SetFocus
            
            OBJ.Close
            Exit Sub
        End If
    End If
    If RST.EOF Then
        MsgBox "Account " & txtbiaya & " Not Found.", vbInformation, "Information"
        txtbiaya = ""
        txtbiaya.SetFocus
        
        OBJ.Close
        Exit Sub
    End If
    lblbiaya = RST!nmac
    OBJ.Close
End Sub

Private Sub txtcom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodefa.SetFocus
End Sub

Private Sub txtcom_LostFocus()
    If txtcom = "" Then Exit Sub
    hapusemua
    txtkodefa = ""
    lblcom = ""
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

Private Sub txtjenis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtlokasi.SetFocus
End Sub

Private Sub txtjenis_LostFocus()
    If txtjenis = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_jenis where kdjenis = '" & txtjenis & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lbljenis = RST!nmjenis
        
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Jenis F/A " & txtjenis & " Not Found.", vbInformation, "Information"
    txtjenis = ""
    txtjenis.SetFocus
End Sub

Private Sub txtkodefa_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnamafa.SetFocus
End Sub

Private Sub txtkodefa_LostFocus()
    If txtkodefa = "" Then Exit Sub
    hapusemua
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom & "' and kdaktiva = '" & txtkodefa & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = "P" Or RST!flag = "J" Then
            posted.Visible = True
        Else
            posted.Visible = False
        End If
        txtnamafa = RST!nmaktiva
        txtlokasi = RST!lokasi
        txtdept = RST!dept
        date1.Value = RST!tglbeli
        txtbeli = RST!hargabeli
        txtaktiva = original(RST!ac_aktiva)
        txtlawan = original(RST!ac_lawan)
        txtsusut = original(RST!ac_susut)
        txtbiaya = original(RST!ac_biaya)
        txtumur = RST!umur
        txtjenis = RST!jenisfa
        txtsisa = RST!nilaisisa
        If RST!jurnal = "F" Then
            opsawal.Value = True
        Else
            opsakhir.Value = True
        End If
        txtkodecur = RST!curr
        txtnilaikurs = RST!kurs
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtkodecur & "'"
        Set RST = OBJ.Execute(SQL)
        lblnamacur = RST!nmkurs
        If RST!base = 1 Then
            lblbase = "1"
        Else
            lblbase = "0"
        End If
        
        SQL = "select * from gl_masterac where noac = '" & x_original(txtaktiva) & "'"
        Set RST = OBJ.Execute(SQL)
        lblaktiva = RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(txtlawan) & "'"
        Set RST = OBJ.Execute(SQL)
        lblawan = RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(txtsusut) & "'"
        Set RST = OBJ.Execute(SQL)
        lblsusut = RST!nmac
        
        SQL = "select * from gl_masterac where noac = '" & x_original(txtbiaya) & "'"
        Set RST = OBJ.Execute(SQL)
        lblbiaya = RST!nmac
        
        SQL = "select * from gl_jenis where kdjenis = '" & txtjenis & "'"
        Set RST = OBJ.Execute(SQL)
        lbljenis = RST!nmjenis
        
        txtlokasi.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
End Sub

Private Sub txtlawan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtsusut.SetFocus
End Sub

Private Sub txtlawan_LostFocus()
    If txtlawan = "" Then Exit Sub
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtlawan) & "' and typeac = 'AS'"
    SQL = "select b.noac, b.nmac, b.typeac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "' and a.noac = '" & x_original(txtlawan) & "' and b.typeac = 'AS'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = 1 Then
            MsgBox "Account " & txtlawan & " is header.", vbInformation, "Information"
            txtlawan = ""
            lblawan = ""
            txtlawan.SetFocus
            
            OBJ.Close
            Exit Sub
        End If
    End If
    If RST.EOF Then
        MsgBox "Account " & txtlawan & " Not Found.", vbInformation, "Information"
        txtlawan = ""
        txtlawan.SetFocus
        
        OBJ.Close
        Exit Sub
    End If
    lblawan = RST!nmac
    OBJ.Close
End Sub

Private Sub txtlokasi_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtdept.SetFocus
End Sub

Private Sub txtnamafa_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtjenis.SetFocus
End Sub

Private Sub txtdept_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtsisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtsusut_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtbiaya.SetFocus
End Sub

Private Sub txtsusut_LostFocus()
    If txtsusut = "" Then Exit Sub
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtsusut) & "' and typeac = 'AS'"
    SQL = "select b.noac, b.nmac, b.typeac, b.flag from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcom & "' and a.noac = '" & x_original(txtsusut) & "' and b.typeac = 'AS'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = 1 Then
            MsgBox "Account " & txtsusut & " is header.", vbInformation, "Information"
            txtsusut = ""
            lblsusut = ""
            txtsusut.SetFocus
            
            OBJ.Close
            Exit Sub
        End If
    End If
    If RST.EOF Then
        MsgBox "Account " & txtsusut & " Not Found.", vbInformation, "Information"
        txtsusut = ""
        txtsusut.SetFocus
        
        OBJ.Close
        Exit Sub
    End If
    lblsusut = RST!nmac
    OBJ.Close
End Sub

Private Sub txtumur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsisa.SetFocus
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
            txtnilaikurs = RST!kurs1
        Case 2
            txtnilaikurs = RST!kurs2
        Case 3
            txtnilaikurs = RST!kurs3
        Case 4
            txtnilaikurs = RST!kurs4
        Case 5
            txtnilaikurs = RST!kurs5
        Case 6
            txtnilaikurs = RST!kurs6
        Case 7
            txtnilaikurs = RST!kurs7
        Case 8
            txtnilaikurs = RST!kurs8
        Case 9
            txtnilaikurs = RST!kurs9
        Case 10
            txtnilaikurs = RST!kurs10
        Case 11
            txtnilaikurs = RST!kurs11
        Case 12
            txtnilaikurs = RST!kurs12
        End Select
        txtnilaikurs.SetFocus
    Else
        MsgBox "Currency " & txtkodecur & " Not Found.", vbInformation, "Information"
        txtkodecur = ""
        txtkodecur.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtbeli.SetFocus
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub cmdsearch8_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch8_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecur = hasil
    carikurs
    hasil = ""
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
