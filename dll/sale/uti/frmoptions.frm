VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options ..."
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tab             =   1
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmoptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check5"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Check4"
      Tab(0).Control(3)=   "Check3"
      Tab(0).Control(4)=   "Check2"
      Tab(0).Control(5)=   "Check1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Access and code"
      TabPicture(1)   =   "frmoptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkm1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkm2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkm3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkm4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chku2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtu2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkm6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkm7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkm8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkm9"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chku3"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtu3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chku4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtu4"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtu1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Head of Account Customer"
      TabPicture(2)   =   "frmoptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "txtkode"
      Tab(2).Control(2)=   "cmdsearch"
      Tab(2).Control(3)=   "txtkodeT"
      Tab(2).Control(4)=   "cmdsearchT"
      Tab(2).Control(5)=   "txtkodeI"
      Tab(2).Control(6)=   "cmdsearchI"
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   36
         Top             =   1680
         Width           =   4335
         Begin VB.Label Label6 
            Caption         =   "Master Account Customer akan otomatis bertambah sesuai 3 digit pertama account piutang dagang diatas."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Use Delivery Date to Calculate Stock."
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
         Left            =   -74760
         TabIndex        =   5
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtu1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   21
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtu4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox chku4 
         Caption         =   "use this code for SuratJalan"
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
         TabIndex        =   19
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtu3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chku3 
         Caption         =   "use this code for SalesOrder"
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
         TabIndex        =   17
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkm9 
         Caption         =   "Disable Collector"
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
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkm8 
         Caption         =   "Disable Salesman"
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
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkm7 
         Caption         =   "Disable Customer"
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
         Left            =   2280
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkm6 
         Caption         =   "Disable Area"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtu2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox chku2 
         Caption         =   "use this code for Customer"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkm4 
         Caption         =   "Disable Gudang"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkm3 
         Caption         =   "Disable Barang Jadi"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkm2 
         Caption         =   "Disable Category"
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
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkm1 
         Caption         =   "Disable Satuan"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Enable Button (Delivery Order)"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   2280
         Width           =   4335
         Begin Chameleon.chameleonButton cmdenable 
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "Enable"
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
            MICON           =   "frmoptions.frx":0054
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "to enable the button on delivery, just click this button."
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
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Automatically show Print/Skip on Save. (Surat Jalan)"
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Save SalesOrder to temporary. (am_soapp)"
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
         Left            =   -74760
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto numbering on delivery."
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calculate Stock."
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
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin TDBText6Ctl.TDBText txtkode 
         Height          =   285
         Left            =   -73320
         TabIndex        =   24
         Top             =   1320
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmoptions.frx":036E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmoptions.frx":03DA
         Key             =   "frmoptions.frx":03F8
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
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
         MaxLength       =   3
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
      Begin Chameleon.chameleonButton cmdsearch 
         Height          =   285
         Left            =   -74760
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Piutang Dagang"
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
         MICON           =   "frmoptions.frx":0434
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBText6Ctl.TDBText txtkodeT 
         Height          =   285
         Left            =   -73320
         TabIndex        =   22
         Top             =   480
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmoptions.frx":074E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmoptions.frx":07BA
         Key             =   "frmoptions.frx":07D8
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
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
      Begin Chameleon.chameleonButton cmdsearchT 
         Height          =   285
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Company Type"
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
         MICON           =   "frmoptions.frx":0814
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBText6Ctl.TDBText txtkodeI 
         Height          =   285
         Left            =   -73320
         TabIndex        =   23
         Top             =   840
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmoptions.frx":0B2E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmoptions.frx":0B9A
         Key             =   "frmoptions.frx":0BB8
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
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
      Begin Chameleon.chameleonButton cmdsearchI 
         Height          =   285
         Left            =   -74760
         TabIndex        =   35
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Company ID"
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
         MICON           =   "frmoptions.frx":0BF4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "LL           0000 (2)"
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
         Left            =   3015
         TabIndex        =   32
         Top             =   2535
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "PP           0000 (2)"
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
         Left            =   2985
         TabIndex        =   31
         Top             =   2295
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "P/L-           00000 (1)"
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
         Left            =   2880
         TabIndex        =   29
         Top             =   2055
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "C-           0000 (1)"
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
         Left            =   3000
         TabIndex        =   28
         Top             =   1815
         Width           =   1335
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   3480
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
      MICON           =   "frmoptions.frx":0F0E
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
      Left            =   2880
      TabIndex        =   25
      Top             =   3480
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
      MICON           =   "frmoptions.frx":1228
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub chkm7_Click()
    If chkm7.Value = 1 Then
        chku2.Value = 0
        chku2.Enabled = False
        txtu2 = ""
        txtu2.Enabled = False
    Else
        chku2.Enabled = True
        txtu2.Enabled = True
    End If
End Sub

Private Sub chku2_Click()
    txtu2 = ""
End Sub

Private Sub chku3_Click()
    txtu3 = ""
End Sub

Private Sub chku4_Click()
    txtu4 = ""
    txtu1 = ""
End Sub

Private Sub cmdclear_Click()
    If chku2.Value = 1 And txtu2 = "" Then
        MsgBox "Code for Customer is empty.", vbInformation, "Information"
        Exit Sub
    End If
    
    If chku3.Value = 1 And txtu3 = "" Then
        MsgBox "Code for SalesOrder is empty.", vbInformation, "Information"
        Exit Sub
    End If
    
    If chku4.Value = 1 And (txtu4 = "" Or txtu1 = "") Then
        MsgBox "Code for SuratJalan is empty.", vbInformation, "Information"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "update am_options set para1 = '" & Check1.Value & "',para2 = '" & Check2.Value & "',"
        SQL = SQL + "para3 = '" & Check3.Value & "',para4 = '" & Check4.Value & "',para6 = '" & Check5.Value & "',"
        SQL = SQL + "c_type = '" & txtkodeT & "',c_id = '" & txtkodeI & "',ac_cust = '" & txtkode & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "INSERT INTO AM_options "
        SQL = SQL + "(para1"
        SQL = SQL + ",para2"
        SQL = SQL + ",para3"
        SQL = SQL + ",para4"
        SQL = SQL + ",para5"
        SQL = SQL + ",para6"
        SQL = SQL + ",c_type"
        SQL = SQL + ",c_id"
        SQL = SQL + ",ac_cust"
        SQL = SQL + ",posted)"
        
        SQL = SQL + " VALUES "
        SQL = SQL + "('" & Check1.Value & "'"
        SQL = SQL + ",'" & Check2.Value & "'"
        SQL = SQL + ",'" & Check3.Value & "'"
        SQL = SQL + ",'" & Check4.Value & "'"
        SQL = SQL + ",'0'"
        SQL = SQL + ",'" & Check5.Value & "'"
        SQL = SQL + ",'" & txtkodeT & "'"
        SQL = SQL + ",'" & txtkodeI & "'"
        SQL = SQL + ",'" & txtkode & "'"
        SQL = SQL + ",convert(datetime,'" & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & "'))"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    par1 = Check1.Value
    par2 = Check2.Value
    par3 = Check3.Value
    par4 = Check4.Value
    par5 = Check5.Value
    
    OBJ.Open dsn
    SQL = "select * from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "delete from am_branch"
        Set RST = OBJ.Execute(SQL)
    End If
    
    SQL = "INSERT INTO AM_branch "
    SQL = SQL + "(dis1"
    SQL = SQL + ",dis2"
    SQL = SQL + ",dis3"
    SQL = SQL + ",dis4"
    SQL = SQL + ",dis5"
    SQL = SQL + ",dis6"
    SQL = SQL + ",dis7"
    SQL = SQL + ",dis8"
    SQL = SQL + ",dis9"
    SQL = SQL + ",dis10"
    SQL = SQL + ",id1"
    SQL = SQL + ",id2"
    SQL = SQL + ",id3"
    SQL = SQL + ",id4"
    SQL = SQL + ",kode1"
    SQL = SQL + ",kode2"
    SQL = SQL + ",kode3"
    SQL = SQL + ",kode4)"
    
    SQL = SQL + " VALUES "
    SQL = SQL + "('" & chkm1.Value & "'" 'satuan
    SQL = SQL + ",'" & chkm2.Value & "'" 'produk
    SQL = SQL + ",'" & chkm3.Value & "'" 'item
    SQL = SQL + ",'" & chkm4.Value & "'" 'gudang
    SQL = SQL + ",'0'" 'currency
    SQL = SQL + ",'" & chkm6.Value & "'" 'area
    SQL = SQL + ",'" & chkm7.Value & "'" 'customer
    SQL = SQL + ",'" & chkm8.Value & "'" 'salesman
    SQL = SQL + ",'" & chkm9.Value & "'" 'collector
    SQL = SQL + ",'0'" 'bank
    
    SQL = SQL + ",'0'"
    SQL = SQL + ",'" & chku2.Value & "'" 'customer
    SQL = SQL + ",'" & chku3.Value & "'" 'salesorder
    SQL = SQL + ",'" & chku4.Value & "'" 'suratjalan
    
    SQL = SQL + ",'" & txtu1 & "'"  'codesuratjalan LL
    SQL = SQL + ",'" & txtu2 & "'"  'codecustomer
    SQL = SQL + ",'" & txtu3 & "'"  'codesalesorder
    SQL = SQL + ",'" & txtu4 & "')" 'codesuratjalan PP
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    'If chkm1.Value = 1 Then
    '    frmainmenu.ssmnuaddsatuan.Enabled = False
    '    frmainmenu.ssmnupdatesatuan.Enabled = False
    'Else
    '    frmainmenu.ssmnuaddsatuan.Enabled = True
    '    frmainmenu.ssmnupdatesatuan.Enabled = True
    'End If
    'If chkm2.Value = 1 Then
    '    frmainmenu.ssmnuaddproduk.Enabled = False
    '    frmainmenu.ssmnupdateproduk.Enabled = False
    'Else
    '    frmainmenu.ssmnuaddproduk.Enabled = True
    '    frmainmenu.ssmnupdateproduk.Enabled = True
    'End If
    'If chkm3.Value = 1 Then frmainmenu.ssmnuaddbarang.Enabled = False Else frmainmenu.ssmnuaddbarang.Enabled = True
    'If chkm4.Value = 1 Then
    '    frmainmenu.ssmnuaddgudang.Enabled = False
    '    frmainmenu.ssmnupdategudang.Enabled = False
    'Else
    '    frmainmenu.ssmnuaddgudang.Enabled = True
    '    frmainmenu.ssmnupdategudang.Enabled = True
    'End If
   ' If chkm6.Value = 1 Then
   '     frmainmenu.ssmnuaddarea.Enabled = False
   '     frmainmenu.ssmnupdatearea.Enabled = False
   ' Else
   '     frmainmenu.ssmnuaddarea.Enabled = True
   '     frmainmenu.ssmnupdatearea.Enabled = True
   ' End If
   ' If chkm7.Value = 1 Then
   '     frmainmenu.ssmnuaddcust.Enabled = False
   '     frmainmenu.ssmnupdatecust.Enabled = False
   ' Else
   '     frmainmenu.ssmnuaddcust.Enabled = True
   '     frmainmenu.ssmnupdatecust.Enabled = True
   ' End If
   ' If chkm8.Value = 1 Then
   '     frmainmenu.ssmnuaddsales.Enabled = False
   '     frmainmenu.ssmnupdatesales.Enabled = False
   ' Else
   '     frmainmenu.ssmnuaddsales.Enabled = True
   '     frmainmenu.ssmnupdatesales.Enabled = True
   ' End If
   ' If chkm9.Value = 1 Then
   '     frmainmenu.ssmnuaddcollect.Enabled = False
   '     frmainmenu.ssmnupdatecollect.Enabled = False
   ' Else
   '     frmainmenu.ssmnuaddcollect.Enabled = True
   '     frmainmenu.ssmnupdatecollect.Enabled = True
   ' End If
    
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdenable_Click()
    OBJ.Open dsn
    SQL = "update am_options set para5 = '0'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Buttons enabled.", vbInformation, "Information"
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Master Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearchI_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearchI_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeI = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearchT_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearchT_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeT = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    OBJ.Open dsn
    SQL = "select * from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Check1.Value = RST!para1
        Check2.Value = RST!para2
        Check3.Value = RST!para3
        Check4.Value = RST!para4
        Check5.Value = RST!para6
        
        txtkodeT = RST!c_type
        txtkodeI = RST!c_id
        txtkode = RST!ac_cust
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        chkm1.Value = RST!dis1 'satuan
        chkm2.Value = RST!dis2 'produk
        chkm3.Value = RST!dis3 'item
        chkm4.Value = RST!dis4 'gudang
        chkm6.Value = RST!dis6 'area
        chkm7.Value = RST!dis7 'cust
        chkm8.Value = RST!dis8 'sales
        chkm9.Value = RST!dis9 'collect
        
        chku2.Value = RST!id2 'cust
        chku3.Value = RST!id3 'so
        chku4.Value = RST!id4 'sj
        
        txtu1 = RST!kode1 'sj ll
        txtu2 = RST!kode2 'cust
        txtu3 = RST!kode3 'so
        txtu4 = RST!kode4 'sj pp
    End If
    OBJ.Close
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtkodeI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkode.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtkodeT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkodeI.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtu1_KeyPress(KeyAscii As Integer)
    If chku4.Value = 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii))) Else KeyAscii = 0
End Sub

Private Sub txtu2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 48 Then KeyAscii = 0
    If chku2.Value = 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii))) Else KeyAscii = 0
End Sub

Private Sub txtu3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 48 Then KeyAscii = 0
    If chku3.Value = 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii))) Else KeyAscii = 0
End Sub

Private Sub txtu4_KeyPress(KeyAscii As Integer)
    If chku4.Value = 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii))) Else KeyAscii = 0
End Sub
