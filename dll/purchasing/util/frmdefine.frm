VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdefine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options..."
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4000
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7064
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Max Order"
      TabPicture(0)   =   "frmdefine.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Master Account Supplier"
      TabPicture(1)   =   "frmdefine.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "txtkode"
      Tab(1).Control(2)=   "cmdsearch"
      Tab(1).Control(3)=   "txtkodeT"
      Tab(1).Control(4)=   "cmdsearchT"
      Tab(1).Control(5)=   "txtkodeI"
      Tab(1).Control(6)=   "cmdsearchI"
      Tab(1).Control(7)=   "txtkode3"
      Tab(1).Control(8)=   "cmdsearch2"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "..."
      TabPicture(2)   =   "frmdefine.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   24
         Top             =   1800
         Width           =   3495
         Begin VB.Label Label5 
            Caption         =   "Master Account Supplier akan otomatis bertambah sesuai 2 digit pertama account hutang dagang diatas."
            Height          =   975
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.ListBox List1 
         Height          =   2985
         ItemData        =   "frmdefine.frx":0054
         Left            =   120
         List            =   "frmdefine.frx":0056
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   600
         Width           =   3495
      End
      Begin TDBText6Ctl.TDBText txtkode 
         Height          =   285
         Left            =   -73320
         TabIndex        =   13
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":0058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":00C4
         Key             =   "frmdefine.frx":00E2
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
         MaxLength       =   2
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Hutang Dagang"
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
         MICON           =   "frmdefine.frx":011E
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
         TabIndex        =   9
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":0438
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":04A4
         Key             =   "frmdefine.frx":04C2
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
         TabIndex        =   8
         Top             =   240
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
         MICON           =   "frmdefine.frx":04FE
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
         TabIndex        =   11
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":0818
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":0884
         Key             =   "frmdefine.frx":08A2
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
         TabIndex        =   10
         Top             =   600
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
         MICON           =   "frmdefine.frx":08DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBText6Ctl.TDBText txtkode3 
         Height          =   285
         Left            =   -73320
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":0BF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":0C64
         Key             =   "frmdefine.frx":0C82
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
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   285
         Left            =   -74760
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "PPn Masukan"
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
         MICON           =   "frmdefine.frx":0CBE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         Caption         =   "This following Purchase Order don't have maximum Quantity Order."
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Define PO Numbering Format"
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtnew 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   3
         Top             =   3720
         Width           =   2655
      End
      Begin Chameleon.chameleonButton cmdset 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   4080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add"
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
         MICON           =   "frmdefine.frx":0FD8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2895
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   15
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin TDBText6Ctl.TDBText txtkode1 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":12F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":135E
         Key             =   "frmdefine.frx":137C
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
         AllowSpace      =   0
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
      Begin TDBText6Ctl.TDBText txtkode2 
         Height          =   285
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   503
         Caption         =   "frmdefine.frx":13C0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdefine.frx":142C
         Key             =   "frmdefine.frx":144A
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   1
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
      Begin Chameleon.chameleonButton cmdremove 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   4080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Remove"
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
         MICON           =   "frmdefine.frx":148E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   14.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Purchase Order                                  001"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   390
         Width           =   3015
      End
      Begin VB.Label lbl1 
         Caption         =   "001"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "add new"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3750
         Width           =   735
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   285
      Left            =   -2040
      TabIndex        =   17
      Top             =   -1440
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   503
      Calculator      =   "frmdefine.frx":17A8
      Caption         =   "frmdefine.frx":17C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdefine.frx":1834
      Keys            =   "frmdefine.frx":1852
      Spin            =   "frmdefine.frx":1894
      AlignHorizontal =   2
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
      MaxValue        =   10
      MinValue        =   5
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   5
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   4260
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
      MICON           =   "frmdefine.frx":18BC
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
Attribute VB_Name = "frmdefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim int1, h, i As Integer

Private Sub cmdremove_Click()
    If MsgBox("Remove " & grid.TextMatrix(grid.Row, 0) & " ?", vbQuestion + vbYesNo, "Remove") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodeproduk = '" & grid.TextMatrix(grid.Row, 0) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not remove, kode already used.", vbInformation, "Information"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "delete from am_kode where kode3 = '" & grid.TextMatrix(grid.Row, 0) & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Clear
    
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode1 = RST!kode1
        txtkode2 = RST!kode2
        
        grid.Row = 0
        Do While Not RST.EOF
            grid.TextMatrix(grid.Row, 0) = RST!kode3
            
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    
    List1.Clear
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!kode3
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    For i = 0 To List1.ListCount - 1
        OBJ.Open dsn
        SQL = "select * from am_nomax where kode='" & List1.List(i) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then List1.Selected(i) = True
        OBJ.Close
    Next i
        
    MsgBox "Data remove, click ok to continue ...", vbInformation, "Information"
End Sub

Private Sub cmdclose_Click()
    OBJ.Open dsn
    SQL = "delete from am_nomax"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            OBJ.Open dsn
            SQL = "insert into am_nomax (kode) values ('" & List1.List(i) & "')"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
    Next i
    
    Unload Me
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
    
    If txtkode = "" Or txtkodeT = "" Or txtkodeI = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkode = ""
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE am_option SET ac_supp = '" & txtkode & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "INSERT INTO am_option "
        SQL = SQL + "(c_type,c_id,ac_supp,ac_ppnsupp)"
        
        SQL = SQL + " VALUES "
        SQL = SQL + "('" & txtkodeT & "','" & txtkodeI & "','" & txtkode & "','')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Master Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode3 = hasil
    hasil = ""
    hasil1 = ""
    
    If txtkode3 = "" Or txtkodeT = "" Or txtkodeI = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkode3 = ""
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "UPDATE am_option SET ac_ppnsupp = '" & txtkode3 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "INSERT INTO am_option "
        SQL = SQL + "(c_type,c_id,ac_supp,ac_ppnsupp)"
        
        SQL = SQL + " VALUES "
        SQL = SQL + "('" & txtkodeT & "','" & txtkodeI & "','','" & txtkode3 & "')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
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

Private Sub cmdset_Click()
    txtkode1 = Trim(txtkode1)
    txtkode2 = Trim(txtkode2)
    txtnew = Trim(txtnew)
    
    If txtkode1 = "" Or txtkode2 = "" Or txtnew = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    'cek kalo kosong semua
    For h = 0 To 14
        If grid.TextMatrix(h, 0) <> "" Then int1 = int1 + 1
    Next h
    
    If int1 = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    'cari yg sama
    For h = 0 To 14
        If grid.TextMatrix(h, 0) = txtnew Then
            MsgBox "Kode already exist.", vbExclamation, "Warning"
            Exit Sub
        End If
    Next h
    'masukin yg baru ke grid
    For h = 0 To 14
        If grid.TextMatrix(h, 0) = "" Then
            grid.TextMatrix(h, 0) = txtnew
            Exit For
        End If
    Next h
    
    OBJ.Open dsn
    SQL = "delete from am_kode"
    Set RST = OBJ.Execute(SQL)
        
    grid.Row = 0
    Do While True
        If grid.TextMatrix(grid.Row, 0) <> "" Then
            SQL = "INSERT INTO AM_KODE"
            SQL = SQL + "(kode1"
            SQL = SQL + ",kode2"
            SQL = SQL + ",kode3)"
    
            SQL = SQL + "VALUES"
            SQL = SQL + "('" & txtkode1 & "'"
            SQL = SQL + ", '" & txtkode2 & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 0) & "')"
            Set RST = OBJ.Execute(SQL)
        End If
        
        If grid.Row = 14 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    List1.Clear
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!kode3
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    For i = 0 To List1.ListCount - 1
        OBJ.Open dsn
        SQL = "select * from am_nomax where kode='" & List1.List(i) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then List1.Selected(i) = True
        OBJ.Close
    Next i
        
    txtnew = ""
    
    MsgBox "Data saved, click ok to continue ...", vbInformation, "Information"
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='204' and b.kodeuser = '2" & kuser & "'"
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
    grid.RowHeightMin = 300
    grid.ColWidth(0) = 1335
    
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode1 = RST!kode1
        txtkode2 = RST!kode2
        grid.Row = 0
        Do While Not RST.EOF
            grid.TextMatrix(grid.Row, 0) = RST!kode3
            List1.AddItem RST!kode3
            
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    End If
    
    SQL = "select * from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkodeT = RST!c_type
        txtkodeI = RST!c_id
        txtkode = RST!ac_supp
        txtkode3 = RST!ac_ppnsupp
    End If
    OBJ.Close
    
    For i = 0 To List1.ListCount - 1
        OBJ.Open dsn
        SQL = "select * from am_nomax where kode='" & List1.List(i) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then List1.Selected(i) = True
        OBJ.Close
    Next i
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkode3.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtkode1_Change()
    lbl1 = txtkode1 & "001" & txtkode2
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 47 Or KeyAscii = 8 Or KeyAscii = 121 Or KeyAscii = 109 Or KeyAscii = 89 Or KeyAscii = 77) Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtkode2_Change()
    lbl1 = txtkode1 & "001" & txtkode2
End Sub

Private Sub txtkode2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii = 8) Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtkode3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkodeT.SetFocus
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

Private Sub txtnew_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
