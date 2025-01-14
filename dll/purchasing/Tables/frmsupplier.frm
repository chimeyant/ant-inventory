VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmsupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Supplier"
   ClientHeight    =   3180
   ClientLeft      =   5715
   ClientTop       =   5520
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcari 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6855
      TabIndex        =   34
      Top             =   2820
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.ListBox List1 
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
      Height          =   2565
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.OptionButton ops1 
      Caption         =   "Bahan Baku"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton ops2 
      Caption         =   "Umum"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Ya"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2490
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Price List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   9975
      Begin MSComCtl2.DTPicker date2 
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
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
         CustomFormat    =   "MMM yyyy"
         Format          =   136511491
         CurrentDate     =   38981
      End
      Begin VB.PictureBox uncheck 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         Picture         =   "frmsupplier.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox check 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         Picture         =   "frmsupplier.frx":034E
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox blank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin TDBText6Ctl.TDBText txtket 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         Caption         =   "frmsupplier.frx":0630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsupplier.frx":069C
         Key             =   "frmsupplier.frx":06BA
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
         MaxLength       =   50
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
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   225
         Left            =   4440
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   397
         Calculator      =   "frmsupplier.frx":06F6
         Caption         =   "frmsupplier.frx":0716
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsupplier.frx":0782
         Keys            =   "frmsupplier.frx":07A0
         Spin            =   "frmsupplier.frx":07E2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000);0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         MergeCells      =   4
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin Chameleon.chameleonButton cmdsavelist 
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Save Price List"
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
         MICON           =   "frmsupplier.frx":080A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdclearlist 
         Height          =   375
         Left            =   8400
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Clear List"
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
         MICON           =   "frmsupplier.frx":0B24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblkode 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4320
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblsat 
         Caption         =   "Nama Satuan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label lblbar 
         Caption         =   "Nama Barang :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   2550
         Width           =   4095
      End
   End
   Begin VB.TextBox txtalamat1 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   3
      Top             =   750
      Width           =   3975
   End
   Begin VB.TextBox txtkontak 
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
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtfax 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox txtelp 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtalamat 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtnama 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2640
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
      MICON           =   "frmsupplier.frx":0E3E
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
      Left            =   3840
      TabIndex        =   11
      Top             =   2640
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
      MICON           =   "frmsupplier.frx":1158
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
      Left            =   2880
      TabIndex        =   10
      Top             =   2640
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
      MICON           =   "frmsupplier.frx":1472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdprice 
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Price List"
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
      MICON           =   "frmsupplier.frx":178C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Category Supplier"
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
      TabIndex        =   33
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Pengusaha kena Pajak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Contact Person"
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
      TabIndex        =   23
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Faxsimile"
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
      TabIndex        =   22
      Top             =   1830
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Telephone"
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
      TabIndex        =   21
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat Supplier"
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
      TabIndex        =   20
      Top             =   510
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Supplier"
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
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow, str21, str99 As String

Private Sub cmdadd_Click()
    OBJ.Open dsn
    SQL = "select * from am_pohdr"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "User can not add supplier.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If txtnama = "" Or txtalamat = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtnama = Trim(txtnama)
    OBJ.Open dsn
    SQL = "select namasupp from am_supplier where namasupp = '" & txtnama & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnama.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    history
    lblkode = str21

    OBJ.Open dsn
    SQL = "select kodesupp from am_supplier where kodesupp = '" & lblkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        history
        lblkode = str21

        GoTo jump98
        Exit Sub
    End If
    OBJ.Close

jump98:
    
    OBJ.Open dsn
    SQL = "INSERT INTO AM_Supplier"
    SQL = SQL + "(kodeSupp"
    SQL = SQL + ",NamaSupp"
    SQL = SQL + ",AlamatSupp1"
    SQL = SQL + ",AlamatSupp2"
    SQL = SQL + ",telpsupp"
    SQL = SQL + ",faxsupp"
    SQL = SQL + ",category"
    SQL = SQL + ",wp"
    SQL = SQL + ",contactperson)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + "('" & lblkode & "'"
    SQL = SQL + ", '" & txtnama & "'"
    SQL = SQL + ", '" & txtalamat & "'"
    SQL = SQL + ", '" & txtalamat1 & "'"
    SQL = SQL + ", '" & txtelp & "'"
    SQL = SQL + ", '" & txtfax & "'"
    If ops1.Value = True Then SQL = SQL + ", '1'" Else SQL = SQL + ", '2'"
    SQL = SQL + ", '" & chk1.Value & "'"
    SQL = SQL + ", '" & txtkontak & "')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtnama = ""
    txtalamat = ""
    txtalamat1 = ""
    txtelp = ""
    txtfax = ""
    txtkontak = ""
    lblkode = ""
    chk1.Value = 0
    ops1.Value = True
    txtnama.SetFocus
End Sub

Private Sub cmdclearlist_Click()
    grid.Clear
    grid.Rows = 2
    
    grid.TextMatrix(0, 1) = "Barang"
    grid.TextMatrix(0, 2) = "NamaBarang"
    grid.TextMatrix(0, 3) = "Satuan"
    grid.TextMatrix(0, 4) = "Harga"
    grid.TextMatrix(0, 5) = "Curr"
    grid.TextMatrix(0, 6) = "Tanggal"
    grid.TextMatrix(0, 7) = "Keterangan"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000
    grid.ColWidth(3) = 600
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 3000
    
    grid.RowHeightMin = 300
    
    lblbar = "Nama Barang :"
    lblsat = "Nama Satuan :"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub


Private Sub cmdprice_Click()
    If frmsupplier.Width = 5985 Then
        frmsupplier.Width = 10305
        frmsupplier.Height = 6975
        
        Me.Top = ((Screen.Height - Me.Height) / 2) - 800
        Me.Left = (Screen.Width - Me.Width) / 2
    Else
        frmsupplier.Width = 5985
        frmsupplier.Height = 3615
        
        Me.Top = ((Screen.Height - Me.Height) / 2)
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
End Sub

Private Sub cmdsavelist_Click()
    If lblkode = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        If grid.TextMatrix(grid.Row, 4) = "0.000" Or grid.TextMatrix(grid.Row, 5) = "" Or grid.TextMatrix(grid.Row, 6) = "" Then
            MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    If MsgBox("The Old Price will updated with The New Price, are you sure want to continue ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "delete from am_price where kodesupp = '" & lblkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        If grid.TextMatrix(grid.Row, 4) <> "0.000" Then
            SQL = "INSERT INTO AM_price"
            SQL = SQL + "(kodesupp"
            SQL = SQL + ",kodebarang"
            SQL = SQL + ",kodesatuan"
            SQL = SQL + ",keterangan"
            SQL = SQL + ",kodecurr"
            SQL = SQL + ",lastupdate"
            SQL = SQL + ",price)"
    
            SQL = SQL + "VALUES"
            SQL = SQL + "('" & lblkode & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 1) & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 3) & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 7) & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 5) & "'"
            SQL = SQL + ", convert(datetime,'" & Format(grid.TextMatrix(grid.Row, 6), "MM/01/yyyy") & "')"
            SQL = SQL + ", convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 4), "general number")) & "'))"
            Set RST = OBJ.Execute(SQL)
        End If
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    MsgBox "Price is set, click Ok to continue...", vbInformation, "Information"
    cmdclearlist_Click
    cmdclear_Click
End Sub

Private Sub date2_CloseUp()
    grid.TextMatrix(posrow, 6) = Format(date2, "MMMM yyyy")
    
    grid.SetFocus
    grid.Row = posrow
    date2.Visible = False
End Sub

Private Sub date2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date2.Visible = False
    If KeyCode = 13 Then
        grid.TextMatrix(posrow, 6) = Format(date2, "MMMM yyyy")
        
        grid.SetFocus
        grid.Row = posrow
        date2.Visible = False
    End If
End Sub

Private Sub date2_LostFocus()
    date2.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then List1.Visible = False
End Sub

Private Sub Form_Load()
    txtnama.ToolTipText = "max length = " & txtnama.MaxLength
    txtalamat.ToolTipText = "max length = " & txtalamat.MaxLength
    txtalamat1.ToolTipText = "max length = " & txtalamat1.MaxLength
    txtkontak.ToolTipText = "max length = " & txtkontak.MaxLength
    txtelp.ToolTipText = "max length = " & txtelp.MaxLength
    txtfax.ToolTipText = "max length = " & txtfax.MaxLength
    
    grid.TextMatrix(0, 1) = "Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Satuan"
    grid.TextMatrix(0, 4) = "Harga"
    grid.TextMatrix(0, 5) = "Curr"
    grid.TextMatrix(0, 6) = "Tanggal"
    grid.TextMatrix(0, 7) = "Keterangan"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000
    grid.ColWidth(3) = 600
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 3000
    
    
    grid.RowHeightMin = 300
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblbar = "Nama Barang : " & RST!namabarang
    Else
        lblbar = "Nama Barang : "
    End If
            
    SQL = "select * from am_apunit where kodesatuan = '" & grid.TextMatrix(grid.Row, 3) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsat = "Nama Satuan : " & RST!namasatuan
    Else
        lblsat = "Nama Satuan : "
    End If
    OBJ.Close
    
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
            
            carisql1 = "select kodebarang, kodesatuan, namabarang from am_apitemmst"
            namatabel = "Bahan Baku "
    
            frmsearch.Show vbModal
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 5, 7
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                        
            If date2.Visible = True Then Exit Sub
            
            date2.Width = grid.ColWidth(grid.Col) - 40
            If grid.TextMatrix(grid.Row, grid.Col) <> "" Then date2 = grid.TextMatrix(grid.Row, 6)
            date2.Left = grid.Left + grid.CellLeft
            date2.Top = grid.Top + grid.CellTop + 20
            date2.Visible = True
            date2 = Date
            date2.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblbar = "Nama Barang : " & RST!namabarang
    Else
        lblbar = "Nama Barang : "
    End If
            
    SQL = "select * from am_apunit where kodesatuan = '" & grid.TextMatrix(grid.Row, 3) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsat = "Nama Satuan : " & RST!namasatuan
    Else
        lblsat = "Nama Satuan : "
    End If
    OBJ.Close
    
    Select Case grid.Col
    Case 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 5, 7
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 6
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                        
        If date2.Visible = True Then Exit Sub
        
        posrow = grid.Row
            
        date2.Width = grid.ColWidth(grid.Col) - 40
        If grid.TextMatrix(grid.Row, grid.Col) <> "" Then date2 = grid.TextMatrix(grid.Row, 6)
        date2.Left = grid.Left + grid.CellLeft
        date2.Top = grid.Top + grid.CellTop + 20
        date2.Visible = True
        date2 = Date
        date2.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
    Case 1
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 1) = hasil Then
                MsgBox "Item Alraedy Exist", vbExclamation, "Warning"
                
                hasil = ""
                hasil1 = ""
                hasil2 = ""
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
        
        grid.Row = posrow
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = hasil
        grid.Col = 2
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 3) = hasil1
        hasil = ""
        hasil1 = ""
        hasil2 = ""
        
        OBJ.Open dsn
        SQL = "select * from am_apitemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblbar = "Nama Barang : " & RST!namabarang
            grid.TextMatrix(grid.Row, 4) = "0.000"
            
            SQL = "select * from am_apunit where kodesatuan = '" & grid.TextMatrix(grid.Row, 3) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then lblsat = "Nama Satuan : " & RST!namasatuan
            
            SetRow grid.Row, True
            grid.SetFocus
            grid.Col = 2
            
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            lblbar = "Nama Barang : "
            lblsat = "Nama Satuan : "
            
            MsgBox "Item Not Found", vbExclamation, "Warning"
        End If
        OBJ.Close
    Case 5
        grid.Row = posrow
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 5) = hasil
        hasil = ""
        hasil1 = ""
        hasil2 = ""
        
        OBJ.Open dsn
        SQL = "select kdkurs from gl_kurs where kdkurs = '" & grid.TextMatrix(grid.Row, 5) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.SetFocus
        Else
            grid.TextMatrix(grid.Row, 5) = ""
            MsgBox "Currency Not Found", vbExclamation, "Warning"
        End If
        OBJ.Close
    End Select
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub

Private Sub List1_DblClick()
    txtnama = List1.text
    txtnama = Trim(txtnama)
    
    OBJ.Open dsn
    SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblkode = RST!KodeSupp
        txtalamat = RST!alamatsupp1
        txtalamat1 = RST!alamatsupp2
        txtelp = RST!telpsupp
        txtfax = RST!faxsupp
        txtkontak = RST!contactperson
        chk1.Value = RST!wp
        If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
    End If
    OBJ.Close
    List1.Visible = False
    
    carilist
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtnama = List1.text
        txtnama = Trim(txtnama)
        OBJ.Open dsn
        SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblkode = RST!KodeSupp
            txtalamat = RST!alamatsupp1
            txtalamat1 = RST!alamatsupp2
            txtelp = RST!telpsupp
            txtfax = RST!faxsupp
            txtkontak = RST!contactperson
            chk1.Value = RST!wp
            If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
        End If
        OBJ.Close
        List1.Visible = False
        
        carilist
    End If
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalamat1.SetFocus
End Sub

Private Sub txtAlamat1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkontak.SetFocus
End Sub

Private Sub txtcari_Change()
If txtcari <> "" Then
    OBJ.Open dsn
    SQL = "select a.kodesupp,a.kodebarang,b.NamaBarang,a.kodesatuan,a.price,a.kodecurr,a.lastupdate,a.keterangan from am_price a"
    SQL = SQL + " left join am_apitemmst b on a.kodebarang = b.KodeBarang Where a.kodesupp = '" & lblkode & "' and b.NamaBarang like '" & txtcari & "%'"
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    
    grid.TextMatrix(0, 1) = "Barang"
    grid.TextMatrix(0, 2) = "NamaBarang"
    grid.TextMatrix(0, 3) = "Satuan"
    grid.TextMatrix(0, 4) = "Harga"
    grid.TextMatrix(0, 5) = "Curr"
    grid.TextMatrix(0, 6) = "Tanggal"
    grid.TextMatrix(0, 7) = "Keterangan"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000
    grid.ColWidth(3) = 600
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 3000
    
   ' grid.RowHeightMin = 300
    OBJ.Close
End If
End Sub

Private Sub txtelp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtfax.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ops1.SetFocus
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 27 Then
        txtket_LostFocus
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 5
                grid.Row = posrow
                
                grid.SetFocus
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 5) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "select * from gl_kurs where kdkurs = '" & grid.TextMatrix(grid.Row, 5) & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    OBJ.Close
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 5) = ""
                    txtket = ""
                    
                    carisql1 = "select kdkurs, nmkurs from gl_kurs"
                    namatabel = "Currency"
   
                    frmsearch.Show vbModal
                End If
                grid.Col = 4
            Case 7
                grid.Row = posrow
                
                grid.SetFocus
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 7) = txtket
                txtket = ""
                txtket.Visible = False
                grid.Col = 7
        End Select
    ElseIf KeyAscii = 27 Then
        txtket = ""
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtKontak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtelp.SetFocus
End Sub

Private Sub txtnama_Change()
    cari
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then List1.Visible = False
    If KeyAscii = 13 Then
        List1.Visible = False
        txtalamat.SetFocus
        
        OBJ.Open dsn
        txtnama = Trim(txtnama)
        SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblkode = RST!KodeSupp
            txtalamat = RST!alamatsupp1
            txtalamat1 = RST!alamatsupp2
            txtelp = RST!telpsupp
            txtfax = RST!faxsupp
            txtkontak = RST!contactperson
            chk1.Value = RST!wp
            If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
            
            OBJ.Close
            carilist
            Exit Sub
        Else
            lblkode = ""
            txtalamat = ""
            txtalamat1 = ""
            txtelp = ""
            txtfax = ""
            txtkontak = ""
            chk1.Value = 0
            ops1.Value = True
        End If
        OBJ.Close
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "##,###,###,##0.0000")
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub cari()
    If txtnama = "" Then
        List1.Visible = False
        Exit Sub
    End If
    List1.Clear
    
    OBJ.Open dsn
    SQL = "select namasupp from am_supplier where namasupp like '" & txtnama & "%' order by namasupp"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            List1.AddItem RST!namasupp
            RST.MoveNext
        Loop
        List1.Visible = True
    Else
        List1.Visible = False
    End If
    OBJ.Close
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

Private Sub carilist()
    grid.Clear
    grid.Rows = 2
    
    grid.TextMatrix(0, 1) = "Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Satuan"
    grid.TextMatrix(0, 4) = "Harga"
    grid.TextMatrix(0, 5) = "Curr"
    grid.TextMatrix(0, 6) = "Tanggal"
    grid.TextMatrix(0, 7) = "Keterangan"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000
    grid.ColWidth(3) = 600
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 3000
    
    grid.RowHeightMin = 300
    
    If lblkode = "" Then Exit Sub
    
    OBJ.Open dsn
    'SQL = "select kodebarang,kodesatuan,price,kodecurr,lastupdate,keterangan from am_price where kodesupp = '" & lblkode & "' order by kodebarang"
    SQL = "select a.kodebarang,b.NamaBarang,a.kodesatuan,a.price,a.kodecurr,a.lastupdate,a.keterangan from am_price a"
    SQL = SQL + " left join am_apitemmst b on a.kodebarang = b.KodeBarang Where a.kodesupp = '" & lblkode & "' order by b.NamaBarang"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        grid.Row = 1
        Do While Not RST.EOF
            SetRow grid.Row, True
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 4) = Format(RST!price, "##,###,###,##0.0000")
            grid.TextMatrix(grid.Row, 5) = RST!kodecurr
            grid.TextMatrix(grid.Row, 6) = Format(RST!lastupdate, "MMMM yyyy")
            grid.TextMatrix(grid.Row, 7) = RST!keterangan
            
            SetRow grid.Row, True
                        
            RST.MoveNext
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
        Loop
    End If
    OBJ.Close
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub history()
    OBJ.Open dsn
    SQL = "select top 1 kodesupp from am_supplier where kodesupp like '0%' order by kodesupp desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str99 = RST!KodeSupp Else str99 = 0
    
    str99 = str99 + 1
    
    If Len(str99) = 1 Then str21 = "0000" & str99
    If Len(str99) = 2 Then str21 = "000" & str99
    If Len(str99) = 3 Then str21 = "00" & str99
    If Len(str99) = 4 Then str21 = "0" & str99
    If Len(str99) = 5 Then str21 = str99
    OBJ.Close
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
