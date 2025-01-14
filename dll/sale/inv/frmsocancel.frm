VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsocancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close/Cancel Sales Order"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
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
   ScaleHeight     =   6135
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cancel Description"
      Height          =   1050
      Left            =   120
      TabIndex        =   18
      Top             =   2925
      Visible         =   0   'False
      Width           =   6495
      Begin TDBText6Ctl.TDBText txtket2 
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmsocancel.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsocancel.frx":0065
         Key             =   "frmsocancel.frx":0083
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
      Begin Chameleon.chameleonButton cmdexit2 
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Exit"
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
         MICON           =   "frmsocancel.frx":00C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsave2 
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Save Description"
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
         MICON           =   "frmsocancel.frx":03E1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdrefresh2 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
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
         MICON           =   "frmsocancel.frx":06FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   5055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8916
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Close Description"
      Height          =   1725
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   6495
      Begin TDBText6Ctl.TDBText txtket 
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmsocancel.frx":0A15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsocancel.frx":0A7A
         Key             =   "frmsocancel.frx":0A98
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
      Begin Chameleon.chameleonButton cmdexit 
         Height          =   375
         Left            =   5520
         TabIndex        =   17
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Exit"
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
         MICON           =   "frmsocancel.frx":0ADC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsavedesc 
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Save Description"
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
         MICON           =   "frmsocancel.frx":0DF6
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
         Height          =   5055
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8916
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
      Begin Chameleon.chameleonButton cmdrefresh 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
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
         MICON           =   "frmsocancel.frx":1110
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
   Begin VB.CheckBox Check2 
      Caption         =   "Check All"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All Purchase Order"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   4110
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   5640
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
      MICON           =   "frmsocancel.frx":142A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpost 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "Submit"
      Top             =   2760
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
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
      MICON           =   "frmsocancel.frx":1744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   134742019
      CurrentDate     =   37694
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   134742019
      CurrentDate     =   37694
   End
   Begin Chameleon.chameleonButton cmdclosedesc 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close Description"
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
      MICON           =   "frmsocancel.frx":1A5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpost1 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   "Submit"
      Top             =   5040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
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
      MICON           =   "frmsocancel.frx":1D78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3600
      Width           =   2895
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel Description"
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
      MICON           =   "frmsocancel.frx":2092
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
      Left            =   2520
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsocancel.frx":23AC
      Caption         =   "frmsocancel.frx":23CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsocancel.frx":2431
      Keys            =   "frmsocancel.frx":244F
      Spin            =   "frmsocancel.frx":2499
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   225
      Left            =   2520
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsocancel.frx":24C1
      Caption         =   "frmsocancel.frx":24E1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsocancel.frx":2546
      Keys            =   "frmsocancel.frx":2564
      Spin            =   "frmsocancel.frx":25AE
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "From                                            To"
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   510
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   25
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "frmsocancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim SP As New ADODB.Command
Dim vsp(0) As Variant

Dim i, j As Integer
Dim posrow, poscol As String

Private Sub chameleonButton1_Click()
    Frame2.Visible = True
End Sub

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        date1.Enabled = True
        date2.Enabled = True
        date1.Value = Date
        date2.Value = Date
    Else
        date1.Enabled = False
        date2.Enabled = False
        
        List1.Clear
        List2.Clear
        List3.Clear
        
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '1' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List2.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '2' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List3.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        If List1.ListCount >= 1 Then
            For i = 0 To List1.ListCount - 1
                List1.Selected(i) = True
            Next i
        End If
    Else
        If List1.ListCount >= 1 Then
            For i = 0 To List1.ListCount - 1
                List1.Selected(i) = False
            Next i
        End If
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdclosedesc_Click()
    Frame1.Visible = True
End Sub

Private Sub cmdexit_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdexit2_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdpost_Click()
    If List1.ListCount = 0 Then Exit Sub
        
    j = 0
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then j = j + 1
    Next i
    
    If j = 0 Then
        MsgBox "To close PO, user must select/check at least one PO.", vbExclamation, "Information"
        Exit Sub
    End If
    
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then
            OBJ.Open dsn
            SQL = "update am_pohdr set flag = '1' where nopo = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            List2.AddItem List1.List(i - 1)
        End If
    Next i
    
    List1.Clear
    If Check1.Value = 1 Then
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
    Else
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    Check2.Value = 0
    
    MsgBox "Closing Complete.", vbInformation, "Information"
End Sub

Private Sub cmdpost1_Click()
    If List1.ListCount = 0 Then Exit Sub
        
    j = 0
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then j = j + 1
    Next i
    
    If j = 0 Then
        MsgBox "To cancel PO, user must select/check at least one PO.", vbExclamation, "Information"
        Exit Sub
    End If
    
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then
            OBJ.Open dsn
            SQL = "update am_pohdr set flag = '2' where nopo = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            List3.AddItem List1.List(i - 1)
        End If
    Next i
    
    List1.Clear
    If Check1.Value = 1 Then
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
    Else
        OBJ.Open dsn
        SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!nopo
        
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    Check2.Value = 0
    
    MsgBox "Canceling Complete.", vbInformation, "Information"
End Sub

Private Sub cmdrefresh_Click()
    grid.Clear
    grid.TextMatrix(0, 0) = "P.Order"
    grid.TextMatrix(0, 1) = "Keterangan"
    grid.ColWidth(0) = 1500
    grid.ColWidth(1) = 4000
    grid.ColWidth(2) = 0
    grid.RowHeightMin = 300
    grid.Rows = 2
    
    grid.Row = 1
    OBJ.Open dsn
    SQL = "select distinct a.nopo,isnull(b.keterangan,'')'ket',substring(a.nopo,11,3)'aa' from am_pohdr a left join am_podrop b on a.nopo=b.nopo where a.flag = '1' order by substring(a.nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 0
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 0) = RST!nopo
        grid.TextMatrix(grid.Row, 1) = RST!ket
        grid.TextMatrix(grid.Row, 2) = RST!aa
        
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdrefresh2_Click()
    grid2.Clear
    grid2.TextMatrix(0, 0) = "P.Order"
    grid2.TextMatrix(0, 1) = "Keterangan"
    grid2.ColWidth(0) = 1500
    grid2.ColWidth(1) = 4000
    grid2.ColWidth(2) = 0
    grid2.RowHeightMin = 300
    grid2.Rows = 2
    
    grid2.Row = 1
    OBJ.Open dsn
    SQL = "select distinct a.nopo,isnull(b.keterangan,'')'ket',substring(a.nopo,11,3)'aa' from am_pohdr a left join am_podrop b on a.nopo=b.nopo where a.flag = '2' order by substring(a.nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid2.Col = 0
        grid2.CellAlignment = 1
        grid2.TextMatrix(grid2.Row, 0) = RST!nopo
        grid2.TextMatrix(grid2.Row, 1) = RST!ket
        grid2.TextMatrix(grid2.Row, 2) = RST!aa
        
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdsave2_Click()
    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_podrop where closecancel = '2'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        
        If grid2.TextMatrix(grid2.Row, 1) <> "" Then
            OBJ.Open dsn
            SQL = "insert into am_podrop ("
            SQL = SQL + "nopo,"
            SQL = SQL + "closecancel,"
            SQL = SQL + "keterangan)"

            SQL = SQL + " values("
            SQL = SQL + "'" & grid2.TextMatrix(grid2.Row, 0) & "',"
            SQL = SQL + "'2',"
            SQL = SQL + "'" & grid2.TextMatrix(grid2.Row, 1) & "')"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
        grid2.Row = grid2.Row + 1
    Loop
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdsavedesc_Click()
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_podrop where closecancel = '1'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Do
        
        If grid.TextMatrix(grid.Row, 1) <> "" Then
            OBJ.Open dsn
            SQL = "insert into am_podrop ("
            SQL = SQL + "nopo,"
            SQL = SQL + "closecancel,"
            SQL = SQL + "keterangan)"

            SQL = SQL + " values("
            SQL = SQL + "'" & grid.TextMatrix(grid.Row, 0) & "',"
            SQL = SQL + "'1',"
            SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "')"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
        grid.Row = grid.Row + 1
    Loop
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub date1_Change()
    List1.Clear
    List2.Clear
    List3.Clear
    
    If Check1.Value = 1 Then Exit Sub
    If date1 > date2 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '1' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '2' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List3.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub date2_Change()
    List1.Clear
    List2.Clear
    List3.Clear
    
    If Check1.Value = 1 Then Exit Sub
    If date1 > date2 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '0' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '1' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '2' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List3.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='154' and b.kodeuser = '2" & kuser & "'"
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
   
    grid.TextMatrix(0, 0) = "P.Order"
    grid.TextMatrix(0, 1) = "Keterangan"
    grid.ColWidth(0) = 1500
    grid.ColWidth(1) = 4000
    grid.ColWidth(2) = 0
    grid.RowHeightMin = 300
    
    grid2.TextMatrix(0, 0) = "P.Order"
    grid2.TextMatrix(0, 1) = "Keterangan"
    grid2.ColWidth(0) = 1500
    grid2.ColWidth(1) = 4000
    grid2.ColWidth(2) = 0
    grid2.RowHeightMin = 300
    
    OBJ.Open dsn
    SQL = "DELETE FROM AM_cek1 WHERE kompi='" & GetTheComputerName & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_check1"
    vsp(0) = GetTheComputerName
    SP.Execute , vsp
    Set SP = Nothing
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_cek1 WHERE kompi='" & GetTheComputerName & "' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nopo
   
        RST.MoveNext
    Loop
    OBJ.Close
        
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '1' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT nopo FROM AM_pohdr WHERE flag = '2' order by substring(nopo,11,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List3.AddItem RST!nopo
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Function tanggal1()
      tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
      tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 1
            If grid.TextMatrix(grid.Row, 0) = "" Then Exit Sub

            If txtket.Visible = True Then Exit Sub

            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub

        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    End Select
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    
    posrow = grid2.Row
    poscol = grid2.Col
    Select Case grid2.Col
        Case 1
            If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub

            If txtket2.Visible = True Then Exit Sub

            txtket2.Width = grid2.ColWidth(grid2.Col) - 40
            txtket2 = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtket2.Left = grid2.Left + grid2.CellLeft
            txtket2.Top = grid2.Top + grid2.CellTop + 20
            txtket2.Visible = True
            txtket2.SetFocus
    End Select
End Sub

Private Sub grid2_EnterCell()
    Select Case grid2.Col
    Case 1
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        If txtket2.Visible = True Then Exit Sub

        posrow = grid2.Row
        poscol = grid2.Col
        txtket2.Width = grid2.ColWidth(grid2.Col) - 40
        txtket2 = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtket2.Left = grid2.Left + grid2.CellLeft
        txtket2.Top = grid2.Top + grid2.CellTop + 20
        txtket2.Visible = True
        txtket2.SetFocus
    End Select
End Sub

Private Sub grid2_Scroll()
    txtket2.Visible = False
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0

    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 1
                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtket2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0

    If KeyAscii = 13 Then
        Select Case grid2.Col
            Case 1
                grid2.SetFocus
                grid2.Col = 1
                grid2.CellAlignment = 1
                grid2.TextMatrix(grid2.Row, 1) = txtket2
                txtket2 = ""
                txtket2.Visible = False
        End Select
    ElseIf KeyAscii = 27 Then
        txtket2.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket2_LostFocus()
    txtket2.Visible = False
End Sub
