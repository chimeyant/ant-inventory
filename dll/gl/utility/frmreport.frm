VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmreport 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmreport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   22
      Top             =   1935
      Width           =   6015
      Begin VB.OptionButton opsbukbes 
         Caption         =   "Buku Besar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton opscash 
         Caption         =   "Cash Flow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton opsbalance 
         Caption         =   "Balance Sheet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opsincome 
         Caption         =   "Income Statement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox txtdesc1 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtitle 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5106
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Group"
      TabPicture(0)   =   "frmreport.frx":2372
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdaddgroup"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdupdategroup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detail Group"
      TabPicture(1)   =   "frmreport.frx":238E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblgroupno"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdupdatedetail"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdaddetail"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "grid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin Chameleon.chameleonButton cmdupdategroup 
         Height          =   495
         Left            =   -68160
         TabIndex        =   10
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   9
         TX              =   "Edit Group"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         MICON           =   "frmreport.frx":23AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdaddgroup 
         Height          =   495
         Left            =   -68160
         TabIndex        =   9
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   9
         TX              =   "Add Group"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         MICON           =   "frmreport.frx":26C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   2655
         Left            =   -74520
         TabIndex        =   8
         Top             =   120
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   14
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   2295
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin Chameleon.chameleonButton cmdaddetail 
         Height          =   495
         Left            =   6840
         TabIndex        =   12
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   9
         TX              =   "Add Detail"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         MICON           =   "frmreport.frx":29DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdupdatedetail 
         Height          =   495
         Left            =   6840
         TabIndex        =   13
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   9
         TX              =   "Edit Detail"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         MICON           =   "frmreport.frx":2CF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblgroupno 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   120
         Width           =   7575
      End
   End
   Begin TDBText6Ctl.TDBText txtreportcode 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Caption         =   "frmreport.frx":3012
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmreport.frx":307E
      Key             =   "frmreport.frx":309C
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
      Left            =   360
      TabIndex        =   26
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Report Code"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":30D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":33F2
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
      TabIndex        =   15
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":370C
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
      TabIndex        =   16
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":3A26
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
      TabIndex        =   17
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":3D40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdcash 
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Cash Flow Column"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmreport.frx":405A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbltotdetail 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Detail : 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lbltotgroup 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Group : 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Layout Report"
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
      TabIndex        =   24
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
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
      Width           =   2655
   End
   Begin VB.Label lblapor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Type Laporan"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   20
      Top             =   1950
      Width           =   1395
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   19
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Report Title"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   18
      Top             =   2310
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New Connection
Dim RST As New Recordset
Dim SQL As String

Dim OBJ1 As New Connection
Dim RST1 As New Recordset
Dim SQL1 As String

Private Sub cmdcash_Click()
    If txtreportcode = "" Then Exit Sub
    setup2 = lblapor
    frmreportcash.Show 1
End Sub

Private Sub cmdelete_Click()
    If txtreportcode = "" Or txtdesc1 = "" Or txtitle = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete Form, Group, And Detail Group ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "DELETE FROM gl_rforms WHERE form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
            
    SQL = "DELETE FROM gl_dforms WHERE form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
        
    SQL = "DELETE FROM gl_gforms WHERE form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE FROM gl_cforms WHERE form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
        
    MsgBox "Data Group And Detail Group Report Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdSave_Click()
    If txtreportcode = "" Or txtdesc1 = "" Or txtitle = "" Or (opsbalance.Value = False And opsincome.Value = False And opscash.Value = False And opsbukbes = False) Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtreportcode)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtreportcode = Trim(txtreportcode)
    
    OBJ.Open dsn
    SQL = "select * from gl_rforms where form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        SQL = "insert into gl_rforms"
        SQL = SQL + "(form_no"
        SQL = SQL + ",description"
        SQL = SQL + ",report_type"
        SQL = SQL + ",report_title)"
    
        SQL = SQL + "VALUES"
        SQL = SQL + "('" & txtreportcode & "'"
        SQL = SQL + ", '" & txtdesc1 & "'"
        SQL = SQL + ", '" & lblapor & "'"
        SQL = SQL + ", '" & txtitle & "')"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "UPDATE gl_rforms SET "
        SQL = SQL + "description = '" & txtdesc1 & "',"
        SQL = SQL + "report_title = '" & txtitle & "'"
        SQL = SQL + "WHERE form_no =  '" & txtreportcode & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    If grid1.Rows > 2 Then
        OBJ.Open dsn
        SQL = "delete from gl_gforms where form_no = '" & txtreportcode & "'"
        Set RST = OBJ.Execute(SQL)
        
        grid1.Row = 1
        Do While True
            If grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Do
            
            SQL = "insert into gl_gforms"
            SQL = SQL + "(form_no"
            SQL = SQL + ",group_no"
            SQL = SQL + ",type_ac"
            SQL = SQL + ",description"
            SQL = SQL + ",space_after"
            SQL = SQL + ",print_cash"
            SQL = SQL + ",print_coloum"
            SQL = SQL + ",print_mode"
            SQL = SQL + ",print_sign)"
            
            SQL = SQL + "VALUES"
            SQL = SQL + "('" & txtreportcode & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 0) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 1) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 2) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 3) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 13) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 4) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 5) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 6) & "')"
            Set RST = OBJ.Execute(SQL)
            
            grid1.Row = grid1.Row + 1
        Loop
        OBJ.Close
    End If
    
    If grid1.Rows > 2 Then
        OBJ.Open dsn
        SQL = "delete from gl_dforms where form_no = '" & txtreportcode & "'"
        Set RST = OBJ.Execute(SQL)
            
        grid1.Row = 1
        Do While True
            If grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Do
            
            For x = 0 To 100
                If myarray(x, 10, 1) = grid1.TextMatrix(grid1.Row, 0) Then
                    z = 1
                    Do While True
                        If myarray(x, 0, z) = "" Then Exit Do
                        
                        SQL = "insert into gl_dforms"
                        SQL = SQL + "(form_no"
                        SQL = SQL + ",group_no"
                        SQL = SQL + ",line_no"
                        SQL = SQL + ",acc_no1"
                        SQL = SQL + ",acc_no2"
                        SQL = SQL + ",acc_no3"
                        SQL = SQL + ",acc_no4"
                        SQL = SQL + ",acc_no5"
                        SQL = SQL + ",acc_no6"
                        SQL = SQL + ",acc_no7"
                        SQL = SQL + ",acc_no8"
                        SQL = SQL + ",acc_no9)"
            
                        SQL = SQL + "VALUES"
                        SQL = SQL + "('" & txtreportcode & "'"
                        SQL = SQL + ", '" & myarray(x, 10, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 0, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 1, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 2, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 3, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 4, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 5, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 6, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 7, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 8, z) & "'"
                        SQL = SQL + ", '" & myarray(x, 9, z) & "')"
                        Set RST = OBJ.Execute(SQL)
            
                        z = z + 1
                    Loop
                    GoTo jump
                End If
            Next x
jump:
            grid1.Row = grid1.Row + 1
        Loop
        OBJ.Close
    End If
    MsgBox "Form, Group and Detail Report Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select Form_no, description from gl_rforms"
    namatabel = "Report Code"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtreportcode = hasil
    hasil = ""
    hasil1 = ""
    cariform
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub grid1_Click()
    setup3 = ""
    If grid1.MouseRow > 0 Then
        If grid1.TextMatrix(grid1.Row, 0) <> "" Then
            setup3 = grid1.Row
            setup4 = grid1.TextMatrix(grid1.Row, 1)
        End If
    End If
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow > 0 Then
        If grid2.TextMatrix(grid2.Row, 0) <> "" Then setup1 = grid2.Row
    End If
End Sub

Private Sub opsbalance_Click()
    lblapor = 1
    cmdcash.Enabled = False
    txtitle.SetFocus
End Sub

Private Sub opsbukbes_Click()
    lblapor = 4
    cmdcash.Enabled = False
    txtitle.SetFocus
End Sub

Private Sub opscash_Click()
    lblapor = 3
    cmdcash.Enabled = True
    txtitle.SetFocus
End Sub

Private Sub opsincome_Click()
    lblapor = 2
    cmdcash.Enabled = False
    txtitle.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        If txtreportcode = "" Or grid1.Rows = 2 Or setup3 = "" Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If (setup3 <> "" And (grid1.TextMatrix(setup3, 1) = "0") Or (setup3 <> "" And grid1.TextMatrix(setup3, 4) <> "1")) Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        
        'If (setup3 <> "" And (grid1.TextMatrix(setup3, 1) = "0") Or (setup3 <> "" And (grid1.TextMatrix(setup3, 4) > "2" Or grid1.TextMatrix(setup3, 4) < "1"))) Then
        cmdsave.Enabled = False
        cmdelete.Enabled = False
        cmdclear.Enabled = False
        cmdclose.Enabled = False
        
        Timer1.Interval = 1
    Else
        cmdsave.Enabled = True
        cmdelete.Enabled = True
        cmdclear.Enabled = True
        cmdclose.Enabled = True
    End If
End Sub

Private Sub cmdupdatedetail_Click()
    If txtreportcode = "" Or setup1 = "" Then Exit Sub
    setup2 = lblapor
    frmreportdetail.Show 1
    grid1.SetFocus
    grid1.Row = 1
    grid1.Col = 0
End Sub

Private Sub CmdUpdategroup_Click()
    If txtreportcode = "" Or setup3 = "" Then Exit Sub
    setup2 = lblapor
    frmreportgroup.Show 1
    grid1.SetFocus
    grid1.Row = 1
    grid1.Col = 0
End Sub

Private Sub Timer1_Timer()
    hapusdetail
    caridetail
    lbltotdetail = "Total Detail : " & grid2.Rows - 2
    grid2.SetFocus
    grid2.Row = 1
    grid2.Col = 0
    Timer1.Interval = 0
End Sub

Private Sub txtdesc1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtreportcode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtreportcode_LostFocus
End Sub

Private Sub txtreportcode_LostFocus()
    cariform
End Sub

Private Sub cariform()
    If txtreportcode = "" Then Exit Sub
    If txtreportcode.SelLength <> 0 Then Exit Sub
    hapusform
    hapusgroup
    hapusarray
    SSTab1.Tab = 0
    OBJ.Open dsn
    SQL = "select * from gl_rforms where form_no = '" & txtreportcode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc1 = RST!Description
        txtitle = RST!report_title
        lblapor = RST!report_type
        If lblapor = 1 Then opsbalance.Value = True
        If lblapor = 2 Then opsincome.Value = True
        If lblapor = 3 Then opscash.Value = True
        If lblapor = 4 Then opsbukbes.Value = True
        Frame1.Enabled = False
        OBJ.Close
        carigroup
        lbltotgroup = "Total Group : " & grid1.Rows - 2
        SSTab1.Tab = 1
        SSTab1.Tab = 0
        txtdesc1.SetFocus
        Exit Sub
    End If
    OBJ.Close
    txtdesc1.SetFocus
End Sub

Private Sub hapusform()
    txtdesc1 = ""
    txtitle = ""
    Frame1.Enabled = True
    opsbalance.Value = True
    opsincome.Value = False
End Sub

Private Sub hapusarray()
    For x = 0 To 100
        For z = 0 To 100
            For y = 0 To 10
                myarray(x, y, z) = ""
            Next y
        Next z
    Next x
End Sub

Private Sub hapusgroup()
    grid1.ColWidth(0) = 800
    grid1.ColWidth(1) = 0
    grid1.ColWidth(2) = 0
    grid1.ColWidth(3) = 0
    grid1.ColWidth(4) = 0
    grid1.ColWidth(5) = 0
    grid1.ColWidth(6) = 0
    grid1.ColWidth(7) = 1100
    grid1.ColWidth(8) = 1500
    grid1.ColWidth(9) = 1100
    grid1.ColWidth(10) = 1000
    grid1.ColWidth(11) = 1000
    grid1.ColWidth(12) = 500
    grid1.ColWidth(13) = 1000
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Do
        
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.TextMatrix(grid1.Row, 7) = ""
        grid1.TextMatrix(grid1.Row, 8) = ""
        grid1.TextMatrix(grid1.Row, 9) = ""
        grid1.TextMatrix(grid1.Row, 10) = ""
        grid1.TextMatrix(grid1.Row, 11) = ""
        grid1.TextMatrix(grid1.Row, 12) = ""
        grid1.TextMatrix(grid1.Row, 13) = ""
        
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    lbltotgroup = "Total Group : " & grid1.Rows - 2
End Sub

Private Sub carigroup()
    grid1.Row = 1
    OBJ.Open dsn
    SQL = "select * from gl_gforms where form_no = '" & txtreportcode & "' order by group_no asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 0) = RST!group_no
        grid1.TextMatrix(grid1.Row, 1) = RST!type_ac
        grid1.TextMatrix(grid1.Row, 2) = RST!Description
        grid1.TextMatrix(grid1.Row, 3) = RST!space_after
        grid1.TextMatrix(grid1.Row, 4) = RST!print_coloum
        grid1.TextMatrix(grid1.Row, 5) = RST!print_mode
        grid1.TextMatrix(grid1.Row, 6) = RST!print_sign
        
        If lblapor = 1 And grid1.TextMatrix(grid1.Row, 1) = "0" Then
            grid1.TextMatrix(grid1.Row, 7) = "Header Only"
        ElseIf lblapor = 1 And grid1.TextMatrix(grid1.Row, 1) = "1" Then
            grid1.TextMatrix(grid1.Row, 7) = "Assets"
        ElseIf lblapor = 1 And grid1.TextMatrix(grid1.Row, 1) = "2" Then
            grid1.TextMatrix(grid1.Row, 7) = "Liability"
        ElseIf lblapor = 1 And grid1.TextMatrix(grid1.Row, 1) = "3" Then
            grid1.TextMatrix(grid1.Row, 7) = "Capital"
        ElseIf lblapor = 1 And grid1.TextMatrix(grid1.Row, 1) = "4" Then
            grid1.TextMatrix(grid1.Row, 7) = "Income Summary"
        ElseIf lblapor = 2 And grid1.TextMatrix(grid1.Row, 1) = "0" Then
            grid1.TextMatrix(grid1.Row, 7) = "Header Only"
        ElseIf lblapor = 2 And grid1.TextMatrix(grid1.Row, 1) = "1" Then
            grid1.TextMatrix(grid1.Row, 7) = "Income"
        ElseIf lblapor = 2 And grid1.TextMatrix(grid1.Row, 1) = "2" Then
            grid1.TextMatrix(grid1.Row, 7) = "Expenses"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "0" Then
            grid1.TextMatrix(grid1.Row, 7) = "Header Only"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "1" Then
            grid1.TextMatrix(grid1.Row, 7) = "Assets"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "2" Then
            grid1.TextMatrix(grid1.Row, 7) = "Liability"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "3" Then
            grid1.TextMatrix(grid1.Row, 7) = "Capital"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "4" Then
            grid1.TextMatrix(grid1.Row, 7) = "Income"
        ElseIf lblapor = 3 And grid1.TextMatrix(grid1.Row, 1) = "5" Then
            grid1.TextMatrix(grid1.Row, 7) = "Expenses"
        ElseIf lblapor = 4 And grid1.TextMatrix(grid1.Row, 1) = "0" Then
            grid1.TextMatrix(grid1.Row, 7) = "Header Only"
        ElseIf lblapor = 4 And grid1.TextMatrix(grid1.Row, 1) = "1" Then
            grid1.TextMatrix(grid1.Row, 7) = "Assets"
        ElseIf lblapor = 4 And grid1.TextMatrix(grid1.Row, 1) = "2" Then
            grid1.TextMatrix(grid1.Row, 7) = "Liability"
        ElseIf lblapor = 4 And grid1.TextMatrix(grid1.Row, 1) = "3" Then
            grid1.TextMatrix(grid1.Row, 7) = "Capital"
        ElseIf lblapor = 4 And grid1.TextMatrix(grid1.Row, 1) = "4" Then
            grid1.TextMatrix(grid1.Row, 7) = "Income Summary"
        End If
        
        grid1.TextMatrix(grid1.Row, 8) = grid1.TextMatrix(grid1.Row, 2)
        
        If grid1.TextMatrix(grid1.Row, 3) = "0" Then
            grid1.TextMatrix(grid1.Row, 9) = "Title"
        ElseIf grid1.TextMatrix(grid1.Row, 3) = "1" Then
            grid1.TextMatrix(grid1.Row, 9) = "Space After 1"
        ElseIf grid1.TextMatrix(grid1.Row, 3) = "2" Then
            grid1.TextMatrix(grid1.Row, 9) = "Space After 2"
        ElseIf grid1.TextMatrix(grid1.Row, 3) = "3" Then
            grid1.TextMatrix(grid1.Row, 9) = "Space After 3"
        ElseIf grid1.TextMatrix(grid1.Row, 3) = "4" Then
            grid1.TextMatrix(grid1.Row, 9) = "Space After 4"
        ElseIf grid1.TextMatrix(grid1.Row, 3) = "5" Then
            grid1.TextMatrix(grid1.Row, 9) = "Eject"
        End If
        
        If grid1.TextMatrix(grid1.Row, 4) = "0" Then
            grid1.TextMatrix(grid1.Row, 10) = "Header"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "1" Then
            grid1.TextMatrix(grid1.Row, 10) = "Detail"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "2" Then
            grid1.TextMatrix(grid1.Row, 10) = "Sub Total"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "3" Then
            grid1.TextMatrix(grid1.Row, 10) = "Total"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "4" Then
            grid1.TextMatrix(grid1.Row, 10) = "Grand Total 1"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "5" Then
            grid1.TextMatrix(grid1.Row, 10) = "Grand Total 2"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "6" Then
            grid1.TextMatrix(grid1.Row, 10) = "Grand Total 3"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "7" Then
            grid1.TextMatrix(grid1.Row, 10) = "Grand Total 4"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "8" Then
            grid1.TextMatrix(grid1.Row, 10) = "Grand Total 5"
        ElseIf grid1.TextMatrix(grid1.Row, 4) = "9" Then
            grid1.TextMatrix(grid1.Row, 10) = ""
        End If
        
        If grid1.TextMatrix(grid1.Row, 5) = "0" Then
            grid1.TextMatrix(grid1.Row, 11) = "Normal"
        ElseIf grid1.TextMatrix(grid1.Row, 5) = "1" Then
            grid1.TextMatrix(grid1.Row, 11) = "Tebal"
        ElseIf grid1.TextMatrix(grid1.Row, 5) = "2" Then
            grid1.TextMatrix(grid1.Row, 11) = "Tebal & Garis Bawah"
        ElseIf grid1.TextMatrix(grid1.Row, 5) = "3" Then
            grid1.TextMatrix(grid1.Row, 11) = "Blok Text & Angka"
        ElseIf grid1.TextMatrix(grid1.Row, 5) = "4" Then
            grid1.TextMatrix(grid1.Row, 11) = "Blok Text, Angka Tebal"
        ElseIf grid1.TextMatrix(grid1.Row, 5) = "5" Then
            grid1.TextMatrix(grid1.Row, 11) = "Text Tebal, Blok Angka"
        End If
        
        grid1.TextMatrix(grid1.Row, 12) = grid1.TextMatrix(grid1.Row, 6)
        grid1.TextMatrix(grid1.Row, 13) = RST!print_cash
        
        OBJ1.Open dsn
        x = 1
        SQL1 = "select * from gl_dforms where form_no = '" & txtreportcode & "' and group_no = '" & RST!group_no & "' order by line_no asc"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            myarray(grid1.Row, 0, x) = RST1!line_no
            myarray(grid1.Row, 1, x) = RST1!acc_no1
            myarray(grid1.Row, 2, x) = RST1!acc_no2
            myarray(grid1.Row, 3, x) = RST1!acc_no3
            myarray(grid1.Row, 4, x) = RST1!acc_no4
            myarray(grid1.Row, 5, x) = RST1!acc_no5
            myarray(grid1.Row, 6, x) = RST1!acc_no6
            myarray(grid1.Row, 7, x) = RST1!acc_no7
            myarray(grid1.Row, 8, x) = RST1!acc_no8
            myarray(grid1.Row, 9, x) = RST1!acc_no9
            myarray(grid1.Row, 10, x) = RST1!group_no
        
            x = x + 1
            RST1.MoveNext
        Loop
        OBJ1.Close
        
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdclear_Click()
    txtreportcode = ""
    hapusform
    hapusgroup
    hapusdetail
    txtreportcode.SetFocus
    SSTab1.Tab = 0
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grid1.TextMatrix(0, 0) = "Group No"
    grid1.TextMatrix(0, 7) = "Type Account"
    grid1.TextMatrix(0, 8) = "Description"
    grid1.TextMatrix(0, 9) = "Space After"
    grid1.TextMatrix(0, 10) = "Print Coloum"
    grid1.TextMatrix(0, 11) = "Print Mode"
    grid1.TextMatrix(0, 12) = "Sign"
    grid1.TextMatrix(0, 13) = "Print Cash"
    grid1.ColWidth(0) = 800
    grid1.ColWidth(1) = 0
    grid1.ColWidth(2) = 0
    grid1.ColWidth(3) = 0
    grid1.ColWidth(4) = 0
    grid1.ColWidth(5) = 0
    grid1.ColWidth(6) = 0
    grid1.ColWidth(7) = 1100
    grid1.ColWidth(8) = 1500
    grid1.ColWidth(9) = 1100
    grid1.ColWidth(10) = 1000
    grid1.ColWidth(11) = 1000
    grid1.ColWidth(12) = 500
    grid1.ColWidth(13) = 1000
    
    grid2.TextMatrix(0, 0) = "Line No."
    grid2.TextMatrix(0, 1) = "Acc. 1"
    grid2.TextMatrix(0, 2) = "Acc. 2"
    grid2.TextMatrix(0, 3) = "Acc. 3"
    grid2.TextMatrix(0, 4) = "Acc. 4"
    grid2.TextMatrix(0, 5) = "Acc. 5"
    grid2.TextMatrix(0, 6) = "Acc. 6"
    grid2.TextMatrix(0, 7) = "Acc. 7"
    grid2.TextMatrix(0, 8) = "Acc. 8"
    grid2.TextMatrix(0, 9) = "Acc. 9"
    grid2.ColWidth(0) = 800
    grid2.ColWidth(1) = 800
    grid2.ColWidth(2) = 800
    grid2.ColWidth(3) = 800
    grid2.ColWidth(4) = 800
    grid2.ColWidth(5) = 800
    grid2.ColWidth(6) = 800
    grid2.ColWidth(7) = 800
    grid2.ColWidth(8) = 800
    grid2.ColWidth(9) = 800
    
    grid1.RowHeightMin = 300
    grid2.RowHeightMin = 300
    lblapor = 1
End Sub

Private Sub caridetail()
    lblgroupno = "Group No. : " & grid1.TextMatrix(setup3, 0) & " |Type Account : " & grid1.TextMatrix(setup3, 7) & " |Description : " & grid1.TextMatrix(setup3, 8)
    For x = 0 To 100
        If myarray(x, 10, 1) = grid1.TextMatrix(setup3, 0) Then
            grid2.Row = 1
            Do While True
                If myarray(x, 10, grid2.Row) = "" Then Exit Do
                
                grid2.TextMatrix(grid2.Row, 0) = myarray(x, 0, grid2.Row)
                If myarray(x, 0, grid2.Row) >= -9 And myarray(x, 0, grid2.Row) <= -1 Then
                    grid2.TextMatrix(grid2.Row, 1) = myarray(x, 1, grid2.Row)
                Else
                    grid2.TextMatrix(grid2.Row, 1) = original(myarray(x, 1, grid2.Row))
                End If
                
                grid2.TextMatrix(grid2.Row, 2) = original(myarray(x, 2, grid2.Row))
                grid2.TextMatrix(grid2.Row, 3) = original(myarray(x, 3, grid2.Row))
                grid2.TextMatrix(grid2.Row, 4) = original(myarray(x, 4, grid2.Row))
                grid2.TextMatrix(grid2.Row, 5) = original(myarray(x, 5, grid2.Row))
                grid2.TextMatrix(grid2.Row, 6) = original(myarray(x, 6, grid2.Row))
                grid2.TextMatrix(grid2.Row, 7) = original(myarray(x, 7, grid2.Row))
                grid2.TextMatrix(grid2.Row, 8) = original(myarray(x, 8, grid2.Row))
                grid2.TextMatrix(grid2.Row, 9) = original(myarray(x, 9, grid2.Row))
        
                grid2.Rows = grid2.Rows + 1
                grid2.Row = grid2.Row + 1
            Loop
            
            Exit Sub
        End If
    Next x
End Sub

Private Sub hapusdetail()
    grid2.ColWidth(0) = 800
    grid2.ColWidth(1) = 800
    grid2.ColWidth(2) = 800
    grid2.ColWidth(3) = 800
    grid2.ColWidth(4) = 800
    grid2.ColWidth(5) = 800
    grid2.ColWidth(6) = 800
    grid2.ColWidth(7) = 800
    grid2.ColWidth(8) = 800
    grid2.ColWidth(9) = 800
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.TextMatrix(grid2.Row, 8) = ""
        grid2.TextMatrix(grid2.Row, 9) = ""
        
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.SetFocus
    lbltotdetail = "Total Detail : " & grid2.Rows - 2
    lblgroupno = ""
End Sub

Private Sub cmdaddetail_Click()
    If txtreportcode = "" Then Exit Sub
    setup2 = lblapor
    setup1 = ""
    frmreportdetail.Show 1
    grid2.SetFocus
    grid2.Row = 1
    grid2.Col = 0
End Sub

Private Sub cmdaddgroup_Click()
    If txtreportcode = "" Then Exit Sub
    setup2 = lblapor
    setup3 = ""
    frmreportgroup.Show 1
    grid1.SetFocus
    grid1.Row = 1
    grid1.Col = 0
End Sub
