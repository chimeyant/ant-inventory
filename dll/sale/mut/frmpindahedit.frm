VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpindahedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Pindah Gudang"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3150
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   120
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
      Left            =   3630
      Picture         =   "frmpindahedit.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   120
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
      Left            =   3390
      Picture         =   "frmpindahedit.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtapply 
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
      Left            =   1470
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox txtnobukti 
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
      Left            =   1455
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtgudang 
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
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtgudang2 
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
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   6510
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin Chameleon.chameleonButton cmdnolot 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "No Lot"
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
         MICON           =   "frmpindahedit.frx":0630
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblkdproduk 
         BackColor       =   &H80000014&
         Height          =   255
         Left            =   1335
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblnolot 
         BackColor       =   &H80000014&
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
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Left            =   1815
      TabIndex        =   5
      Text            =   "PG0-160923"
      Top             =   4605
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   5670
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "frmpindahedit.frx":094A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":09B6
      Key             =   "frmpindahedit.frx":09D4
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
      Height          =   255
      Left            =   7470
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpindahedit.frx":0A10
      Caption         =   "frmpindahedit.frx":0A30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":0A9C
      Keys            =   "frmpindahedit.frx":0ABA
      Spin            =   "frmpindahedit.frx":0AFC
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
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   480
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
      Format          =   192282627
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2415
      Left            =   0
      TabIndex        =   15
      Top             =   2010
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8310
      TabIndex        =   16
      Top             =   4560
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
      MICON           =   "frmpindahedit.frx":0B24
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
      Left            =   7350
      TabIndex        =   17
      Top             =   4545
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
      MICON           =   "frmpindahedit.frx":0E3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdUpdate 
      Height          =   375
      Left            =   5430
      TabIndex        =   18
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "frmpindahedit.frx":1158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   270
      TabIndex        =   19
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "dari Gudang"
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
      MICON           =   "frmpindahedit.frx":1472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   225
      Left            =   8430
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpindahedit.frx":178C
      Caption         =   "frmpindahedit.frx":17AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":1818
      Keys            =   "frmpindahedit.frx":1836
      Spin            =   "frmpindahedit.frx":1878
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   225
      Left            =   8430
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpindahedit.frx":18A0
      Caption         =   "frmpindahedit.frx":18C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":192C
      Keys            =   "frmpindahedit.frx":194A
      Spin            =   "frmpindahedit.frx":198C
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   270
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "ke Gudang"
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
      MICON           =   "frmpindahedit.frx":19B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil3 
      Height          =   225
      Left            =   8430
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpindahedit.frx":1CCE
      Caption         =   "frmpindahedit.frx":1CEE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":1D5A
      Keys            =   "frmpindahedit.frx":1D78
      Spin            =   "frmpindahedit.frx":1DBA
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   3990
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   192282627
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3990
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   192282627
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdel 
      Height          =   375
      Left            =   6390
      TabIndex        =   26
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpindahedit.frx":1DE2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   270
      TabIndex        =   27
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Bukti"
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
      MICON           =   "frmpindahedit.frx":20FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil4 
      Height          =   225
      Left            =   8430
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpindahedit.frx":2416
      Caption         =   "frmpindahedit.frx":2436
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpindahedit.frx":24A2
      Keys            =   "frmpindahedit.frx":24C0
      Spin            =   "frmpindahedit.frx":2502
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
      MinValue        =   0
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
   Begin Chameleon.chameleonButton cmdprint 
      Height          =   375
      Left            =   105
      TabIndex        =   29
      Top             =   4575
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Test"
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
      MICON           =   "frmpindahedit.frx":252A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   1050
      Top             =   4530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
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
      Left            =   6870
      TabIndex        =   34
      Top             =   1710
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
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
      Left            =   270
      TabIndex        =   33
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Desc/Reference"
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
      Left            =   270
      TabIndex        =   32
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
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
      Left            =   3150
      TabIndex        =   31
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblcust 
      BackColor       =   &H80000014&
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
      Left            =   3150
      TabIndex        =   30
      Top             =   1335
      Width           =   2295
   End
End
Attribute VB_Name = "frmpindahedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim posrow, poscol As String
Dim boo1, hitunginout As Boolean

Private Sub cmdclear_Click()
    txtnobukti.Enabled = True
    txtgudang.Enabled = True
    txtgudang2.Enabled = True
    cmdsearch1.Enabled = True
    cmdsearch2.Enabled = True
    cmdsearch3.Enabled = True
    date1.Enabled = True
    Frame1.Visible = False
    lblnolot = ""
    lblkdproduk = ""
    hapusemua
    txtnobukti = ""
    boo1 = False
    txtnobukti.SetFocus
End Sub

Private Sub hapusemua()
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1 = Date
    End If
    
    txtgudang2 = ""
    lblcust = ""
    txtgudang = ""
    lblgudang = ""
    txtapply = ""
    
    hapusgrid
    
    lbltotal.Caption = "    Total Barang : 0"
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.ColWidth(6) = 1000
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry in the grid is incomplete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If par1 = "1" Then hitunginout = True Else hitunginout = False
    
    If boo1 Then
        OBJ.Open dsn
        SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            OBJ.Close
            MsgBox "Can not delete, out of date range.", vbExclamation, "Warning"
            Exit Sub
        End If
        OBJ.Close
        
        grid.Row = 1
        Do While True
            If grid.Rows = grid.Row + 1 Then Exit Do
            If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Then
                MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
        
        If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            cmdclear_Click
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "select nobpb from am_bpbhdr where nobpb = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Delete aborted, data not found.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
        OBJ.Close
        
        If hitunginout Then
            OBJ2.Open dsn
            SQL2 = "select * from am_bpblin where nobpb = '" & txtnobukti & "' and type = '88'"
            Set RST2 = OBJ2.Execute(SQL2)
            Do While Not RST2.EOF
                OBJ.Open dsn
                SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
                
                If par5 = "0" Then
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Else
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
                
                SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj < '" & tanggalinv & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
                OBJ.Close
                
                txtnil3 = txtnil1 - txtnil2 - RST2!qty - txtnil4
                date2 = date1
                date3 = date1
                
                OBJ.Open dsn
                SQL = "select isnull(max(a.tglbpb),01/01/1900)'tanggal' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                
                If par5 = "0" Then
                    SQL = "select isnull(max(a.tglsj),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Else
                    SQL = "select isnull(max(b.tglkirim),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                
                SQL = "select isnull(max(tglsj),01/01/1900)'tanggal' from am_sjsby where kodegudang = '" & txtgudang & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                OBJ.Close
                
                Do While True
                    OBJ.Open dsn
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb = '" & tanggal2 & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
                    
                    If par5 = "0" Then
                        SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    Else
                        SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    End If
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
                    
                    SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj = '" & tanggal2 & "' and kodegudang = '" & txtgudang & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
                    OBJ.Close
                    
                    txtnil3 = txtnil3 + txtnil1 - txtnil2 - txtnil4
                    
                    If txtnil3 < 0 Then
                        MsgBox "Stock Limited on " & RST2!kodebarang, vbOKOnly + vbExclamation, "Warning"
                        Exit Sub
                    End If
                                
                    If date2 = date3 Then Exit Do
                    
                    date2 = date2 + 1
                Loop
                
                OBJ.Open dsn
                SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb < '" & tanggalinv & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
                
                If par5 = "0" Then
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Else
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim < '" & tanggalinv & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
                
                SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj < '" & tanggalinv & "' and kodegudang = '" & txtgudang2 & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
                OBJ.Close
                
                txtnil3 = txtnil1 - txtnil2 - RST2!qty - txtnil4
                date2 = date1
                date3 = date1
                
                OBJ.Open dsn
                SQL = "select isnull(max(a.tglbpb),01/01/1900)'tanggal' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                
                If par5 = "0" Then
                    SQL = "select isnull(max(a.tglsj),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                Else
                    SQL = "select isnull(max(b.tglkirim),01/01/1900)'tanggal' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                End If
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                
                SQL = "select isnull(max(tglsj),01/01/1900)'tanggal' from am_sjsby where kodegudang = '" & txtgudang2 & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then If date3 < RST!tanggal Then date3 = RST!tanggal
                OBJ.Close
                
                Do While True
                    OBJ.Open dsn
                    SQL = "select isnull(sum(a.qty),0)'qty' from am_bpblin a left join am_bpbhdr b on a.type=b.type and a.nobpb=b.nobpb and a.tglbpb=b.tglbpb where a.tglbpb = '" & tanggal2 & "' and a.nobpb <> '" & txtnobukti & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil1 = RST!qty Else txtnil1 = 0
                    
                    If par5 = "0" Then
                        SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and a.tglsj = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    Else
                        SQL = "select isnull(sum(a.qty),0)'qty' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.tglkirim = '" & tanggal2 & "' and b.kodegudang = '" & txtgudang2 & "' and a.kodebarang = '" & RST2!kodebarang & "' and a.kodesatuan = '" & RST2!kodesatuan & "'"
                    End If
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil2 = RST!qty Else txtnil2 = 0
                    
                    SQL = "select isnull(sum(qty),0)'qty' from am_sjsby where tglsj = '" & tanggal2 & "' and kodegudang = '" & txtgudang2 & "' and kodebarang = '" & RST2!kodebarang & "' and kodesatuan = '" & RST2!kodesatuan & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then txtnil4 = RST!qty Else txtnil4 = 0
                    OBJ.Close
                    
                    txtnil3 = txtnil3 + txtnil1 - txtnil2 - txtnil4
                    
                    If txtnil3 < 0 Then
                        MsgBox "Stock Limited on " & RST2!kodebarang, vbOKOnly + vbExclamation, "Warning"
                        Exit Sub
                    End If
                                
                    If date2 = date3 Then Exit Do
                    
                    date2 = date2 + 1
                Loop
                
                RST2.MoveNext
            Loop
            OBJ2.Close
        End If
        '------------
        
        OBJ.Open dsn
        SQL = "delete from am_bpbhdr where nobpb = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_bpblin where nobpb = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
    Else
        MsgBox "Delete aborted, data not found.", vbExclamation, "Warning"
    End If
End Sub

Private Sub cmdnolot_Click()
    namatabel = "nolot"
    carisql1 = "Select tanggal,kode_produk,nolot from list_produksi_master "
    'carisql1 = carisql1 + " where flagprint =4"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnolot_GotFocus()
    If hasil = "" Then Exit Sub
    lblkdproduk = hasil
    lblnolot = hasil1
    txtapply = hasil1
    hasil = ""
    hasil1 = ""
    carisql1 = ""
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang2 = hasil
    lblcust = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtapply.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    lblgudang = hasil1
    hasil = ""
    hasil1 = ""
    If txtgudang = "G3" Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
    txtgudang2.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PG0-%' and type = '99' and tglbpb >= '" & batas1 & "' and tglbpb <= '" & batas2 & "'"
    Else
        carisql1 = "select nobpb, convert(char(11),tglbpb )'tglbpb' from am_bpbhdr where nobpb like 'PG0-%' and type = '99'"
    End If
    namatabel = "Pindah Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    boo1 = True
    hasil = ""
    hasil1 = ""
    Call loaddata
   ' txtapply.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, AppName
        Exit Sub
    End If
    If grid.Rows = 2 Then
        MsgBox "Data Entry on Grid Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Update am_bpbhdr set "
    SQL = SQL + "tglbpb=convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "kodegudang='" & txtgudang & "',"
    SQL = SQL + "keterangan='" & txtapply & "',"
    SQL = SQL + "noref='" & lblnolot & "',"
    SQL = SQL + "idupdate='" & UserOnline & "',"
    SQL = SQL + "dateupdate=convert(datetime,'" & tanggalsekarang & "') "
    SQL = SQL + "Where nobpb='" & txtnobukti & "' and type='99'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "Update am_bpbhdr set "
    SQL = SQL + "tglbpb=convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "kodegudang='" & txtgudang2 & "',"
    SQL = SQL + "keterangan='" & txtapply & "',"
    SQL = SQL + "noref='" & lblnolot & "',"
    SQL = SQL + "idupdate='" & UserOnline & "',"
    SQL = SQL + "dateupdate=convert(datetime,'" & tanggalsekarang & "') "
    SQL = SQL + "Where nobpb='" & txtnobukti & "' and type='88'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE from am_bpblin WHERE nobpb='" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
'TIDAK BISA MENGGUNAKAN METODE UPDATE, KARENA ADA KONDISI UPDATE
'DENGAN MENGHAPUS 1 BARIS BARANG PADA GRID (MISAL 3 BARANG{ROW} MENJADI 1 ATAU 2 BARANG{ROW})
'ATAU DENGAN MENAMBAH BARANG(ROW) JADI HARUS DELETE BARU INSERT
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "insert into am_bpblin ("
        SQL = SQL + "type,"
        SQL = SQL + "nobpb,"
        SQL = SQL + "tglbpb,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values("
        SQL = SQL + "'99',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") * -1 & "'),"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "insert into am_bpblin ("
        SQL = SQL + "type,"
        SQL = SQL + "nobpb,"
        SQL = SQL + "tglbpb,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values("
        SQL = SQL + "'88',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Close
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "K/Satuan"
    grid.TextMatrix(0, 5) = "N/Satuan"
    grid.TextMatrix(0, 6) = "Qty WIP"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.ColWidth(6) = 1000
    
    grid.RowHeightMin = 300
    boo1 = False
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
    
    If grid.Rows = 2 Then
        lbltotal.Caption = "    Total Barang : 0"
    Else
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    End If
End Sub


Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then Exit Sub
    If txtgudang = "G3" Then
        If lblnolot = "" Then Exit Sub
    End If
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 200 Then
                MsgBox "Maximum line 200", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtnilai.Visible = True Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then Exit Sub
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        If txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select

End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    If grid.Col = 4 Then
        grid.Row = 1
        Do While True
            If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 4) = hasil And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
            
                MsgBox "Kode Barang already exist.", vbInformation, "Information"
                hasil = ""
                grid.Row = posrow
                grid.Col = 4
                grid.SetFocus
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
    End If

    grid.Row = posrow
    grid.Col = poscol
    grid.CellAlignment = 1
    grid.TextMatrix(grid.Row, grid.Col) = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""

    If grid.Col = 1 Then
        OBJ.Open dsn
        SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = "0.00"
            
            SetRow grid.Row, True
            lbltotal.Caption = "    Total Barang : " & grid.Rows - 1
            grid.SetFocus
            grid.Col = 2
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
        End If
        OBJ.Close
    End If

    If grid.Col = 4 Then
        OBJ.Open dsn
        SQL = "SELECT a.namabarang,b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            grid.SetFocus
            grid.Col = 5
            hitung_qty_wip
        Else
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        If OBJ.State = 1 Then OBJ.Close
    End If

End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub


Private Sub txtapply_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    
    
    If KeyAscii = 13 Then
        If txtgudang = "G3" Then
            If lblnolot = "" Then
                MsgBox "Silahkan pilih nomor lot terlebih dahulu", vbCritical, "Peringatan"
                Exit Sub
            End If
            
        End If
        Select Case grid.Col
            Case 1
                grid.Row = posrow
                
                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and len(kodebarang)=8"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    grid.TextMatrix(grid.Row, 2) = RST!namabarang
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    
                    OBJ.Close
    
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    
                    If txtgudang = "G3" Then
                        carisql1 = "Select distinct a.kodebarang,b.namabarang from am_bpblin a inner join am_itemmst b "
                        carisql1 = carisql1 + "on a.kodebarang = b.KodeBarang inner join am_bpbhdr c on a.nobpb =c.nobpb "
                        carisql1 = carisql1 + "Where c.kodegudang = 'G3' and c.noref like '" + lblnolot + "%'"
                        namatabel = "Item Gudang WIF"
                    Else
                
                    carisql1 = "select kodebarang, namabarang from am_itemmst"
                    namatabel = "Item"
                    End If
                    frmsearch.Show vbModal
                End If
                grid.Col = 1
            Case 4
                grid.Row = 1
                Do While True
                    If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                    If grid.TextMatrix(grid.Row, 4) = txtket And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
                    
                        MsgBox "Kode Barang already exist.", vbInformation, "Information"
                        txtket = ""
                        grid.Row = posrow
                        grid.Col = 4
                        grid.SetFocus
                        Exit Sub
                    End If
                    grid.Row = grid.Row + 1
                Loop
                
                grid.Row = posrow
                
                grid.SetFocus
                grid.Col = 4
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "SELECT namabarang,kodesatuan FROM AM_ITEMDTL WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    grid.TextMatrix(posrow, 4) = ""
                    
                    txtket = ""
                    
                    carisql1 = "SELECT b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b on a.kodesatuan = b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                    namatabel = "Satuan "
                        
                    frmsearch.Show vbModal
                Else
                    grid.TextMatrix(grid.Row, 2) = RST!namabarang
                    
                    SQL = "SELECT namasatuan FROM AM_unit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then grid.TextMatrix(grid.Row, 2) = RST!namasatuan
                    
                End If
                OBJ.Close
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

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then
        txtnilai_LostFocus
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Select SUM(a.qty)as qty From am_bpblin a inner join am_bpbhdr b "
        SQL = SQL + "on a.nobpb = b.nobpb and (a.type = '01' or a.type = '99') "
        SQL = SQL + "Where a.kodebarang = '" + grid.TextMatrix(grid.Row, 1) + "' and b.kodegudang = 'G3' "
        SQL = SQL + "and b.noref like '" + lblnolot + "%'"
        Set RST = OBJ.Execute(SQL)
        
        If RST!qty < txtnilai Then
            MsgBox "Qty tidak mencukupi" & Chr(13) _
            & "" & Chr(13) _
            & "KODE BARANG : " & grid.TextMatrix(grid.Row, 1) & Chr(13) _
            & "TOTAL Qty        : " & RST!qty, vbCritical, "WARNING"
            grid.TextMatrix(grid.Row, grid.Col) = "0.00"
        Else
            grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        End If
        OBJ.Close
bawah:
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
        grid.Col = poscol
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call loaddata
    End If
End Sub

Private Sub loaddata()
        hapusgrid
        OBJ.Open dsn
        SQL = "Select a.*,b.namagudang From am_bpbhdr a inner join am_gudang b"
        SQL = SQL + " On a.kodegudang=b.kodegudang Where a.nobpb='" & txtnobukti.text & "' and a.type='99'"
        Set RST = OBJ.Execute(SQL)
        
        If RST.EOF Then
            OBJ.Close
            MsgBox "Data tidak ditemukan.", vbCritical, AppName
            Exit Sub
        End If
        
        date1 = RST!tglbpb
        txtgudang = RST!kodegudang
        lblgudang = RST!namagudang
        txtapply = RST!keterangan
        lblnolot = RST!noref
        If Left(txtnobukti.text, 3) = "PG0" And RST!kodegudang = "G3" Then Frame1.Visible = True

        SQL = "Select a.kodegudang,b.namagudang From am_bpbhdr a inner join am_gudang b"
        SQL = SQL + " On a.kodegudang=b.kodegudang Where nobpb='" & txtnobukti.text & "' and type='88'"
        Set RST = OBJ.Execute(SQL)
        txtgudang2 = RST!kodegudang
        lblcust = RST!namagudang
        OBJ.Close
        
        OBJ1.Open dsn
        SQL1 = "Select a.kodebarang,b.NamaBarang,a.qty,a.keterangan,a.kodesatuan,c.NamaSatuan"
        SQL1 = SQL1 + " From am_bpblin a inner join am_itemmst b on a.kodebarang = b.KodeBarang"
        SQL1 = SQL1 + " inner join am_unit c on a.kodesatuan = c.KodeSatuan"
        SQL1 = SQL1 + " Where a.nobpb='" & txtnobukti.text & "' and a.type='88'"
        'SQL1 = SQL1 + " Order By a.lineitem Asc"
        Set RST1 = OBJ1.Execute(SQL1)

        'Do Until RST1.EOF
        Do While Not RST1.EOF
            grid.Col = 0
            Set grid.CellPicture = uncheck
            grid.TextMatrix(grid.Row, 1) = RST1!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST1!namabarang
            grid.TextMatrix(grid.Row, 3) = Format(RST1!qty, "##0.00")
            grid.TextMatrix(grid.Row, 4) = RST1!kodesatuan
            grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            hitung_qty_wip
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            lbltotal.Caption = "    Total Barang : " & grid.Row
            DoEvents
            RST1.MoveNext
        Loop
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
        OBJ1.Close
End Sub

Private Function hitung_qty_wip(Optional ByVal baris As Integer)
     If OBJ.State = 0 Then OBJ.Open dsn
        SQL = "Select SUM(a.qty)as qty From am_bpblin a inner join am_bpbhdr b "
        SQL = SQL + "on a.nobpb = b.nobpb and (a.type = '01' or a.type = '99') "
        SQL = SQL + "Where a.kodebarang = '" + grid.TextMatrix(grid.Row, 1) + "' and b.kodegudang = 'G3' "
        SQL = SQL + "and b.noref like '" + lblnolot + "%'"
        Set RST = OBJ.Execute(SQL)
        grid.TextMatrix(grid.Row, 6) = Format(RST!qty, "###,###,##0.00")
        OBJ.Close
        Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function
Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
Function tanggal3()
    tanggal3 = Month(date3) & "/" & Day(date3) & "/" & Year(date3)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function
Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub
