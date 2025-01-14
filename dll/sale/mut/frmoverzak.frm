VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmoverzak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Over Zak"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnobukti2 
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
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   34
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtsatuan 
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtkode 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   31
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtbrg 
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
      Left            =   2760
      TabIndex        =   30
      Top             =   960
      Width           =   3495
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
      Left            =   840
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   5520
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
      Left            =   1320
      Picture         =   "frmoverzak.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   5520
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
      Left            =   1080
      Picture         =   "frmoverzak.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   5520
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
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2880
      Width           =   7335
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
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   19
      Top             =   600
      Width           =   855
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
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtnolot 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   6360
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
      MICON           =   "frmoverzak.frx":0630
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
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   6840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2566
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
   Begin MSComCtl2.DTPicker Date1 
      Height          =   285
      Left            =   6960
      TabIndex        =   12
      Top             =   240
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
      CustomFormat    =   "ddMMMMyyyy"
      Format          =   134807555
      CurrentDate     =   42052
   End
   Begin Chameleon.chameleonButton cmdgudang 
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Gudang"
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
      MICON           =   "frmoverzak.frx":094A
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
      Left            =   6960
      TabIndex        =   9
      Top             =   6360
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
      MICON           =   "frmoverzak.frx":0C64
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdSave 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   6360
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
      MICON           =   "frmoverzak.frx":0F7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmoverzak.frx":1298
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoverzak.frx":1304
      Key             =   "frmoverzak.frx":1322
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
      Left            =   4440
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmoverzak.frx":135E
      Caption         =   "frmoverzak.frx":137E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoverzak.frx":13EA
      Keys            =   "frmoverzak.frx":1408
      Spin            =   "frmoverzak.frx":144A
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin Chameleon.chameleonButton cmdasal 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Kode Asal"
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
      MICON           =   "frmoverzak.frx":1472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtqty 
      Height          =   285
      Left            =   1440
      TabIndex        =   35
      Top             =   1320
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Calculator      =   "frmoverzak.frx":178C
      Caption         =   "frmoverzak.frx":17AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmoverzak.frx":1818
      Keys            =   "frmoverzak.frx":1836
      Spin            =   "frmoverzak.frx":1878
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   8454143
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblKsatuan 
      Height          =   255
      Left            =   5400
      TabIndex        =   36
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Transaksi"
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
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      TabIndex        =   32
      Top             =   1320
      Width           =   855
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
      Left            =   6480
      TabIndex        =   25
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lbltype2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Terima Over Zak"
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
      TabIndex        =   29
      Top             =   2520
      Width           =   1335
   End
   Begin MSForms.ComboBox cmbtype2 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   735
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1296;503"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblnotif 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Width           =   5775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   21
      Top             =   2910
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
      Left            =   6960
      TabIndex        =   20
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label lbltype 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Keluar Over Zak"
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
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin MSForms.ComboBox cmbtype 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   735
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1296;503"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl. Mutasi"
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
      Left            =   6000
      TabIndex        =   15
      Top             =   290
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Lot"
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
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Transaksi"
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
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "66666"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   8895
   End
End
Attribute VB_Name = "frmoverzak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim posrow, poscol As String
Dim str99 As String

Private Sub cmbtype_Change()
Dim strformat As String
    strformat = Format(Date, "yymm")
    'txtnobukti = ""
    'txtnobukti.SetFocus
    
    If cmbtype2 = "02" Then
        lbltype2 = "Terima Over Zak"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TOZ0-' + '" + strformat + "%' order by dateentry desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If
        OBJ.Close

        str99 = str99 + 1

        If Len(str99) = 1 Then txtnobukti2 = "TOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti2 = "TOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti2 = "TOZ0-" & strformat & str99
    End If

    If cmbtype = "03" Then
        lbltype = "Keluar Over Zak"
        
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KOZ0-' + '" + strformat + "%' order by dateentry desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If
        OBJ.Close

        str99 = str99 + 1

        If Len(str99) = 1 Then txtnobukti = "KOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti = "KOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti = "KOZ0-" & strformat & str99
    End If
End Sub

Private Sub cmdasal_Click()
    hasil8 = txtnolot
    namatabel = "Item Lot"
    carisql1 = "Select kodebarang,namabarang,SUM(qin)-SUM(qout)'qty',satuan From am_stokgudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdasal_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    txtbrg = hasil1
    hasil = "": hasil1 = ""
    showdata
End Sub

Private Sub cmdclear_Click()
    hapusgrid1
    hapusgrid
    txtnolot = ""
    cmbtype = "03"
    cmbtype2 = "02"
    txtgudang = ""
    lblgudang = ""
    date1 = Date
    txtkode = ""
    txtbrg = ""
    txtqty = "0.00"
    txtsatuan = ""
    cmbtype_Change
    txtapply = ""
    lblnotif = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdgudang_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdgudang_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    lblgudang = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsave_Click()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    If txtnolot = "" Then Exit Sub
    If txtnobukti = "" Then
        MsgBox "Kolom Nomor bukti kosong" & vbCrLf & "silahkan pilih type transaksi terlebih dahulu", vbCritical, AppName
        Exit Sub
    End If
    
    If txtgudang = "" Then
        MsgBox "Kolom Gudang kosong", vbCritical, AppName
        Exit Sub
    End If
    
    If txtkode = "" Then
        MsgBox "kolom kode asal kosong", vbCritical, AppName
        Exit Sub
    End If
    
    
    lblnotif = "checking auto numbering format ..."
    OBJ.Open dsn
    If cmbtype2 = "02" Then
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'TOZ0-' + '" + strformat + "%' order by dateentry desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If

        str99 = str99 + 1
            
        If Len(str99) = 1 Then txtnobukti2 = "TOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti2 = "TOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti2 = "TOZ0-" & strformat & str99
    End If
    If cmbtype = "03" Then
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'KOZ0-' + '" + strformat + "%' order by dateentry desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 3)
        Else
            str99 = 0
        End If
            
        str99 = str99 + 1
        
        If Len(str99) = 1 Then txtnobukti = "KOZ0-" & strformat & "00" & str99
        If Len(str99) = 2 Then txtnobukti = "KOZ0-" & strformat & "0" & str99
        If Len(str99) = 3 Then txtnobukti = "KOZ0-" & strformat & str99
    End If
    OBJ.Close
    
    lblnotif = "Inserting data keluar over zak..."
    OBJ.Open dsn
    'keluar over zak
    SQL = "insert into am_bpbhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nobpb,"
    SQL = SQL + "tglbpb,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "noref,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'" & cmbtype & "',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'" & txtgudang & "',"
    SQL = SQL + "'" & txtnolot & " : " & txtapply & "',"
    SQL = SQL + "'" & txtnolot & "',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
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
    SQL = SQL + "'" & cmbtype & "',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'" & txtkode & "',"
    SQL = SQL + "convert(money,'" & txtqty & "' * -1),"
    SQL = SQL + "'" & txtapply & "',"
    SQL = SQL + "'1',"
    SQL = SQL + "'" & lblKsatuan & "')"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "insert into am_stokgudang ("
    SQL = SQL + "nolot,"
    SQL = SQL + "palet,"
    SQL = SQL + "tanggal,"
    SQL = SQL + "ref,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "kodebarang,"
    SQL = SQL + "namabarang,"
    SQL = SQL + "kg,"
    SQL = SQL + "kgperpalet,"
    SQL = SQL + "hppperkg,"
    SQL = SQL + "qin,"
    SQL = SQL + "qout,"
    SQL = SQL + "kdsatuan,"
    SQL = SQL + "satuan,"
    SQL = SQL + "gudang,"
    SQL = SQL + "username,"
    SQL = SQL + "flag)"
        
    SQL = SQL + " values("
    SQL = SQL + "'" & txtnolot & "',"
    SQL = SQL + "'01' + '" & txtnolot & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "'" & lbltype & "',"
    SQL = SQL + "'" & txtkode & "',"
    SQL = SQL + "'" & txtbrg & "',"
    SQL = SQL + "'" & grid1.TextMatrix(1, 2) & "',"
    SQL = SQL + "'" & grid1.TextMatrix(1, 2) * txtqty & "',"
    SQL = SQL + "'" & grid1.TextMatrix(1, 3) & "',"
    SQL = SQL + "'0.00',"
    SQL = SQL + "convert(money,'" & txtqty & "'),"
    SQL = SQL + "'" & lblKsatuan & "',"
    SQL = SQL + "'" & txtsatuan & "',"
    SQL = SQL + "'" & txtgudang & "',"
    SQL = SQL + "'" & nmuser & "',"
    SQL = SQL + "'1')"
    Set RST = OBJ.Execute(SQL)
    
    lblnotif = "Inserting data terima over zak..."
    SQL = "insert into am_bpbhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nobpb,"
    SQL = SQL + "tglbpb,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "noref,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'" & cmbtype2 & "',"
    SQL = SQL + "'" & txtnobukti2 & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'" & txtgudang & "',"
    SQL = SQL + "'" & txtnolot & " : " & txtapply & "',"
    SQL = SQL + "'" & txtnolot & "',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
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
        SQL = SQL + "'" & cmbtype2 & "',"
        SQL = SQL + "'" & txtnobukti2 & "',"
        SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "insert into am_stokgudang ("
        SQL = SQL + "nolot,"
        SQL = SQL + "palet,"
        SQL = SQL + "tanggal,"
        SQL = SQL + "ref,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "namabarang,"
        SQL = SQL + "kg,"
        SQL = SQL + "kgperpalet,"
        SQL = SQL + "hppperkg,"
        SQL = SQL + "qin,"
        SQL = SQL + "qout,"
        SQL = SQL + "kdsatuan,"
        SQL = SQL + "satuan,"
        SQL = SQL + "gudang,"
        SQL = SQL + "username,"
        SQL = SQL + "flag)"
            
        SQL = SQL + " values("
        SQL = SQL + "'" & txtnolot & "',"
        SQL = SQL + "'01' + '" & txtnolot & "',"
        SQL = SQL + "convert(datetime,'" & tanggaloz & "'),"
        SQL = SQL + "'" & txtnobukti2 & "',"
        SQL = SQL + "'" & lbltype2 & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid1.TextMatrix(1, 2) & "',"
        SQL = SQL + "'" & grid1.TextMatrix(1, 2) * txtqty & "',"
        SQL = SQL + "'" & grid1.TextMatrix(1, 3) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'0.00',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'" & txtgudang & "',"
        SQL = SQL + "'" & nmuser & "',"
        SQL = SQL + "'0')"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    MsgBox "Data is save successfully", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    If txtnolot = "" Then txtnolot.SetFocus
End Sub

Private Sub Form_Load()
    'cmbtype.Clear
    'cmbtype.ColumnCount = 2
    'cmbtype.ListWidth = "7 cm"
    'cmbtype.ColumnWidths = "2 cm; 5 cm"

    'cmbtype.AddItem "02"
    'cmbtype.AddItem "03"
    'cmbtype.List(0, 1) = "Terima Over Zak"
    'cmbtype.List(1, 1) = "Keluar Over Zak"
    cmbtype2 = "02"
    cmbtype = "03"
    date1 = Date
    grid1.Cols = 6
    grid1.TextMatrix(0, 0) = "kode"
    grid1.TextMatrix(0, 1) = "Item"
    grid1.TextMatrix(0, 2) = "Kg"
    grid1.TextMatrix(0, 3) = "Hpp/Kg"
    grid1.TextMatrix(0, 4) = "K/Satuan"
    grid1.TextMatrix(0, 5) = "Satuan"
    
    grid1.ColWidth(0) = 800
    grid1.ColWidth(1) = 1800
    grid1.ColWidth(2) = 800
    grid1.ColWidth(3) = 1200
    grid1.ColWidth(4) = 800
    grid1.ColWidth(5) = 1000
    
    grid.Cols = 6
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "K/Satuan"
    grid.TextMatrix(0, 5) = "N/Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.RowHeightMin = 300

End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    If txtnobukti = "" Or txtgudang = "" Then Exit Sub
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
    If txtnobukti = "" Or txtgudang = "" Then Exit Sub
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

Private Sub txtnolot_LostFocus()
    txtnolot_KeyPress 13
End Sub

Private Sub txtnolot_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Select * From am_stokgudang Where nolot = '" & txtnolot & "'"
        Set RST = OBJ.Execute(SQL)
        
        If RST.EOF Then
            MsgBox "Nomor Lot tidak ditemukan", vbCritical, AppName
            OBJ.Close
            Exit Sub
        End If
        OBJ.Close
    End If
End Sub

Private Sub showdata()
    hapusgrid1
    OBJ.Open dsn
    'SQL = "Select distinct kodebarang,namabarang,kg,hppperkg,kdsatuan,satuan from am_stokgudang where nolot='" & txtnolot & "'"
    'SQL = SQL + " and keterangan='Produksi Lem' and kodebarang= '" & txtkode & "'"
    
    SQL = "Select distinct a.kodebarang,b.namabarang,a.kg,a.hppperkg,b.kdsatuan,b.satuan from list_hpp_produksi a"
    SQL = SQL + " left join am_stokgudang b on a.nolot=b.nolot and a.kodebarang=b.kodebarang where a.nolot='" & txtnolot & "'"
    SQL = SQL + " and a.kodebarang= '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid1.Row = 1
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 0) = RST!kodebarang
        grid1.TextMatrix(grid1.Row, 1) = RST!NamaBarang
        grid1.TextMatrix(grid1.Row, 2) = Format(RST!kg, "##,##0.00")
        grid1.TextMatrix(grid1.Row, 3) = Format(RST!hppperkg, "##,##0.00")
        grid1.TextMatrix(grid1.Row, 4) = RST!kdsatuan
        grid1.TextMatrix(grid1.Row, 5) = RST!satuan
        txtsatuan = RST!satuan
        lblKsatuan = RST!kdsatuan
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
    txtqty.SetFocus
End Sub

Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
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
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
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
        If cmbtype = "07" Then
            SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodeproduk = 'C999'"
        Else
            SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        End If
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
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
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            
            grid.SetFocus
            grid.Col = 5
        Else
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
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
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    
                    OBJ.Close
    
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    carisql1 = "select kodebarang, namabarang from am_itemmst"
                    namatabel = "Item"
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
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
                    
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
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
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

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Function tanggaloz()
    tanggaloz = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtsatuan.SetFocus
    End If
End Sub

Private Sub txtqty_LostFocus()
    txtsatuan.SetFocus
End Sub

Private Sub txtsatuan_Click()
    namatabel = "Satuan."
    carisql1 = "Select a.kodesatuan,b.namasatuan from am_itemdtl a inner join am_unit b"
    carisql1 = carisql1 + " on a.kodesatuan=b.kodesatuan Where a.kodebarang = '" & txtkode & "'"
    frmsearch.Show vbModal
End Sub

Private Sub txtsatuan_GotFocus()
    If hasil = "" Then Exit Sub
    lblKsatuan = hasil
    txtsatuan = hasil1
    hasil = "": hasil1 = ""
End Sub

Private Sub txtsatuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtsatuan_Click
    End If
End Sub
