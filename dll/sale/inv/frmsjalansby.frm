VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsjalansby 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Surat Jalan Surabaya"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   Icon            =   "frmsjalansby.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbuktihapus 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   15
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Surat Jalan langsung dari Pabrik"
      Height          =   255
      Left            =   6360
      TabIndex        =   36
      Top             =   3000
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker date4 
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
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
      Format          =   103415811
      CurrentDate     =   37426
   End
   Begin VB.TextBox txtso 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtgudang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtvia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   200
      TabIndex        =   7
      Top             =   2640
      Width           =   7575
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "frmsjalansby.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalansby.frx":23DE
      Key             =   "frmsjalansby.frx":23FC
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
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmsjalansby.frx":2438
      Caption         =   "frmsjalansby.frx":2458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalansby.frx":24C4
      Keys            =   "frmsjalansby.frx":24E2
      Spin            =   "frmsjalansby.frx":2524
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtsales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
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
      Left            =   5040
      Picture         =   "frmsjalansby.frx":254C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   360
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
      Left            =   5040
      Picture         =   "frmsjalansby.frx":289A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   255
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
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
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
      Format          =   103415811
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
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
      Left            =   8280
      TabIndex        =   14
      Top             =   5520
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
      MICON           =   "frmsjalansby.frx":2B7C
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
      Left            =   7320
      TabIndex        =   13
      Top             =   5520
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
      MICON           =   "frmsjalansby.frx":2E96
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
      Left            =   6360
      TabIndex        =   12
      Top             =   5520
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
      MICON           =   "frmsjalansby.frx":31B0
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
      Left            =   240
      TabIndex        =   31
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmsjalansby.frx":34CA
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
      Left            =   240
      TabIndex        =   33
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Customer"
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
      MICON           =   "frmsjalansby.frx":37E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   35
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Salesman"
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
      MICON           =   "frmsjalansby.frx":3AFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlast 
      Height          =   285
      Left            =   3120
      TabIndex        =   38
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Last Number"
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
      MICON           =   "frmsjalansby.frx":3E18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel 
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmsjalansby.frx":4132
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line3 
      X1              =   9240
      X2              =   9000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   9240
      X2              =   9000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   9240
      X2              =   9240
      Y1              =   600
      Y2              =   3120
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Surat Jalan dari Surabaya tetap dimasukan disini, tetapi tidak akan muncul di faktur. (Tujuan untuk record saja)"
      Height          =   615
      Left            =   5760
      TabIndex        =   37
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Barang : 0"
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Sales Order"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   870
      Width           =   975
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   32
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label lblalamatcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   1560
      Width           =   7575
   End
   Begin VB.Label lblsat 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5730
      Width           =   4065
   End
   Begin VB.Label Label3 
      Caption         =   "Kirim Via"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Kirim"
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "No Surat Jalan"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "No. PO"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1950
      Width           =   1095
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal SJ"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Barang :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5490
      Width           =   4065
   End
End
Attribute VB_Name = "frmsjalansby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow, poscol As String

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    txtnobukti = Trim(txtnobukti)
    
    If txtnobukti = "" Or txtsales = "" Or txtgudang = "" Or txtkodecust = "" Or txtso = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If date4 < date1 Then
        MsgBox "Tanggal Kirim tidak boleh lebih kecil dari Tanggal Surat Jalan.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtapply = "" Then
        If MsgBox("Continue with blank PO number ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Or Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) < Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_sjapp where nosj = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    OBJ.Close
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        OBJ.Open dsn
        If Check1.Value = 0 Then
            SQL = "INSERT INTO AM_sjapp"
            SQL = SQL + " (nosj"
            SQL = SQL + ", Tglsj"
            SQL = SQL + ", kodecust"
            SQL = SQL + ", kodesales"
            SQL = SQL + ", nopo"
            SQL = SQL + ", noso"
            SQL = SQL + ", Kodegudang"
            SQL = SQL + ", tglkirim"
            SQL = SQL + ", via"
            SQL = SQL + ", Kodebarang"
            SQL = SQL + ", qty"
            SQL = SQL + ", keterangan"
            SQL = SQL + ", kodesatuan"
            SQL = SQL + ", lineitem"
            SQL = SQL + ", bn"
            SQL = SQL + ", flag1"
            SQL = SQL + ", flag2)"
        
            SQL = SQL + "VALUES"
            SQL = SQL + " ('" & txtnobukti & "'"
            SQL = SQL + ",Convert(dateTime, '" & tanggalinv & "')"
            SQL = SQL + ", '" & txtkodecust & "'"
            SQL = SQL + ", '" & txtsales & "'"
            SQL = SQL + ", '" & txtapply & "'"
            SQL = SQL + ", '" & txtso & "'"
            SQL = SQL + ", '" & txtgudang & "'"
            SQL = SQL + ",Convert(dateTime, '" & tanggal4 & "')"
            SQL = SQL + ", '" & txtvia & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 1) & "'"
            SQL = SQL + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 2) & "'"
            SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 4) & "'"
            SQL = SQL + ",Convert (numeric, '" & grid.Row & "')"
            SQL = SQL + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "')"
            SQL = SQL + ", '1'"
            SQL = SQL + ", '2')"
            Set RST = OBJ.Execute(SQL)
        End If
        
        SQL = "INSERT INTO AM_sjsby"
        SQL = SQL + " (nosj"
        SQL = SQL + ", Tglsj"
        SQL = SQL + ", kodecust"
        SQL = SQL + ", kodesales"
        SQL = SQL + ", nopo"
        SQL = SQL + ", noso"
        SQL = SQL + ", Kodegudang"
        SQL = SQL + ", tglkirim"
        SQL = SQL + ", via"
        SQL = SQL + ", Kodebarang"
        SQL = SQL + ", qty"
        SQL = SQL + ", keterangan"
        SQL = SQL + ", kodesatuan"
        SQL = SQL + ", lineitem"
        SQL = SQL + ", pabrik"
        SQL = SQL + ", bn)"
    
        SQL = SQL + "VALUES"
        SQL = SQL + " ('" & txtnobukti & "'"
        SQL = SQL + ",Convert(dateTime, '" & tanggalinv & "')"
        SQL = SQL + ", '" & txtkodecust & "'"
        SQL = SQL + ", '" & txtsales & "'"
        SQL = SQL + ", '" & txtapply & "'"
        SQL = SQL + ", '" & txtso & "'"
        SQL = SQL + ", '" & txtgudang & "'"
        SQL = SQL + ",Convert(dateTime, '" & tanggal4 & "')"
        SQL = SQL + ", '" & txtvia & "'"
        SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 1) & "'"
        SQL = SQL + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL = SQL + ", '" & grid.TextMatrix(grid.Row, 4) & "'"
        SQL = SQL + ",Convert (numeric, '" & grid.Row & "')"
        If Check1.Value = 0 Then SQL = SQL + ", '0'" Else SQL = SQL + ", '1'"
        SQL = SQL + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        grid.Row = grid.Row + 1
    Loop
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbuktihapus)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbuktihapus.SetFocus
        Exit Sub
    End If

    txtbuktihapus = Trim(txtbuktihapus)
    
    If txtbuktihapus = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from AM_invhdr where nosj = '" & txtbuktihapus & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Data Already used on invoice (delete aborted), Click OK To Continue ...", vbInformation, "Information"
        txtbuktihapus.SetFocus
        Exit Sub
    Else
        SQL = "delete from AM_sjapp where nosj = '" & txtbuktihapus & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from AM_sjsby where nosj = '" & txtbuktihapus & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    MsgBox "Data Deleted, Click OK To Continue ...", vbInformation, "Information"
    txtbuktihapus = ""
    txtbuktihapus.SetFocus
End Sub

Private Sub cmdlast_Click()
    OBJ.Open dsn
    SQL = "select top 1 nosj from am_sjsby order by nosj desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then MsgBox "Last No : " & RST!nosj, vbInformation, "Information"
    OBJ.Close
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
    namatabel = "Sales"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    lblsales = ""
    txtsales = hasil
    lblsales = hasil1
    hasil = ""
    hasil1 = ""
    carisales
    grid.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    carigudang
    hasil = ""
    hasil1 = ""
    txtvia.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecust = hasil
    caricustomer
    hasil = ""
    hasil1 = ""
    txtapply.SetFocus
End Sub

Private Sub date1_Change()
    date4 = date1
End Sub

Private Sub Form_Activate()
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "The period is empty !!" & vbCrLf & _
        "Please define Period on proces, Starting period date and Ending period date.", vbCritical, "Critical"
        
        OBJ.Close
        Unload Me
        Exit Sub
    End If
    OBJ.Close
    
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Keterangan"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "Satuan"
    grid.TextMatrix(0, 5) = "BN"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 0
    grid.RowHeightMin = 300
    
    date1.Value = Date
    date4.Value = Date
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblitem = "    Nama Barang : " & RST!namabarang
            lblsat = "    Nama Satuan : " & RST!namasatuan
        Else
            lblitem = "    Nama Barang : "
            lblsat = "    Nama Satuan : "
        End If
        OBJ.Close
    End If
    
    If txtnobukti = "" Or txtkodecust = "" Or txtgudang = "" Or txtsales = "" Or txtso = "" Then Exit Sub
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
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
        
            If txtket.Visible = True Then Exit Sub

            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 2, 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            OBJ.Open dsn
            SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If RST!kodeproduk = "C999" Then
                    OBJ.Close
                    Exit Sub
                Else
                    OBJ.Close
                End If
            Else
                OBJ.Close
                Exit Sub
            End If

            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtkodecust = "" Or txtgudang = "" Or txtsales = "" Or txtso = "" Then Exit Sub
    
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub

        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 2, 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 5
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        OBJ.Open dsn
        SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!kodeproduk = "C999" Then
                OBJ.Close
                Exit Sub
            Else
                OBJ.Close
            End If
        Else
            OBJ.Close
            Exit Sub
        End If

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
            grid.TextMatrix(grid.Row, 3) = "0.00"
            grid.TextMatrix(grid.Row, 5) = "0.00"
            
            lblitem = "    Nama Barang : " & RST!namabarang

            SetRow grid.Row, True
            grid.SetFocus
            grid.Col = 2
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
        End If
        OBJ.Close
    ElseIf grid.Col = 4 Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsat = "    Nama Satuan : " & RST!namasatuan
            lblitem = "    Nama Barang : " & RST!namabarang

            grid.SetFocus
            grid.Col = 5
        Else
            lblsat = "    Nama Satuan :"
            lblitem = "    Nama Barang :"
            
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date4.SetFocus
End Sub

Private Sub txtgudang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtvia.SetFocus
End Sub

Private Sub txtgudang_LostFocus()
    carigudang
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
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    grid.TextMatrix(grid.Row, 5) = "0.00"
                    
                    lblitem = "    Nama Barang : " & RST!namabarang
                    OBJ.Close
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""

                    carisql1 = "select kodebarang, namabarang from am_itemmst"
                    namatabel = "Item"

                    frmsearch.Show vbModal
                End If
            Case 2
                grid.SetFocus
                grid.Col = 2
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 2) = txtket
                txtket = ""
                txtket.Visible = False
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
                SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    grid.TextMatrix(posrow, 4) = ""
                    lblsat = "    Nama Satuan :"
                    lblitem = "    Nama Barang :"

                    txtket = ""

                    carisql1 = "SELECT b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b on a.kodesatuan = b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                    namatabel = "Satuan "

                    frmsearch.Show vbModal
                Else
                    lblsat = "    Nama Satuan : " & RST!namasatuan
                    lblitem = "    Nama Barang : " & RST!namabarang
                End If
                OBJ.Close
                grid.Col = 5
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
        txtnilai.Visible = False
    ElseIf KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        
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
    If KeyAscii = 13 Then date1.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Private Sub txtso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal4()
    tanggal4 = Month(date4) & "/" & Day(date4) & "/" & Year(date4)
End Function

Private Sub carigudang()
    If txtgudang = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_gudang where kodegudang = '" & txtgudang & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblgudang = RST!namagudang
    Else
        MsgBox "Gudang " & txtgudang & " Not Found.", vbExclamation, "Warning"
        txtgudang = ""
        lblgudang = ""
        txtgudang.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusemua

    OBJ.Open dsn
    SQL = "select * from am_sjapp where nosj = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnobukti.SetFocus
        txtnobukti = ""
    End If
    OBJ.Close
End Sub

Private Sub caricustomer()
    If txtkodecust = "" Then Exit Sub

    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacust = RST!namacust
        lblalamatcust = RST!alamatcust
    Else
        MsgBox "Customer " & txtkodecust & " Not found.", vbExclamation, "Warning"
        txtkodecust = ""
        lblnamacust = ""
        lblalamatcust = ""
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

Private Sub hapusemua()
    date1 = Date
    txtkodecust = ""
    lblnamacust = ""
    lblalamatcust = ""
    txtsales = ""
    lblsales = ""
    txtgudang = ""
    lblgudang = ""
    txtapply = ""
    txtvia = ""
    txtso = ""
    date4 = Date
    Check1.Value = 0
    
    hapusgrid
    
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
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
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
    grid.ColWidth(6) = 0
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
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
    If grid.Rows = 2 Then
        lbltotal.Caption = "    Total Barang : 0"
    Else
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    End If
End Sub

Private Sub carisales()
    If txtsales = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtsales & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & lblsales & " is not active !", vbExclamation, "Warning"
        lblsales = ""
        txtsales = ""
    End If
    OBJ.Close
End Sub
