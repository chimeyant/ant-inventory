VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpenerimaan_app 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Penerimaan Barang"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
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
   Icon            =   "frmpenerimaan_app.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtalamat 
      Height          =   345
      Left            =   11175
      TabIndex        =   43
      Top             =   1305
      Width           =   1605
   End
   Begin VB.TextBox txtnpwp 
      Height          =   345
      Left            =   11190
      TabIndex        =   42
      Top             =   900
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker date4 
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   1680
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
      Format          =   133824515
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   9840
      TabIndex        =   40
      Top             =   2790
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Format          =   133824515
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   9840
      TabIndex        =   39
      Top             =   3150
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Format          =   133824515
      CurrentDate     =   37426
   End
   Begin VB.TextBox ket4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3300
      Width           =   5295
   End
   Begin VB.TextBox ket3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3570
      Width           =   5295
   End
   Begin VB.TextBox ket2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3030
      Width           =   5295
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2670
      Width           =   5295
   End
   Begin VB.TextBox txtsup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   5295
   End
   Begin VB.TextBox txtref2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtref1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1665
      Width           =   1815
   End
   Begin VB.TextBox txtcurr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   6105
      ItemData        =   "frmpenerimaan_app.frx":2372
      Left            =   120
      List            =   "frmpenerimaan_app.frx":2374
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmpenerimaan_app.frx":2376
      Caption         =   "frmpenerimaan_app.frx":2396
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_app.frx":2402
      Keys            =   "frmpenerimaan_app.frx":2420
      Spin            =   "frmpenerimaan_app.frx":2462
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.0000;(##,###,##0.0000);0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.0000;(##,###,##0.0000)"
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
   Begin VB.TextBox txtpo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   17
      TabIndex        =   2
      Top             =   600
      Width           =   1815
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
      Left            =   7800
      Picture         =   "frmpenerimaan_app.frx":248A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
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
      Left            =   8040
      Picture         =   "frmpenerimaan_app.frx":27D8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
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
      Left            =   7560
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   9840
      TabIndex        =   20
      Top             =   3510
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Format          =   133824515
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   2760
      TabIndex        =   16
      Top             =   3990
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3413
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   6030
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
      MICON           =   "frmpenerimaan_app.frx":2ABA
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
      Left            =   8160
      TabIndex        =   18
      Top             =   6030
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
      MICON           =   "frmpenerimaan_app.frx":2DD4
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
      Left            =   7200
      TabIndex        =   17
      Top             =   6030
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Confirm"
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
      MICON           =   "frmpenerimaan_app.frx":30EE
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
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Submit"
      Top             =   3510
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
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
      MICON           =   "frmpenerimaan_app.frx":3408
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   5160
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "frmpenerimaan_app.frx":3722
      Caption         =   "frmpenerimaan_app.frx":3742
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_app.frx":37AE
      Keys            =   "frmpenerimaan_app.frx":37CC
      Spin            =   "frmpenerimaan_app.frx":380E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
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
      ValueVT         =   2003566597
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   285
      Left            =   7920
      TabIndex        =   10
      Top             =   2040
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmpenerimaan_app.frx":3836
      Caption         =   "frmpenerimaan_app.frx":3856
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_app.frx":38C2
      Keys            =   "frmpenerimaan_app.frx":38E0
      Spin            =   "frmpenerimaan_app.frx":3922
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99
      MinValue        =   0
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   6720
      TabIndex        =   37
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Invoice"
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
      MICON           =   "frmpenerimaan_app.frx":394A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtpotongan 
      Height          =   285
      Left            =   7920
      TabIndex        =   44
      Top             =   2355
      Width           =   1830
      _Version        =   65536
      _ExtentX        =   3228
      _ExtentY        =   503
      Calculator      =   "frmpenerimaan_app.frx":3C64
      Caption         =   "frmpenerimaan_app.frx":3C84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_app.frx":3CF0
      Keys            =   "frmpenerimaan_app.frx":3D0E
      Spin            =   "frmpenerimaan_app.frx":3D50
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   10
      MinValue        =   0
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
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Jumah Potong (RP)"
      Height          =   255
      Left            =   6405
      TabIndex        =   45
      Top             =   2385
      Width           =   1410
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Invoice"
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   1710
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Keterangan PO"
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "PPn (%)"
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Keterangan LPB"
      Height          =   255
      Left            =   2880
      TabIndex        =   35
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   2880
      TabIndex        =   34
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "No. Voucher"
      Height          =   255
      Left            =   2880
      TabIndex        =   33
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   7920
      TabIndex        =   32
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm"
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
      Left            =   2760
      TabIndex        =   31
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblket2 
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label lblket1 
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   6000
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Currency"
      Height          =   255
      Left            =   2880
      TabIndex        =   28
      Top             =   2070
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unconfirm"
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
      TabIndex        =   27
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Nomor LPB"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   630
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nomor P.O."
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   1350
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal LPB"
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmpenerimaan_app"
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

Dim str2, str3 As String
Dim posrow, posrow1 As String
Dim i, j As Integer
Dim boo1 As Boolean

Private Sub cmdadd_Click()
Dim Nkurs As String
Dim Rkurs As String
    Rkurs = txtnilaikurs
    If txtcurr <> "IDR" Then
        Nkurs = InputBox("Masukkan nilaikurs !", "Konfirmasi Nilai Kurs", 0)
            If Nkurs = "" Then
                txtnilaikurs = Rkurs
                Exit Sub
            Else
                txtnilaikurs = Nkurs
            End If
    End If
'Exit Sub
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtpo = "" Or txtcurr = "" Or txtnilaikurs = 0 Or txtref2 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        
        Exit Sub
    End If
    
    If txtppn > 0 And txtppn < 11 Then
        MsgBox "PPn Value must 11.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If MsgBox("Are you sure want to confirm (Penerimaan Barang)?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    If str3 = "" Then
        MsgBox "Please check on tabel currency.", vbExclamation, "Warning"
        Exit Sub
    Else
        If str3 = "0" And txtnilaikurs <= 1 Then
            MsgBox "Rate on non base currency must more then 1.", vbExclamation, "Warning"
            Exit Sub
        End If
    End If
    
    OBJ1.Open dsn
    SQL1 = "select * from am_beliapp where ref2 = '" & txtref2 & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If RST1.EOF Then
        If MsgBox("Faktur not found, continue with new one ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ1.Close
            Exit Sub
        End If
    End If
    OBJ1.Close
        
    OBJ1.Open dsn
    SQL1 = "select * from am_apopnfil where nobeli = '" & txtnobukti & "'"
    Set RST1 = OBJ1.Execute(SQL1)

    If Not RST1.EOF Then
        OBJ1.Close
        MsgBox "Can not confirm, please check LPB.", vbExclamation, "Warning"
        Exit Sub
    End If

    SQL = "select * from am_voucherhdr where novoucher='" + txtref1 + "'"
   
    Set RST1 = OBJ1.Execute(SQL)
    If Not RST1.EOF Then
        If MsgBox("Data Telah Ada Apakah anda akan melakukan update data tersebut...?", vbQuestion + vbYesNo, AppName) = vbNo Then
            OBJ1.Close
            Exit Sub
        End If
    End If
    
    SQL1 = "delete from am_beliapp where nobeli = '" & txtnobukti & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    
    SQL1 = "delete from am_apopnfil where nobeli = '" & txtnobukti & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    
    txtref1 = GetNoVoucher
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do

        SQL1 = "INSERT INTO AM_beliapp"
        SQL1 = SQL1 + " (NoBeli"
        SQL1 = SQL1 + ", TglBeli"
        SQL1 = SQL1 + ", nopo"
        SQL1 = SQL1 + ", ref1"
        SQL1 = SQL1 + ", ref2"
        SQL1 = SQL1 + ", kodesupp"
        SQL1 = SQL1 + ", kodecur"
        SQL1 = SQL1 + ", nilaikurs"
        SQL1 = SQL1 + ", Kodebarang"
        SQL1 = SQL1 + ", qty"
        SQL1 = SQL1 + ", Price"
        SQL1 = SQL1 + ", kodesatuan"
        SQL1 = SQL1 + ", keterangan"
        SQL1 = SQL1 + ", keterangan2"
        SQL1 = SQL1 + ", keterangan3"
        SQL1 = SQL1 + ", keterangan4"
        SQL1 = SQL1 + ", ppn"
        SQL1 = SQL1 + ", lineitem"
        SQL1 = SQL1 + ", flag1"
        SQL1 = SQL1 + ", flag2)"
    
        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + " ('" & txtnobukti & "'"
        SQL1 = SQL1 + ",Convert(dateTime, '" & tanggalpo & "')"
        SQL1 = SQL1 + ", '" & txtpo & "'"
        SQL1 = SQL1 + ", '" & txtref1 & "'"
        SQL1 = SQL1 + ", '" & txtref2 & "'"
        SQL1 = SQL1 + ", '" & str2 & "'"
        SQL1 = SQL1 + ", '" & txtcurr & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & txtnilaikurs & "')"
        SQL1 = SQL1 + ", '" & grid.TextMatrix(grid.Row, 1) & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        SQL1 = SQL1 + ",Convert (Money, '" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "')"
        SQL1 = SQL1 + ", '" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL1 = SQL1 + ", '" & txtket & "'"
        SQL1 = SQL1 + ", '" & ket2 & "'"
        SQL1 = SQL1 + ", '" & ket3 & "'"
        SQL1 = SQL1 + ", '" & ket4 & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & txtppn & "')"
        SQL1 = SQL1 + ",Convert (numeric, '" & grid.Row & "')"
        SQL1 = SQL1 + ", '1'"
        SQL1 = SQL1 + ", '1')"
        Set RST1 = OBJ1.Execute(SQL1)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ1.Close
    
    OBJ1.Open dsn
    SQL1 = "select distinct b.nopo,b.kodesupp,b.kodecur,b.nilaikurs,b.ppn,"
    SQL1 = SQL1 + "(select sum(a.qty*a.price) from am_beliapp a where a.nobeli=b.nobeli)'amount'"
    SQL1 = SQL1 + " from am_beliapp b where b.nobeli = '" & txtnobukti & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ.Open dsn
        SQL = "insert into am_apopnfil ("
        SQL = SQL + "kodesupp, "
        SQL = SQL + "nobeli, "
        SQL = SQL + "tglbeli, "
        SQL = SQL + "noapply, "
        SQL = SQL + "transtype, "
        SQL = SQL + "keterangan, "
        SQL = SQL + "amount, "
        SQL = SQL + "potongan, "
        SQL = SQL + "ppn, "
        SQL = SQL + "selisih, "
        SQL = SQL + "kodecur, "
        SQL = SQL + "nilaikurs)"
        
        SQL = SQL + " values ('" & RST1!kodesupp & "',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
        SQL = SQL + "'" & txtref2 & "',"
        SQL = SQL + "'I',"
        SQL = SQL + "'" & Format(date4, "yyyyMMdd") & "',"
        SQL = SQL + "convert(money,'" & RST1!amount & "'),"
        SQL = SQL + "convert(money,'" & txtpotongan & "'),"
        If RST1!ppn = 0 Then SQL = SQL + "convert(money,'0')," Else SQL = SQL + "convert(money,'" & RST1!amount * RST1!ppn * 0.01 & "'),"
        SQL = SQL + "convert(money,'0'),"
        SQL = SQL + "'" & RST1!kodecur & "',"
        SQL = SQL + "convert(money,'" & RST1!nilaikurs & "'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    OBJ1.Close
    
    'simpan ke table no voucher
    OBJ1.Open dsn
    SQL = "Insert into am_voucherhdr("
    SQL = SQL + "novoucher,tgl,kepada,npwp,alamat,kdkurs,nilai,ppn,username,ispajak) VALUES('"
    SQL = SQL + txtref1 + "',"
    SQL = SQL + "convert(datetime,'" + Format(Date, "MM/dd/yyyy") + "'),'"
    SQL = SQL + txtsup + "','"
    SQL = SQL + txtnpwp + "','"
    SQL = SQL + txtalamat + "','"
    SQL = SQL + txtcurr + "',"
    SQL = SQL + "convert(money,'" + Format(txtnilaikurs, "general number") + "'),"
    SQL = SQL + "convert(money,'" + Format(txtppn, "general number") + "'),'"
    SQL = SQL + nmuser + "','1')"
    OBJ1.Execute SQL
    
    'simpan ke table vocherin
    grid.Row = 1
    Do While True
        With grid
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "insert into am_voucherin ("
            SQL = SQL + "novoucher,nonota,tgl,keterangan,perkiraan,jumlah) VALUES('"
            SQL = SQL + txtref1 + "','"
            SQL = SQL + txtref2 + "',"
            SQL = SQL + "convert(datetime,'" + Format(Date, "MM/dd/yyyy") + "'),'"
            SQL = SQL + grid.TextMatrix(.Row, 1) + "','"
            SQL = SQL + "" + "',"
            SQL = SQL + "convert(money,'" + Format(grid.TextMatrix(.Row, 5), "general number") + "'))"
            OBJ1.Execute (SQL)
            .Row = .Row + 1
        End With
    Loop
    
    OBJ1.Close
    
    PrintVoucher
    List1.RemoveItem (i)
    
    MsgBox "Data confirm, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub PrintVoucher()
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_ppn As Long
    Dim nilai_potongan As Long
    Dim nilai_hutang As Long
    
    SQL = "SELECT NoApply,nilaikurs,Amount,Selisih,potongan,(PPN * nilaikurs) AS nilaippn,kodecur, TransType, Amount - Potongan + PPN AS jumlah"
    SQL = SQL + " From am_apopnfil"
    SQL = SQL + " Where NoBeli='" + txtnobukti + "'"

    OBJ1.Open dsn
    Set RST = OBJ1.Execute(SQL)

    Do While Not RST.EOF
        kode_kurs = RST!kodecur
        nilai_kurs = RST!nilaikurs
        nilai_jumlah = RST!amount
        nilai_ppn = RST!nilaippn
        'nilai_ppn = Format(RST!amount, "###,###,##0.00") * RST!nilaikurs * 0.1
        'nilai_ppn = Format(RST!nilaikurs, "###,###,##0.00") * RST!PPN
        nilai_potongan = RST!potongan
        nilai_hutang = RST!jumlah
        RST.MoveNext
    Loop
    OBJ1.Close
        
    SQL = "Select  a.* ,b.namabarang,a.qty , d.namasatuan ,(a.qty * a.price) as jumlah,c.noapply "
    SQL = SQL + " from am_beliapp   as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang "
    SQL = SQL + " inner join am_apopnfil c on c.nobeli= a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    SQL = SQL + " Where a.nobeli = '" + txtnobukti + "'"

    With rptvoucher
        .lblsupp = txtsup
        .lblnpwp = ""
        .lblkurs = kode_kurs
        .lblnilaikurs = Format(nilai_kurs, "###,###,##0.00")
        .lbljumlah = Format(nilai_jumlah, "###,###,##0.00")
        .lblppn = Format(nilai_ppn, "###,###,##0.00")
        .lblpotongan = Format(nilai_potongan, "###,###,##0.00")
        If txtcurr = "IDR" Then
            .lblhutang = Format(nilai_jumlah + nilai_ppn - nilai_potongan, "###,###,##0.00")
        Else
            .lblhutang = .lbljumlah
        End If
        
         grid.Row = 1
        .lblalamat = txtalamat
        .lblnovoucher = ": " + txtref1
        '.lbltanggal = ": " + Format(Date, "dd/MM/yyyy")
        .lbltanggal = ": " + Format(date4, "dd/MM/yyyy")
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub

Private Sub cmdclear_Click()

    hapusgrid
    
    txtnobukti = ""
    date1.Value = Date
    txtpo = ""
    txtcurr = ""
    txtnilaikurs = "0.00"
    Label6 = ""
    txtref1 = ""
    txtref2 = ""
    txtsup = ""
    txtppn = "0.00"
    txtket = ""
    ket2 = ""
    ket3 = ""
    ket4 = ""
    date4.Enabled = True
    date4 = Date
    boo1 = True
    lblket1 = "Nama Barang : "
    lblket2 = "Nama Satuan : "
    txtalamat = ""
    txtnpwp = ""
    txtpotongan = "0.00"
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdpost_Click()
    If List1.text = "" Then Exit Sub

    hapusgrid
    date1 = Date
    txtpo = ""
    txtcurr = ""
    Label6 = ""
    txtsup = ""
    txtppn = 0
    ket2 = ""
    ket3 = ""
    ket4 = ""
    txtref1 = ""
    txtref2 = ""
    txtnilaikurs = 0
    txtnobukti = List1.text
    i = List1.ListIndex
    date4.Enabled = True
    date4 = Date
    boo1 = True
    str2 = ""
    str3 = ""
    
    txtref1.text = GetNoVoucher

    OBJ.Open dsn
    SQL = "select distinct * from am_beliapp where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbeli
        Label6 = Format(RST!tglbeli, "dd MMMM yyyy")
        txtpo = RST!NOPO
        txtppn = RST!ppn
        'txtref1 = RST!ref1
        txtref2 = RST!ref2
        txtcurr = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtket = RST!keterangan
        ket2 = RST!keterangan2
        ket3 = RST!keterangan3
        ket4 = RST!keterangan4
        str2 = RST!kodesupp
        
        SQL = "select * from am_supplier where kodesupp = '" & str2 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtsup = RST!namasupp
            txtalamat = RST!alamatsupp1
        Else
            txtsup = ""
            txtalamat = ""
        End If
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtcurr & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then str3 = RST!base Else str3 = ""
        
        grid.Row = 1
        SQL = "select * from am_beliapp where nobeli = '" & txtnobukti & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.Col = 2
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 2) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 4) = Format(RST!Price, "###,###,##0.0000")
            grid.TextMatrix(grid.Row, 5) = Format(RST!qty * RST!Price, "###,###,##0.00")
            'grid.TextMatrix(grid.Row, 5) = Format(SpyRound(grid.TextMatrix(grid.Row, 5)), "###,###,##0.00")
            grid.Col = 0
            Set grid.CellPicture = uncheck.Picture

            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdsearch1_Click()
    date2 = date1 - 30
    date3 = date1 + 30
    
    carisql1 = "select distinct ref2 from am_beliapp where tglbeli > '" & tanggal2 & "' and tglbeli < '" & tanggal3 & "' and ref2 <> '' and kodecur = '" & txtcurr & "' and kodesupp = '" & str2 & "'"
    namatabel = "Faktur"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtref2 = hasil
    carinofaktur
    hasil = ""
    hasil1 = ""
    If date4.Enabled = True Then date4.SetFocus Else txtppn.SetFocus
End Sub

Private Sub cmdvoucher_Click()
    frmvoucher.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "K/Sat."
    grid.TextMatrix(0, 3) = "Qty"
    grid.TextMatrix(0, 4) = "Price"
    grid.TextMatrix(0, 5) = "Jumlah"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 1500
    
    grid.RowHeightMin = 300
    
    date1.Value = Date
    date4.Value = Date
    
    List1.Clear
    
    OBJ.Open dsn
    SQL = "SELECT distinct nobeli FROM AM_beliapp WHERE flag2 = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nobeli
        RST.MoveNext
    Loop
    OBJ.Close
    
    lblket1 = "Nama Barang : "
    lblket2 = "Nama Satuan : "
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
    OBJ1.Open dsn
    SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket1 = "Nama Barang : " & RST1!namabarang
    If RST1.EOF Then lblket1 = "Nama Barang : "

    SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 2) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket2 = "Nama Satuan : " & RST1!namasatuan
    If RST1.EOF Then lblket2 = "Nama Satuan : "
    OBJ1.Close
    
    posrow = grid.Row
    Select Case grid.Col
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
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
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub ket2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ket2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket4.SetFocus
    KeyAscii = 0
End Sub

Private Sub ket3_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ket3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
    KeyAscii = 0
End Sub

Private Sub ket4_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ket4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket3.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtcurr_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtcurr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtket_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket2.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 4) = Format(txtnilai, "###,###,##0.0000")
        grid.TextMatrix(grid.Row, 5) = Format(Format(grid.TextMatrix(grid.Row, 3), "general number") * Format(grid.TextMatrix(grid.Row, 4), "general number"), "###,###,##0.0000")
    
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai_LostFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not boo1 Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtppn.SetFocus
    If Not boo1 Then KeyAscii = 0
End Sub

Private Sub txtnobukti_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsup.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtpo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtref1.SetFocus
    KeyAscii = 0
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
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 1500
End Sub

Function tanggalpo()
      tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
      tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggal3()
      tanggal3 = Month(date3) & "/" & Day(date3) & "/" & Year(date3)
End Function

Function tanggal4()
      tanggal4 = Month(date4) & "/" & Day(date4) & "/" & Year(date4)
End Function

Private Sub txtppn_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not boo1 Then KeyCode = 0
End Sub

Private Sub txtppn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtket.SetFocus
    If Not boo1 Then KeyAscii = 0
End Sub

Private Sub txtref1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtref2.SetFocus
End Sub

Private Sub txtref2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then If date4.Enabled = True Then date4.SetFocus Else txtppn.SetFocus
End Sub

Private Sub txtref2_LostFocus()
    If txtref2 <> "" Then carinofaktur
End Sub

Private Sub txtsup_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtpo.SetFocus
    KeyAscii = 0
End Sub

Private Sub carinofaktur()
    If txtref2 = "" Then Exit Sub
    
    boo1 = True
    OBJ.Open dsn
    SQL = "select distinct nilaikurs,ppn from am_beliapp where ref2 = '" & txtref2 & "' and kodecur = '" & txtcurr & "' and kodesupp = '" & str2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilaikurs = RST!nilaikurs
        txtppn = RST!ppn
        boo1 = False
        
        SQL = "select distinct keterangan from am_apopnfil where noapply = '" & txtref2 & "' and len(keterangan)=8"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            date4.Year = Mid(RST!keterangan, 1, 4)
            date4.Month = Mid(RST!keterangan, 5, 2)
            date4.Day = Mid(RST!keterangan, 7, 2)
            date4.Enabled = False
        Else
            date4 = Date
            date4.Enabled = True
        End If
    End If
    OBJ.Close
End Sub


Private Function GetNoVoucher() As String
    SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='1'"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    GetNoVoucher = Trim(Str(RST!maxno + 1))
    OBJ.Close
Exit Function
    
    Dim tempyear As String
    Dim temp_kode As String
    Dim int_kode As Long
    tempyear = Format(Date, "yy") & "-"
    
    'SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='1'"
    OBJ.Open dsn
    SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='1' and novoucher like '" & tempyear & "%'"
    Set RST = OBJ.Execute(SQL)
    If RST!maxno = "" Or IsNull(RST!maxno) Then
        temp_kode = "0001"
    End If
        
    If RST!maxno <> "" Then
        'int_kode = RST!maxno
        int_kode = Right(RST!maxno, 4) 'new (yg diubah ini aja)
        int_kode = int_kode + 1
            
        If (Len(Trim(Str(Right(int_kode, 4)))) = 1) Then
            temp_kode = "000" + Trim(Str(Right(int_kode, 1)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 2) Then
            temp_kode = "00" + Trim(Str(Right(int_kode, 2)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 3) Then
            temp_kode = "0" + Trim(Str(Right(int_kode, 3)))
        End If
        If (Len(Trim(Str(Right(int_kode, 4)))) = 4) Then
            temp_kode = Trim(Str(Right(int_kode, 4)))
        End If
    End If
    GetNoVoucher = Format(Date, "yy") & "-" & temp_kode
    OBJ.Close
End Function

Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function
