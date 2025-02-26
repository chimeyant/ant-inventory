VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frminvoicedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Faktur Penjualan"
   ClientHeight    =   6045
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
   Icon            =   "frminvoicedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkdevelop 
      Alignment       =   1  'Right Justify
      Caption         =   "Mode Edit"
      Height          =   255
      Left            =   8040
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtgudang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   48
      Top             =   810
      Width           =   2070
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox txtAlamat 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtsales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin TDBNumber6Ctl.TDBNumber txtotal 
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   2520
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":2372
      Caption         =   "frminvoicedit.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":23FE
      Keys            =   "frminvoicedit.frx":241C
      Spin            =   "frminvoicedit.frx":245E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483631
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   16777215
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin VB.TextBox txtposup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":2486
      Caption         =   "frminvoicedit.frx":24A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":2512
      Keys            =   "frminvoicedit.frx":2530
      Spin            =   "frminvoicedit.frx":2572
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
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
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
      Left            =   3960
      Picture         =   "frminvoicedit.frx":259A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
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
      Left            =   4200
      Picture         =   "frminvoicedit.frx":28E8
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
      Left            =   3720
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
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
      Format          =   143327235
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frminvoicedit.frx":2BCA
      Caption         =   "frminvoicedit.frx":2BEA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":2C56
      Keys            =   "frminvoicedit.frx":2C74
      Spin            =   "frminvoicedit.frx":2CB6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0.00;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   20
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -1355808763
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtdisc 
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frminvoicedit.frx":2CDE
      Caption         =   "frminvoicedit.frx":2CFE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":2D6A
      Keys            =   "frminvoicedit.frx":2D88
      Spin            =   "frminvoicedit.frx":2DCA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0.00;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2175
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frminvoicedit.frx":2DF2
      Caption         =   "frminvoicedit.frx":2E12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":2E7E
      Keys            =   "frminvoicedit.frx":2E9C
      Spin            =   "frminvoicedit.frx":2EDE
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Currency"
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
      MICON           =   "frminvoicedit.frx":2F06
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
      TabIndex        =   29
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frminvoicedit.frx":3220
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtdiscount 
      Height          =   255
      Left            =   7680
      TabIndex        =   32
      Top             =   1560
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":353A
      Caption         =   "frminvoicedit.frx":355A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":35C6
      Keys            =   "frminvoicedit.frx":35E4
      Spin            =   "frminvoicedit.frx":3626
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   7440
      TabIndex        =   33
      Top             =   1320
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":364E
      Caption         =   "frminvoicedit.frx":366E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":36DA
      Keys            =   "frminvoicedit.frx":36F8
      Spin            =   "frminvoicedit.frx":373A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8280
      TabIndex        =   15
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
      MICON           =   "frminvoicedit.frx":3762
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
      TabIndex        =   14
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
      MICON           =   "frminvoicedit.frx":3A7C
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
      Left            =   6360
      TabIndex        =   13
      Top             =   5520
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
      MICON           =   "frminvoicedit.frx":3D96
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
      Left            =   5400
      TabIndex        =   12
      Top             =   5520
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
      MICON           =   "frminvoicedit.frx":40B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtpotong 
      Height          =   255
      Left            =   7710
      TabIndex        =   42
      Top             =   1080
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":43CA
      Caption         =   "frminvoicedit.frx":43EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":4456
      Keys            =   "frminvoicedit.frx":4474
      Spin            =   "frminvoicedit.frx":44B6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtbruto 
      Height          =   255
      Left            =   7440
      TabIndex        =   44
      Top             =   840
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":44DE
      Caption         =   "frminvoicedit.frx":44FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":456A
      Keys            =   "frminvoicedit.frx":4588
      Spin            =   "frminvoicedit.frx":45CA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn1 
      Height          =   240
      Left            =   7440
      TabIndex        =   31
      Top             =   2280
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   432
      Calculator      =   "frminvoicedit.frx":45F2
      Caption         =   "frminvoicedit.frx":4612
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":467E
      Keys            =   "frminvoicedit.frx":469C
      Spin            =   "frminvoicedit.frx":46DE
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   2011627525
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtDpp 
      Height          =   255
      Left            =   7680
      TabIndex        =   51
      Top             =   2040
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":4706
      Caption         =   "frminvoicedit.frx":4726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":4792
      Keys            =   "frminvoicedit.frx":47B0
      Spin            =   "frminvoicedit.frx":47F2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtNetotal 
      Height          =   255
      Left            =   7680
      TabIndex        =   52
      Top             =   1800
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoicedit.frx":481A
      Caption         =   "frminvoicedit.frx":483A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoicedit.frx":48A6
      Keys            =   "frminvoicedit.frx":48C4
      Spin            =   "frminvoicedit.frx":4906
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      ValueVT         =   -1355808763
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000014&
      Caption         =   "DPP"
      Height          =   255
      Left            =   6960
      TabIndex        =   54
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "Total Net"
      Height          =   255
      Left            =   6960
      TabIndex        =   53
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Gudang"
      Height          =   255
      Left            =   3045
      TabIndex        =   49
      Top             =   825
      Width           =   810
   End
   Begin VB.Label Label8 
      Caption         =   "Salesman"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1320
      TabIndex        =   46
      Top             =   2910
      Width           =   5295
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6990
      TabIndex        =   34
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000011&
      Height          =   345
      Left            =   6840
      TabIndex        =   35
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "Bruto"
      Height          =   255
      Left            =   6990
      TabIndex        =   45
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "Potongan"
      Height          =   255
      Left            =   6990
      TabIndex        =   43
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblsat 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   5730
      Width           =   5385
   End
   Begin VB.Label Label1 
      Caption         =   "Surat Jalan"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Netto"
      Height          =   255
      Left            =   6990
      TabIndex        =   38
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "Discount"
      Height          =   255
      Left            =   6990
      TabIndex        =   37
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000014&
      Caption         =   "PPN"
      Height          =   255
      Left            =   6990
      TabIndex        =   36
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   2550
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   270
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "PPN    %"
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Barang :"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   5490
      Width           =   5385
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   6840
      TabIndex        =   39
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frminvoicedit"
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

Dim SP As New ADODB.Command
Dim vsp(2) As Variant

Dim str1, str2, str3, str4 As String
Dim posrow, poscol As String

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    'If txtppn <> 0 Then
        'If txtppn <> 10 Then
            'MsgBox "PPn <> 10.", vbExclamation, "Warning"
            'txtppn.SetFocus
            'Exit Sub
        'End If
    'End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_invhdr where nobkt = '" & txtnobukti & "' and posted='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If chkdevelop.Value = Checked Then GoTo langkahi:
        OBJ.Close
        MsgBox "Can not update, data already process.", vbExclamation, "Warning"
        Exit Sub
langkahi:
    End If
    OBJ.Close
    
    If txtsales = "" Or txtnobukti = "" Or txtposup = "" Or txtkurs = "" Or txtnilaikurs = 0 Or txtNama = "" Or txtAlamat = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        hitamount
        grid.Row = grid.Row + 1
    Loop
    hitneto
    
    If txtneto = 0 Then
        MsgBox "There Is No Data To Save.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        If grid.TextMatrix(grid.Row, 3) = "" Or grid.TextMatrix(grid.Row, 6) = "0.00" Then
            MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    'cek pm
    OBJ.Open dsn
    SQL = "select * from am_aropnfil where noapply = '" & txtnobukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If chkdevelop.Value = Checked Then GoTo langkahi2:
        MsgBox "Can not update data, invoice already have payment.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
langkahi2:
    End If
    OBJ.Close
    'cek cndn
    OBJ.Open dsn
    SQL = "select * from am_aropnfil where noapply = '" & txtnobukti & "' and noapply <> nobkt and (transtype = 'CN' or transtype = 'DN')"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not update data, invoice already have correction.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    ops_tf = False
    frminvoicedesc.Show 1
    If Not ops_tf Then
        MsgBox "Save aborted, user must supply a description/comment for update.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "update am_invhdr set nilaikurs = convert(money,'" & txtnilaikurs & "'),namacust = '" & txtNama & "',alamatcust = '" & txtAlamat & "',idupdate = '" & kuser & "',dateupdate = convert(datetime,'" & tanggalsekarang & "'),ppn = convert(money,'" & txtppn & "') where nobkt = '" & txtnobukti & "' and type = 'I'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_invlin where nobkt = '" & txtnobukti & "'  and type = 'I'"
    Set RST = OBJ.Execute(SQL)
        
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "insert into am_invlin (type, nobkt, kodebarang, qty, price, lineitem, kodesatuan, bn, discline)"
        SQL = SQL + " values('I','" & txtnobukti & "','" & grid.TextMatrix(grid.Row, 1) & "',convert(money,'" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'),convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "'),convert(numeric,'" & grid.Row & "'),'" & grid.TextMatrix(grid.Row, 3) & "',convert(money,'" & Format(grid.TextMatrix(grid.Row, 8), "general number") & "'),convert(money,'" & Format(grid.TextMatrix(grid.Row, 9), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "delete from am_aropnfil where nobkt = '" & txtnobukti & "' and transtype = 'I'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    'SP.CommandText = "am_postinginv"   'Kembalikan Jika hitungan ppn normal kembali tanpa 11/12
    SP.CommandText = "am_postinginv_12"
    vsp(0) = txtnobukti
    vsp(1) = Format(date1, "yyyyMMdd")
    vsp(2) = "sj"
    SP.Execute , vsp
    Set SP = Nothing
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti.Enabled = True
    cmdsearch1.Enabled = True
    
    txtnobukti = ""
    date1.Value = Date
    txtkodecust = ""
    txtNama = ""
    txtAlamat = ""
    txtposup = ""
    txtsales = ""
    lblsales = ""
    txtgudang = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_invhdr where nobkt = '" & txtnobukti & "' and posted='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, data already process.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close

    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    'cek pm
    OBJ.Open dsn
    SQL = "select * from am_aropnfil where noapply = '" & txtnobukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete data, invoice already have payment.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    'cek cndn
    OBJ.Open dsn
    SQL = "select * from am_aropnfil where noapply = '" & txtnobukti & "' and noapply <> nobkt and (transtype = 'CN' or transtype = 'DN')"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete data, invoice already have correction.", vbOKOnly + vbExclamation, "Warning"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    ops_tf = False
    frminvoicedesc2.Show 1
    If Not ops_tf Then
        MsgBox "Save aborted, user must supply a description/comment for delete.", vbExclamation, "Warning"
        Exit Sub
    End If

    OBJ.Open dsn
    SQL = "delete from am_aropnfil where nobkt = '" & txtnobukti & "' and transtype = 'I'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete am_invhdr where type = 'I' and nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete am_invlin where type = 'I' and nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs = hasil
    carikurs
    hasil = ""
    txtnilaikurs.SetFocus
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from am_invhdr where type = 'I' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from am_invhdr where type = 'I'"
    End If
    
    namatabel = "Penjualan"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    carinvoice
    hasil = ""
    hasil1 = ""
    txtkodecust.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='172' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='173' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    If nmuser = "Creator" Then chkdevelop.Visible = True
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 3) = "Satuan"
    grid.TextMatrix(0, 5) = "Quantity"
    grid.TextMatrix(0, 6) = "Price"
    grid.TextMatrix(0, 7) = "Amount"
    grid.TextMatrix(0, 8) = "Bonus"
    grid.TextMatrix(0, 9) = "Disc (%)"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 700
    grid.ColWidth(9) = 1000
    
    grid.RowHeightMin = 300
        
    date1.Value = Date
    
    OBJ.Open dsn
    SQL = "select * from gl_kurs where base = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblbase = "1"
        txtkurs = RST!kdkurs
        txtnilaikurs = 1
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 3) & "'"
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
    
    If txtnobukti = "" Or txtkodecust = "" Or txtposup = "" Then Exit Sub
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 6, 9
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
    If txtnobukti = "" Or txtkodecust = "" Or txtposup = "" Then Exit Sub
    Select Case grid.Col
    Case 6, 9
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
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

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub txtkodecust_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkodecust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNama.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
End Sub

Private Sub txtkurs_LostFocus()
    carikurs
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamat.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        hitamount
        hitneto
        grid.Row = posrow
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecust.SetFocus
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Private Sub txtposup_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtposup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtppn.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtppn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If txtppn <> 0 Then
            'If txtppn <> 10 Then
                'MsgBox "PPn <> 10.", vbExclamation, "Warning"
                'txtppn = 0
                'txtppn.SetFocus
                'Exit Sub
            'End If
        'End If
        If txtppn <> 0 Then
            txtppn1 = ((txtneto - txtdiscount) * txtppn) / 100
        Else
            txtppn1 = 0
        End If
        txtotal = txtneto - txtdiscount + txtppn1
        txtkurs.SetFocus
    End If
End Sub

Private Sub txtppn_LostFocus()
    If txtneto = 0 Then Exit Sub
    'If txtppn <> 0 And txtppn <> 10 Then
        'MsgBox "PPn <> 10.", vbExclamation, "Warning"
        'txtppn = 0
        'txtppn.SetFocus
        'Exit Sub
    'End If
    If txtppn <> 0 Then
        txtppn1 = ((txtneto - txtdiscount) * txtppn) / 100
    Else
        txtppn1 = 0
    End If
    txtotal = txtneto - txtdiscount + txtppn1
End Sub

Function tanggalinv()
      tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
      tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub

    hapusemua
    date1 = Date
    txtkodecust = ""
    txtNama = ""
    txtAlamat = ""
    txtsales = ""
    lblsales = ""
    txtposup = ""
    
    OBJ.Open dsn
    'SQL = "select * from am_invhdr where nobkt = '" & txtnobukti & "' and type = 'I'"
    SQL = "Select a.*,c.NamaGudang from am_invhdr a left join am_sjhdr b on a.nosj = b.NoSJ"
    SQL = SQL + " left join am_gudang c on b.KodeGudang = c.KodeGudang"
    SQL = SQL + " Where a.nobkt = '" & txtnobukti & "' and a.type = 'I'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbkt
        txtkodecust = RST!kodecust
        txtNama = RST!namacust
        txtAlamat = RST!alamatcust
        txtposup = RST!nosj
        txtdisc = RST!discprc
        txtppn = RST!ppn
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtsales = RST!kodesales
        txtgudang = RST!namagudang
                
        SQL = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
        Set RST = OBJ.Execute(SQL)
        If RST!base = 1 Then
            lblbase = "1"
            
            txtbruto.Format = "###,###,###,##0;(###,###,###,##0)"
            txtbruto.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtbruto.Value = 0
            txtpotong.Format = "###,###,###,##0;(###,###,###,##0)"
            txtpotong.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtpotong.Value = 0
            txtneto.Format = "###,###,###,##0;(###,###,###,##0)"
            txtneto.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtneto.Value = 0
            txtdiscount.Format = "###,###,###,##0;(###,###,###,##0)"
            txtdiscount.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtdiscount.Value = 0
            txtppn1.Format = "###,###,###,##0;(###,###,###,##0)"
            txtppn1.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtppn1.Value = 0
            txtotal.Format = "###,###,###,##0;(###,###,###,##0)"
            txtotal.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtotal.Value = 0
        Else
            lblbase = "0"
            
            txtbruto.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtbruto.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtbruto.Value = 0
            txtpotong.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtpotong.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtpotong.Value = 0
            txtneto.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtneto.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtneto.Value = 0
            txtdiscount.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtdiscount.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtdiscount.Value = 0
            txtppn1.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtppn1.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtppn1.Value = 0
            txtotal.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtotal.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtotal.Value = 0
        End If
        
        SQL = "select * from am_salesman where kodesales = '" & txtsales & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsales = RST!namasales
                                
        grid.Row = 1
        SQL = "select * from am_invlin where type = 'I' and nobkt = '" & txtnobukti & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 3) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 5) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 6) = Format(RST!Price, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 7) = Format(RST!qty * RST!Price, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 8) = Format(RST!bn, "###,##0.00")
            grid.TextMatrix(grid.Row, 9) = Format(RST!discline, "##0.00")
                                
            SetRow grid.Row, True
            hitamount
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            
            RST.MoveNext
        Loop
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
        hitneto
        
        txtnobukti.Enabled = False
        cmdsearch1.Enabled = False
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub carikurs()
    If txtkurs = "" Then Exit Sub
    
    OBJ2.Open dsn
    SQL2 = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If Not RST2.EOF Then
        If RST2!base = 1 Then
            lblbase = "1"
            
            txtbruto.Format = "###,###,###,##0;(###,###,###,##0)"
            txtbruto.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtbruto.Value = 0
            txtpotong.Format = "###,###,###,##0;(###,###,###,##0)"
            txtpotong.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtpotong.Value = 0
            txtneto.Format = "###,###,###,##0;(###,###,###,##0)"
            txtneto.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtneto.Value = 0
            txtdiscount.Format = "###,###,###,##0;(###,###,###,##0)"
            txtdiscount.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtdiscount.Value = 0
            txtppn1.Format = "###,###,###,##0;(###,###,###,##0)"
            txtppn1.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtppn1.Value = 0
            txtotal.Format = "###,###,###,##0;(###,###,###,##0)"
            txtotal.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            txtotal.Value = 0
        Else
            lblbase = "0"
            
            txtbruto.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtbruto.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtbruto.Value = 0
            txtpotong.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtpotong.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtpotong.Value = 0
            txtneto.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtneto.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtneto.Value = 0
            txtdiscount.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtdiscount.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtdiscount.Value = 0
            txtppn1.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtppn1.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtppn1.Value = 0
            txtotal.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
            txtotal.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            txtotal.Value = 0
        End If
        hitneto
        
        Select Case Month(date1)
        Case 1
            txtnilaikurs = RST2!kurs1
        Case 2
            txtnilaikurs = RST2!kurs2
        Case 3
            txtnilaikurs = RST2!kurs3
        Case 4
            txtnilaikurs = RST2!kurs4
        Case 5
            txtnilaikurs = RST2!kurs5
        Case 6
            txtnilaikurs = RST2!kurs6
        Case 7
            txtnilaikurs = RST2!kurs7
        Case 8
            txtnilaikurs = RST2!kurs8
        Case 9
            txtnilaikurs = RST2!kurs9
        Case 10
            txtnilaikurs = RST2!kurs10
        Case 11
            txtnilaikurs = RST2!kurs11
        Case 12
            txtnilaikurs = RST2!kurs12
        End Select
    Else
        MsgBox "Currency " & txtkurs & " Not Found.", vbInformation, "Information"
        txtkurs = ""
        txtkurs.SetFocus
    End If
    OBJ2.Close
End Sub

Private Sub hitamount()
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
    str3 = grid.TextMatrix(grid.Row, 5) * grid.TextMatrix(grid.Row, 6)
    grid.TextMatrix(grid.Row, 7) = Format(str3, "###,###,##0.00")
End Sub

Private Sub hitneto()
    grid.Row = 1
    str4 = 0
    str2 = 0
    str1 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do
        If grid.TextMatrix(grid.Row, 7) <> "0.00" Then
            str4 = Val(str4) + (Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) * Val(Format(grid.TextMatrix(grid.Row, 6), "general number")))
        End If
        
        str1 = Val(str1) + (Val(Format(grid.TextMatrix(grid.Row, 8), "general number")) * Val(Format(grid.TextMatrix(grid.Row, 6), "general number")))
        
        If grid.TextMatrix(grid.Row, 9) <> "0.00" Then
            str2 = Val(str2) + (((Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) * Val(Format(grid.TextMatrix(grid.Row, 6), "general number"))) - (Val(Format(grid.TextMatrix(grid.Row, 8), "general number")) * Val(Format(grid.TextMatrix(grid.Row, 6), "general number")))) * 0.01 * Val(Format(grid.TextMatrix(grid.Row, 9), "general number")))
        End If
        
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    txtbruto = str4
    txtpotong = str1
    txtneto = str4 - str1
    txtdiscount = 0
    txtppn1 = 0
    txtotal = 0
    If txtneto = 0 Then Exit Sub
    
    If txtdisc <> 0 Then
        txtdiscount = ((txtneto * txtdisc) / 100) + str2
    Else
        txtdiscount = 0 + str2
    End If
    
    If Left(txtnobukti, 1) = "P" Then
        txtDpp = (txtneto - txtdiscount) * (11 / 12) ' hitung dpp ; total net x 11/12
    Else
        txtDpp = 0
    End If
    txtNetotal = txtneto - txtdiscount
    
    If txtppn <> 0 Then
        'txtppn1 = ((txtneto - txtdiscount) * txtppn) / 100
        If txtppn <> 12 Then
            txtppn1 = ((txtneto - txtdiscount) * txtppn) / 100
        Else
            txtppn1 = (txtDpp * txtppn) / 100
        End If
    Else
        txtppn1 = 0
    End If
    txtotal = txtneto - txtdiscount + txtppn1
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
    txtdisc = 0
    txtppn = 0
    txtneto = 0
    txtbruto = 0
    txtpotong = 0
    txtdiscount = 0
    txtppn1 = 0
    txtotal = 0
    txtkurs = ""
    txtnilaikurs = 0
    txtNetotal = 0
    txtDpp = 0
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
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 700
    grid.ColWidth(9) = 1000
End Sub
