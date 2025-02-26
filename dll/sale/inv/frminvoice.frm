VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frminvoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Faktur Penjualan"
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
   Icon            =   "frminvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
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
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":2372
      Caption         =   "frminvoice.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":23FE
      Keys            =   "frminvoice.frx":241C
      Spin            =   "frminvoice.frx":245E
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
      EditMode        =   1
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
      ValueVT         =   6946821
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
      Left            =   4680
      Picture         =   "frminvoice.frx":2486
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
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
      Left            =   4920
      Picture         =   "frminvoice.frx":27D4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
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
      Left            =   4440
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
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
      Calculator      =   "frminvoice.frx":2AB6
      Caption         =   "frminvoice.frx":2AD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":2B42
      Keys            =   "frminvoice.frx":2B60
      Spin            =   "frminvoice.frx":2BA2
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
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frminvoice.frx":2BCA
      Caption         =   "frminvoice.frx":2BEA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":2C56
      Keys            =   "frminvoice.frx":2C74
      Spin            =   "frminvoice.frx":2CB6
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
      ValueVT         =   2001469445
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
      Calculator      =   "frminvoice.frx":2CDE
      Caption         =   "frminvoice.frx":2CFE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":2D6A
      Keys            =   "frminvoice.frx":2D88
      Spin            =   "frminvoice.frx":2DCA
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
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   27
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
      MICON           =   "frminvoice.frx":2DF2
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
      TabIndex        =   28
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frminvoice.frx":310C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtotal 
      Height          =   255
      Left            =   7440
      TabIndex        =   29
      Top             =   2520
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":3426
      Caption         =   "frminvoice.frx":3446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":34B2
      Keys            =   "frminvoice.frx":34D0
      Spin            =   "frminvoice.frx":3512
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn1 
      Height          =   255
      Left            =   7440
      TabIndex        =   30
      Top             =   2160
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":353A
      Caption         =   "frminvoice.frx":355A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":35C6
      Keys            =   "frminvoice.frx":35E4
      Spin            =   "frminvoice.frx":3626
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
   Begin TDBNumber6Ctl.TDBNumber txtdiscount 
      Height          =   255
      Left            =   7680
      TabIndex        =   31
      Top             =   1440
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":364E
      Caption         =   "frminvoice.frx":366E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":36DA
      Keys            =   "frminvoice.frx":36F8
      Spin            =   "frminvoice.frx":373A
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
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   7440
      TabIndex        =   32
      Top             =   1200
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":3762
      Caption         =   "frminvoice.frx":3782
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":37EE
      Keys            =   "frminvoice.frx":380C
      Spin            =   "frminvoice.frx":384E
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
      MICON           =   "frminvoice.frx":3876
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
      MICON           =   "frminvoice.frx":3B90
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
      MICON           =   "frminvoice.frx":3EAA
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
      TabIndex        =   39
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Surat Jalan"
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
      MICON           =   "frminvoice.frx":41C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtpotong 
      Height          =   255
      Left            =   7710
      TabIndex        =   41
      Top             =   960
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":44DE
      Caption         =   "frminvoice.frx":44FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":456A
      Keys            =   "frminvoice.frx":4588
      Spin            =   "frminvoice.frx":45CA
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
   Begin TDBNumber6Ctl.TDBNumber txtbruto 
      Height          =   255
      Left            =   7440
      TabIndex        =   43
      Top             =   720
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":45F2
      Caption         =   "frminvoice.frx":4612
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":467E
      Keys            =   "frminvoice.frx":469C
      Spin            =   "frminvoice.frx":46DE
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
   Begin TDBNumber6Ctl.TDBNumber txtDpp 
      Height          =   255
      Left            =   7680
      TabIndex        =   48
      Top             =   1920
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":4706
      Caption         =   "frminvoice.frx":4726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":4792
      Keys            =   "frminvoice.frx":47B0
      Spin            =   "frminvoice.frx":47F2
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
      TabIndex        =   50
      Top             =   1680
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frminvoice.frx":481A
      Caption         =   "frminvoice.frx":483A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoice.frx":48A6
      Keys            =   "frminvoice.frx":48C4
      Spin            =   "frminvoice.frx":4906
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
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "Total Net"
      Height          =   255
      Left            =   6990
      TabIndex        =   49
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "DPP"
      Height          =   255
      Left            =   6990
      TabIndex        =   47
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Salesman"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1320
      TabIndex        =   45
      Top             =   2910
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "Bruto"
      Height          =   255
      Left            =   6990
      TabIndex        =   44
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "Potongan"
      Height          =   255
      Left            =   6990
      TabIndex        =   42
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblsat 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   5730
      Width           =   6105
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6990
      TabIndex        =   33
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Netto"
      Height          =   255
      Left            =   6990
      TabIndex        =   37
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "Discount"
      Height          =   255
      Left            =   6990
      TabIndex        =   36
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000014&
      Caption         =   "PPN"
      Height          =   255
      Left            =   6990
      TabIndex        =   35
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   2550
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "PPN (%)"
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
      Height          =   255
      Left            =   6840
      TabIndex        =   20
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Barang :"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5490
      Width           =   6105
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000011&
      Height          =   345
      Left            =   6840
      TabIndex        =   34
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   6840
      TabIndex        =   38
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frminvoice"
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

Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim SP As New ADODB.Command
Dim vsp(2) As Variant

Dim str1, str2, str3, str4 As String
Dim posrow, poscol As String
Dim dpp, ppn As Double

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    'If txtppn <> 0 And txtppn <> 10 Then
        'MsgBox "PPn <> 10.", vbExclamation, "Warning"
        'txtppn.SetFocus
        'Exit Sub
    'End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not add, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close

    txtnobukti = Trim(txtnobukti)

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

    OBJ.Open dsn
    SQL = "select * from am_invhdr where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click

        Exit Sub
    End If
    
    SQL = "select * from am_invhdr where nosj = '" & txtposup & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Invalid SJ, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click

        Exit Sub
    End If
    OBJ.Close

    OBJ.Open dsn
    SQL = "insert into am_invhdr ("
    SQL = SQL + "nosj, "
    SQL = SQL + "noapply, "
    SQL = SQL + "type, "
    SQL = SQL + "nobkt, "
    SQL = SQL + "tglbkt, "
    SQL = SQL + "kodecust, "
    SQL = SQL + "namacust, "
    SQL = SQL + "alamatcust, "
    SQL = SQL + "kodesales, "
    SQL = SQL + "discprc, "
    SQL = SQL + "discamt, "
    SQL = SQL + "ppn, "
    SQL = SQL + "ppnbm, "
    SQL = SQL + "termpay, "
    SQL = SQL + "identry, "
    SQL = SQL + "dateentry, "
    SQL = SQL + "idupdate, "
    SQL = SQL + "dateupdate, "
    SQL = SQL + "Posted, "
    SQL = SQL + "kodecur, "
    SQL = SQL + "nilaikurs, "
    SQL = SQL + "noseri)"

    SQL = SQL + " values('" & txtposup & "',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "'I',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtkodecust & "',"
    SQL = SQL + "'" & txtNama & "',"
    SQL = SQL + "'" & txtAlamat & "',"
    SQL = SQL + "'" & txtsales & "',"
    SQL = SQL + "convert(money,'0'),"
    SQL = SQL + "convert(money,'0'),"
    SQL = SQL + "convert(money,'" & txtppn & "'),"
    SQL = SQL + "convert(money,'0'),"
    SQL = SQL + "'0',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'0',"
    SQL = SQL + "'" & txtkurs & "',"
    SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
    SQL = SQL + "'')"
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

    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    'SP.CommandText = "am_postinginv" 'Kembalikan jika hitungan ppn normal kembali tanpa 11/12
    SP.CommandText = "am_postinginv_12"
    vsp(0) = txtnobukti
    vsp(1) = Format(date1, "yyyyMMdd")
    vsp(2) = "sj"
    SP.Execute , vsp
    Set SP = Nothing

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti = ""
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1 = Date
    End If
    txtkodecust = ""
    txtNama = ""
    txtAlamat = ""
    txtposup = ""
    txtsales = ""
    lblsales = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
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
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecust = hasil
    caricustomer
    hasil = ""
    hasil1 = ""
    txtNama.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select distinct a.nosj, convert(char(11),a.tglsj )'tglsj' from am_sjapp a where a.kodecust = '" & txtkodecust & "' and a.tglsj >= '" & batas1 & "' and a.tglsj <= '" & batas2 & "' and a.tglsj <= '" & tanggalinv & "' and (select count(b.nosj)'no' from am_invhdr b where b.nosj=a.nosj)=0"
    Else
        carisql1 = "select distinct a.nosj, convert(char(11),a.tglsj )'tglsj' from am_sjapp a where a.kodecust = '" & txtkodecust & "' and a.tglsj <= '" & tanggalinv & "' and (select count(b.nosj)'no' from am_invhdr b where b.nosj=a.nosj)=0"
    End If
    namatabel = "Surat Jalan "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
        
    txtposup = hasil
    caripo
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub date1_Change()
    hapusemua
    
    txtkodecust = ""
    txtNama = ""
    txtAlamat = ""
    txtposup = ""
    txtsales = ""
    lblsales = ""
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
    
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='171' and b.kodeuser = '1" & kuser & "'"
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
    
    SQL = "Select * From am_ppn Where tahun=Year(getdate())"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        dpp = RST!dpp
        ppn = RST!ppn
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

Private Sub txtkodecust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub txtkodecust_LostFocus()
    caricustomer
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
    If KeyAscii = 13 Then date1.SetFocus
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
        txtkurs.SetFocus
    End If
End Sub

Private Sub txtppn_LostFocus()
    'If txtneto = 0 Then Exit Sub
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

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Function tanggalinv()
      tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
      tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub

    hapusemua
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1 = Date
    End If
    txtkodecust = ""
    txtNama = ""
    txtAlamat = ""
    txtsales = ""
    lblsales = ""
    txtposup = ""

    OBJ.Open dsn
    SQL = "select * from am_invhdr where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub caricustomer()
    If txtkodecust = "" Then Exit Sub
    
    hapusemua
    txtposup = ""
    txtsales = ""
    lblsales = ""

    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtNama = RST!namacust
        txtAlamat = RST!alamatcust
    Else
        MsgBox "Customer " & txtkodecust & " Not Found.", vbExclamation, "Warning"
        txtkodecust = ""
        txtNama = ""
        txtAlamat = ""
        txtkodecust.SetFocus

        cmdsearch1_Click
        If hasil <> "" Then
            txtkodecust = hasil

            SQL = "SELECT * FROM AM_customer WHERE Kodecust = '" & txtkodecust & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                txtNama = RST!namacust
                txtAlamat = RST!alamatcust
            End If

            hasil = ""
            txtposup.SetFocus
        Else
            txtkodecust = ""
            txtkodecust.SetFocus
        End If
    End If
    OBJ.Close
End Sub

Private Sub caripo()
    If txtnobukti = "" Or txtposup = "" Then Exit Sub
    
    hapusemua
    txtsales = ""
    lblsales = ""

    OBJ.Open dsn
    SQL = "select a.*,"
    SQL = SQL + "(select b.konversi from am_itemdtl b where a.kodebarang=b.kodebarang and a.kodesatuan=b.kodesatuan)'konversi'"
    SQL = SQL + " from am_sjapp a where a.nosj = '" & txtposup & "' and flag2<>'9' order by a.lineitem asc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtsales = RST!kodesales
        
        grid.Row = 1
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 5) = Format((RST!qty * RST!konversi), "###,###,##0.00")
            grid.TextMatrix(grid.Row, 6) = "0.00"
            grid.TextMatrix(grid.Row, 7) = "0.00"
            grid.TextMatrix(grid.Row, 8) = Format((RST!bn * RST!konversi), "###,###,##0.00")
            grid.TextMatrix(grid.Row, 9) = "0.00"
            
            OBJ1.Open dsn
            SQL1 = "select * from am_itemdtl where kodebarang = '" & RST!kodebarang & "' and level_ = 0"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                grid.TextMatrix(grid.Row, 3) = RST1!kodesatuan
            End If
            OBJ1.Close
            
            SetRow grid.Row, True
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            
            RST.MoveNext
        Loop
        
        SQL = "select * from am_salesman where kodesales = '" & txtsales & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsales = RST!namasales
    End If
    lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    OBJ.Close
End Sub

Private Sub carikurs()
    If txtkurs = "" Then Exit Sub

    OBJ.Open dsn
    SQL2 = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
    Set RST2 = OBJ.Execute(SQL2)
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
    OBJ.Close
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
    txtneto = txtbruto - txtpotong
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
    'txtotal = txtDpp + txtppn1
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture

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
