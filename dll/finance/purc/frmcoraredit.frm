VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmcoraredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Koreksi Hutang"
   ClientHeight    =   4620
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   6255
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
   Icon            =   "frmcoraredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtreklawan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cmbtype 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3120
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   130416641
      CurrentDate     =   38516
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin TDBNumber6Ctl.TDBNumber txtjumlah 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":2372
      Caption         =   "frmcoraredit.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":23FE
      Keys            =   "frmcoraredit.frx":241C
      Spin            =   "frmcoraredit.frx":2466
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBText6Ctl.TDBText txtbukti 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmcoraredit.frx":248E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":24FA
      Key             =   "frmcoraredit.frx":2518
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
      MaxLength       =   20
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
   Begin VB.TextBox txtsup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   960
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
      Format          =   130416643
      CurrentDate     =   37421
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2520
      Width           =   4575
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No. Bukti"
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
      MICON           =   "frmcoraredit.frx":255C
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
      Left            =   5160
      TabIndex        =   16
      Top             =   4080
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
      MICON           =   "frmcoraredit.frx":2876
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
      Left            =   4200
      TabIndex        =   15
      Top             =   4080
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
      MICON           =   "frmcoraredit.frx":2B90
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
      Left            =   3240
      TabIndex        =   14
      Top             =   4080
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
      MICON           =   "frmcoraredit.frx":2EAA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtkoreksi 
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":31C4
      Caption         =   "frmcoraredit.frx":31E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3250
      Keys            =   "frmcoraredit.frx":326E
      Spin            =   "frmcoraredit.frx":32B8
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBNumber6Ctl.TDBNumber txtotal 
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   885
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmcoraredit.frx":32E0
      Caption         =   "frmcoraredit.frx":3300
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":336C
      Keys            =   "frmcoraredit.frx":338A
      Spin            =   "frmcoraredit.frx":33CC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483631
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   16777215
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
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
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   360
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmcoraredit.frx":33F4
      Caption         =   "frmcoraredit.frx":3414
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3480
      Keys            =   "frmcoraredit.frx":349E
      Spin            =   "frmcoraredit.frx":34E0
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
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
   Begin TDBNumber6Ctl.TDBNumber txtdiscount 
      Height          =   240
      Left            =   4440
      TabIndex        =   28
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   423
      Calculator      =   "frmcoraredit.frx":3508
      Caption         =   "frmcoraredit.frx":3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3594
      Keys            =   "frmcoraredit.frx":35B2
      Spin            =   "frmcoraredit.frx":35F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
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
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   4080
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
      MICON           =   "frmcoraredit.frx":361C
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
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":3936
      Caption         =   "frmcoraredit.frx":3956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":39C2
      Keys            =   "frmcoraredit.frx":39E0
      Spin            =   "frmcoraredit.frx":3A22
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtjumlahppn 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":3A4A
      Caption         =   "frmcoraredit.frx":3A6A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3AD6
      Keys            =   "frmcoraredit.frx":3AF4
      Spin            =   "frmcoraredit.frx":3B3E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBNumber6Ctl.TDBNumber txtkoreksippn 
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":3B66
      Caption         =   "frmcoraredit.frx":3B86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3BF2
      Keys            =   "frmcoraredit.frx":3C10
      Spin            =   "frmcoraredit.frx":3C5A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   240
      TabIndex        =   41
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Rek/ Lawan"
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
      MICON           =   "frmcoraredit.frx":3C82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblreklawan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   42
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "PPn"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Koreksi PPn"
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   38
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   2280
      TabIndex        =   37
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Netto"
      Height          =   255
      Left            =   3870
      TabIndex        =   32
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3870
      TabIndex        =   30
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Koreksi Nilai"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   3270
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3270
      Width           =   975
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2550
      Width           =   975
   End
   Begin VB.Label lblsup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Type"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "No Invoice"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000011&
      Height          =   345
      Left            =   3720
      TabIndex        =   31
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "Koreksi"
      Height          =   255
      Left            =   3870
      TabIndex        =   33
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3720
      TabIndex        =   34
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmcoraredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str2 As String

Private Sub cmbtype_Click()
    Select Case cmbtype
        Case "Credit Note"
            lbltype = "CN"
        Case "Debit Note"
            lbltype = "DN"
    End Select
    txtbukti.SetFocus
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    If cmbtype = "" Then Exit Sub
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtbukti = ""
    txtapply = ""
    txtkurs = ""
    txtnilaikurs = 0
    lblbase = ""
    txtketerangan = ""
    txtreklawan = ""
    lblreklawan = ""
    txtjumlah = 0
    txtkoreksi = 0
    txtjumlahppn = 0
    txtkoreksippn = 0
End Sub

Private Sub cmdadd_Click()
On Error GoTo Err_handler:
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtsup = "" Or txtreklawan = "" Or txtbukti = "" Or cmbtype = "" Or txtapply = "" Or (txtkoreksi = 0 And txtkoreksippn = 0) Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "SELECT posted FROM am_apcashhdr WHERE kodebayar = '" & lbltype & "' and nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!posted = "1" Then
            MsgBox "Can not update data already posted.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
    End If
    
    If lbltype = "CN" And txtapply = txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtbukti & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply = txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtbukti & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    If lbltype = "CN" And txtapply <> txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtapply & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply <> txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtapply & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not update data already have apply.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!identry
        date2 = RST!dateentry
    End If
    
    SQL = "delete from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashlin where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashlinppn where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apopnfil where nobeli = '" & txtbukti & "' and transtype = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_apCashHdr"
    SQL = SQL + " (Kodesupp"
    SQL = SQL + ", NoBkt"
    SQL = SQL + ", TglBkt"
    SQL = SQL + ", kodebayar"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", Amount"
    SQL = SQL + ", Posted"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", IdEntry"
    SQL = SQL + ", DateEntry"
    SQL = SQL + ", IdUpdate"
    SQL = SQL + ", DateUpdate)"
        
    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ", Convert(Money,'" & txtjumlah & "')"
    SQL = SQL + ", '0'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ", Convert(Money,'" & txtnilaikurs & "')"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal2 & "')"
    SQL = SQL + ", '" & txtreklawan & "'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_apCashLin "
    SQL = SQL + " (NoBkt"
    SQL = SQL + ", tglbkt"
    SQL = SQL + ", KodeBayar"
    SQL = SQL + ", Kodesupp"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Jumlah"
    SQL = SQL + ", selisih"
    SQL = SQL + ", selisihkurs"
    SQL = SQL + ", Potongan)"

    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtsup & "'"
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ",convert(money,'" & txtkoreksi & "')"
    SQL = SQL + ",convert(money,'0')"
    SQL = SQL + ",convert(money,'0')"
    SQL = SQL + ",convert(money,'0'))"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_apCashLinppn "
    SQL = SQL + " (NoBkt"
    SQL = SQL + ", tglbkt"
    SQL = SQL + ", KodeBayar"
    SQL = SQL + ", Kodesupp"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Jumlahppn"
    SQL = SQL + ", koreksippn)"

    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtsup & "'"
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ",convert(money,'" & txtjumlahppn & "')"
    SQL = SQL + ",convert(money,'" & txtkoreksippn & "'))"
    Set RST = OBJ.Execute(SQL)
            
    SQL = "INSERT INTO AM_Apopnfil"
    SQL = SQL + " (Kodesupp"
    SQL = SQL + ", NoBeli"
    SQL = SQL + ", TglBeli"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", TransType"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", Amount"
    SQL = SQL + ", Potongan"
    SQL = SQL + ", selisih"
    SQL = SQL + ", PPN)"

    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
    If lbltype = "DN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksi * -1 & "')"
    If lbltype = "CN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksi & "')"
    SQL = SQL + ",Convert (Money, 0)"
    SQL = SQL + ",Convert (Money, 0)"
    If lbltype = "DN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksippn * -1 & "'))"
    If lbltype = "CN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksippn & "'))"
    OBJ.Execute (SQL)
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdclear_Click()
    txtbukti.Enabled = True
    cmdsearch.Enabled = True
    cmbtype.Enabled = True
    date1.Enabled = True
    
    txtbukti = ""
    cmbtype = ""
    lbltype = ""
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtapply = ""
    txtketerangan = ""
    txtreklawan = ""
    lblreklawan = ""
    txtkurs = ""
    txtnilaikurs = 0
    lblbase = ""
    txtjumlah = 0
    txtkoreksi = 0
    txtjumlahppn = 0
    txtkoreksippn = 0
    cmbtype.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtsup = "" Or txtreklawan = "" Or txtbukti = "" Or cmbtype = "" Or txtapply = "" Or (txtkoreksi = 0 And txtkoreksippn = 0) Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "SELECT posted FROM am_apcashhdr WHERE kodebayar = '" & lbltype & "' and nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!posted = "1" Then
            MsgBox "Can not delete data already posted.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
    End If
            
    If lbltype = "CN" And txtapply = txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtbukti & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply = txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtbukti & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    If lbltype = "CN" And txtapply <> txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtapply & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply <> txtbukti Then SQL = "select * from am_apopnfil where noapply = '" & txtapply & "' and tglbeli >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete data already have apply.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "delete from am_apcashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashlin where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apcashlinppn where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_apopnfil where nobeli = '" & txtbukti & "' and transtype = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_apCashhdr where kodebayar = '" & lbltype & "' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_apCashhdr where kodebayar = '" & lbltype & "'"
    End If
    namatabel = "Koreksi"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    caricor
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Master Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtreklawan = hasil
    lblreklawan = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtjumlah.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='162' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='163' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    cmbtype.AddItem "Debit Note"
    cmbtype.AddItem "Credit Note"
End Sub

Private Sub txtapply_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtketerangan.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    caricor
End Sub

Private Sub txtjumlah_Change()
    txtneto = txtjumlah + txtjumlahppn
    txtdiscount = txtkoreksi + txtkoreksippn
    
    If lbltype = "DN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtjumlahppn_Change()
    txtneto = txtjumlah + txtjumlahppn
    txtdiscount = txtkoreksi + txtkoreksippn
    
    If lbltype = "DN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtreklawan.SetFocus
End Sub

Private Sub txtkoreksi_Change()
    txtneto = txtjumlah + txtjumlahppn
    txtdiscount = txtkoreksi + txtkoreksippn
    
    If lbltype = "DN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtkoreksippn_Change()
    txtneto = txtjumlah + txtjumlahppn
    txtdiscount = txtkoreksi + txtkoreksippn
    
    If lbltype = "DN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtkurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtreklawan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtjumlah.SetFocus
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtreklawan_LostFocus()
    If txtreklawan = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select noac, nmac from gl_masterac where noac = '" & txtreklawan & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblreklawan = RST!nmac
        
        OBJ.Close
    Else
        OBJ.Close
        
        MsgBox "Account " & txtreklawan & " Not Found.", vbExclamation, "Warning"
        txtreklawan = ""
        lblreklawan = ""
        txtreklawan.SetFocus
    End If
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub caricor()
    If txtbukti = "" Or cmbtype = "" Then Exit Sub
    If txtbukti.SelLength <> 0 Then Exit Sub
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtreklawan = ""
    lblreklawan = ""
    txtjumlah = 0
    txtkoreksi = 0
    txtjumlahppn = 0
    txtkoreksippn = 0
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_apCashHdr WHERE NoBkt = '" & txtbukti & "' And kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtsup = RST!kodesupp
        date1 = RST!tglbkt
        txtketerangan = RST!keterangan
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtreklawan = RST!idupdate
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblbase = RST!base
        
        SQL = "select * from gl_masterac where noac = '" & txtreklawan & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblreklawan = RST!nmac
        
        SQL = "SELECT isnull(sum(Amount+potongan+selisih),0)'total',isnull(sum(ppn),0)'totalppn' from AM_Apopnfil WHERE Kodesupp = '" & txtsup & "' AND Nobeli <> '" & txtbukti & "' AND Noapply = '" & txtapply & "' and kodecur = '" & txtkurs & "' and tglbeli <= '" & tanggal1 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtjumlah = RST!Total
            txtjumlahppn = RST!totalppn
        End If
        
        SQL = "SELECT * FROM AM_apCashlin WHERE NoBkt = '" & txtbukti & "' And kodebayar = '" & lbltype & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtkoreksi = RST!jumlah
            txtapply = RST!noapply
        End If
        
        SQL = "SELECT * FROM AM_apCashlinppn WHERE NoBkt = '" & txtbukti & "' And kodebayar = '" & lbltype & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtkoreksippn = RST!koreksippn
        
        SQL = "SELECT * FROM AM_supplier WHERE kodesupp = '" & txtsup & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsup = RST!namasupp
        Else
            lblsup = ""
        End If
    
        txtbukti.Enabled = False
        cmdsearch.Enabled = False
        cmbtype.Enabled = False
        date1.Enabled = False
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

