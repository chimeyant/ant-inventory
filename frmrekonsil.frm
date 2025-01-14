VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmrekonsil 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   17415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtacc1 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7680
      Picture         =   "frmrekonsil.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      Picture         =   "frmrekonsil.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   7695
      _Version        =   851970
      _ExtentX        =   13573
      _ExtentY        =   450
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   -2147483637
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   6120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   720
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
      Format          =   106889219
      CurrentDate     =   37749
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   720
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
      Format          =   106889219
      CurrentDate     =   37749
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   16440
      TabIndex        =   10
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmrekonsil.frx":0630
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton btnview 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "View"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmrekonsil.frx":094A
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
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmrekonsil.frx":0C64
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmrekonsil.frx":0F7E
      Caption         =   "frmrekonsil.frx":0F9E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekonsil.frx":100A
      Keys            =   "frmrekonsil.frx":1028
      Spin            =   "frmrekonsil.frx":106A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      HighlightText   =   1
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
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtjumlah 
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmrekonsil.frx":1092
      Caption         =   "frmrekonsil.frx":10B2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekonsil.frx":111E
      Keys            =   "frmrekonsil.frx":113C
      Spin            =   "frmrekonsil.frx":117E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      HighlightText   =   1
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
      ReadOnly        =   -1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   15480
      TabIndex        =   20
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmrekonsil.frx":11A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtselisih 
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   2040
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmrekonsil.frx":14C0
      Caption         =   "frmrekonsil.frx":14E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekonsil.frx":154C
      Keys            =   "frmrekonsil.frx":156A
      Spin            =   "frmrekonsil.frx":15AC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      HighlightText   =   1
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
      ReadOnly        =   -1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   14520
      TabIndex        =   24
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmrekonsil.frx":15D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblacc 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Selisih                       :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblwait 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000011&
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Nilai Buku Besar        :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Nilai Rekening Koran :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblrecord 
      Appearance      =   0  'Flat
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "  Rekonsiliasi Bank"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1455
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   17415
   End
End
Attribute VB_Name = "frmrekonsil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim norec As String

Private Sub btnview_Click()
    If txtacc1 = "" Then Exit Sub
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    hapusgrid
    lblwait.Visible = True
    DoEvents
    grid.MousePointer = vbHourglass
    opendata
    grid.MousePointer = vbDefault
    lblwait.Visible = False
    If grid.Rows = 1 Then
        grid.Rows = 2
        Adodc1.Refresh
        grid.Refresh
    End If
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    txtacc1 = ""
    txtnilai.Value = "0.00"
    txtjumlah.Value = "0.00"
    txtselisih.Value = "0.00"
    date1 = Date
    date2 = Date
    lblacc = ""
    lblrecord = "0 Record"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_handler:
    Dim baris As Integer
    If lblstatus = "Status : UnBalance" Then
        MsgBox "Unbalance value" + vbCrLf + "Save is abort.", vbCritical, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    grid.MousePointer = vbHourglass
    
    grid.Row = 1
    baris = 0
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.Col = 0
        If grid.CellPicture = check Then
            baris = baris + 1
            SQL = "insert into am_reconsil ("
            SQL = SQL + "noreconsil, "
            SQL = SQL + "tglfrom, "
            SQL = SQL + "tglto, "
            SQL = SQL + "notrans, "
            SQL = SQL + "noacc, "
            SQL = SQL + "line, "
            SQL = SQL + "tgl)"
    
            SQL = SQL + " values ('" & norec & "',"
            SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
            SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
            SQL = SQL + "'" & grid.TextMatrix(grid.Row, 3) & "',"
            SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
            SQL = SQL + "convert(numeric,'" & baris & "'),"
            SQL = SQL + "convert(datetime,'" & tglsekarang & "'))"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "Update gl_transaksi set reconsil='1' "
            SQL = SQL + "Where notrx = '" & grid.TextMatrix(grid.Row, 3) & "' "
            SQL = SQL + "and noactrx = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        If grid.Row = grid.Rows - 1 Then Exit Do
        grid.Row = grid.Row + 1
        DoEvents
    Loop
    OBJ.Close
    
    hapusgrid
    txtnilai.Value = "0.00"
    txtjumlah.Value = "0.00"
    txtselisih.Value = "0.00"
    opendata
    grid.MousePointer = vbDefault
    MsgBox "Reconsiliasi is save successfuly", vbInformation, AppName
    norec = getreconsil
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '01' and a.kdcomp <= '01'"
    namatabel = "Company Account  "
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    lblacc = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    setgrid
    norec = getreconsil
End Sub

Private Sub opendata()
    Adodc1.ConnectionString = dsn
    If txtacc1 = "1410300" Or txtacc1 = "25201001" Then
        SQL = "Select tgltrx,kdtrx,notrx,noactrx,desctrx,dbkrtrx,currtrx,kurs,cekbg,amounttrx From gl_transaksi "
        SQL = SQL + "Where tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
        SQL = SQL + "and kdtrx in('JB','JU','JP','GL') "
        SQL = SQL + "and noactrx = '" & txtacc1 & "' and reconsil in('0','is NULL')"
        SQL = SQL + "Order By notrx desc"
    ElseIf txtacc1 = "11101003" Or txtacc1 = "11101004" Or txtacc1 = "11101006" Then
        SQL = "Select tgltrx,kdtrx,notrx,noactrx,desctrx,dbkrtrx,currtrx,kurs,cekbg,amounttrx From gl_transaksi "
        SQL = SQL + "Where tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
        SQL = SQL + "and kdtrx in('KK','KM') "
        SQL = SQL + "and noactrx = '" & txtacc1 & "' and reconsil in('0','is NULL')"
        SQL = SQL + "Order By notrx desc"
    Else
        SQL = "Select tgltrx,kdtrx,notrx,noactrx,desctrx,dbkrtrx,currtrx,kurs,cekbg,amounttrx From gl_transaksi "
        'SQL = SQL + "Where dbkrtrx = 'K' and tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
        SQL = SQL + "Where tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
        SQL = SQL + "and kdtrx in('BM','MB','BP','PH','BK','DA','TB','JU') "
        SQL = SQL + "and noactrx = '" & txtacc1 & "' and reconsil in('0','is NULL')" 'eitdb_old
        'SQL = SQL + "and noactrx = '" & txtacc1 & "' and reconsil is null or reconsil = '0' "
        SQL = SQL + "Order By notrx desc"
    End If
    
    Adodc1.RecordSource = SQL
    Set grid.DataSource = Adodc1
    Adodc1.Refresh
    Adodc1.Recordset.Requery -1
    Pg.Visible = True
    setdata
    grid.Refresh
End Sub

Private Sub setgrid()
    With grid
        .Cols = 11
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "Tanggal"
        .TextMatrix(0, 2) = "Kode"
        .TextMatrix(0, 3) = "No.Transaksi"
        .TextMatrix(0, 4) = "Account"
        .TextMatrix(0, 5) = "Description"
        
        .TextMatrix(0, 6) = "Currency" '8
        .TextMatrix(0, 7) = "Kurs" '9
        .TextMatrix(0, 8) = "Cek/Giro" '10
        .TextMatrix(0, 9) = "Debet"
        .TextMatrix(0, 10) = "Kredit"
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignRightCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 750
        .ColWidth(3) = 1500
        .ColWidth(4) = 1200
        .ColWidth(5) = 5550 '5000
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1200
        .ColWidth(9) = 2000
        .ColWidth(10) = 2000
    End With
End Sub
Private Sub setdata()
On Error GoTo Err_handler:
Dim jml As String
    setgrid
    jml = Adodc1.Recordset.RecordCount
    If jml = "0" Then Pg.Visible = False: Exit Sub
    Pg.Min = 0
    Pg.Max = jml
    Pg.Value = 0
    Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
    With Adodc1.Recordset
        .MoveFirst
        Do While Not .EOF
            With grid
            grid.Col = 0
            Set grid.CellPicture = uncheck.Picture
            .TextMatrix(.Row, 0) = grid.Row
            .TextMatrix(.Row, 1) = Format(Adodc1.Recordset!tgltrx, "yyyy/MM/dd")
            .TextMatrix(.Row, 2) = Adodc1.Recordset!kdtrx
            .TextMatrix(.Row, 3) = Adodc1.Recordset!notrx
            .TextMatrix(.Row, 4) = Adodc1.Recordset!noactrx
            .TextMatrix(.Row, 5) = Adodc1.Recordset!desctrx
            .TextMatrix(.Row, 6) = Adodc1.Recordset!currtrx
            .TextMatrix(.Row, 7) = Adodc1.Recordset!kurs
            .TextMatrix(.Row, 8) = Adodc1.Recordset!cekbg
            If Adodc1.Recordset!dbkrtrx = "D" Then
                .TextMatrix(.Row, 9) = Format(Adodc1.Recordset!amounttrx, "#,##0.00")
                .TextMatrix(.Row, 10) = ""
            Else
                .TextMatrix(.Row, 9) = ""
                .TextMatrix(.Row, 10) = Format(Adodc1.Recordset!amounttrx, "#,##0.00")
            End If
            
            Pg.Value = Pg.Value + 1
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            lblrecord = grid.Row & " Record"
            If grid.Row = jml Then Exit Do
            .Row = .Row + 1
            End With
            DoEvents
            Adodc1.Recordset.MoveNext
        Loop
    End With
    Pg.Value = 0
    Pg.Visible = False
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub
Private Sub hapusgrid()
On Error Resume Next
    Dim jml As Integer
    Dim j As Integer

    jml = grid.Rows
    j = 0
    grid.Row = 1
    Do While True
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.TextMatrix(grid.Row, 10) = ""
        If grid.Row = jml - 1 Then Exit Do
        grid.Row = grid.Row + 1
        lblrecord = lblrecord & " Record"
        DoEvents
    Loop
    grid.Rows = 2
    For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
    Next
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 0) = ""
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    grid.TextMatrix(grid.Row, 10) = ""
    Do While True
        grid.TextMatrix(grid.Row, 0) = grid.TextMatrix(grid.Row + 1, 0)
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.TextMatrix(grid.Row, 10) = grid.TextMatrix(grid.Row + 1, 10)
        If grid.Row = grid.Rows - 2 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub
Function tglsekarang()
    tglsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub Form_Resize()
    Pg.Move (Me.Width - Pg.Width) / 2, (Me.Height - Pg.Height) / 2
    Me.Top = 2250
    Me.Height = Screen.Height - 3150
    grid.Height = Me.Height - grid.Top - 900
    cmdclose.Top = Me.Height - 800
    cmdclear.Top = Me.Height - 800
    cmdsave.Top = Me.Height - 800
    lblrecord.Top = Me.Height - 800
End Sub

Private Sub grid_Click()
    Dim j As Integer
    j = 0
    
    If grid.MouseRow = 0 Then
        If grid.MouseCol = 1 Then
            If grid.TextMatrix(0, 1) = "Tanggal" Or grid.TextMatrix(0, 1) = "< Tanggal" Then
                grid.TextMatrix(0, 1) = "> Tanggal"
                grid.Sort = flexSortStringDescending
                sortnumber
            Else
                grid.TextMatrix(0, 1) = "< Tanggal"
                grid.Sort = flexSortStringAscending
                sortnumber
            End If
            grid.TextMatrix(0, 3) = "No.Transaksi"
            Exit Sub
        ElseIf grid.MouseCol = 3 Then
            If grid.TextMatrix(0, 3) = "No.Transaksi" Or grid.TextMatrix(0, 3) = "< No.Transaksi" Then
                grid.TextMatrix(0, 3) = "> No.Transaksi"
                grid.Sort = flexSortStringDescending
                sortnumber
            Else
                grid.TextMatrix(0, 3) = "< No.Transaksi"
                grid.Sort = flexSortStringAscending
                sortnumber
            End If
            grid.TextMatrix(0, 1) = "Tanggal"
            Exit Sub
        Else
            grid.TextMatrix(0, 1) = "Tanggal"
            grid.TextMatrix(0, 3) = "No.Transaksi"
            Exit Sub
        End If
    Else
        grid.TextMatrix(0, 1) = "Tanggal"
        grid.TextMatrix(0, 3) = "No.Transaksi"
    End If
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            grid.Col = 0
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HE0E0E0
                Next
                If grid.TextMatrix(grid.Row, 9) = "" Then
                    txtjumlah.Value = CDbl(grid.TextMatrix(grid.Row, 10) + CLng(txtjumlah.Value))
                Else
                    'txtjumlah.Value = CDbl(grid.TextMatrix(grid.Row, 9) + CLng(txtjumlah.Value))
                    txtjumlah.Value = CDbl(CLng(txtjumlah.Value) - grid.TextMatrix(grid.Row, 9))
                End If
                txtselisih.Value = txtnilai.Value - txtjumlah.Value

            ElseIf grid.CellPicture = check Then
                Set grid.CellPicture = uncheck
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HFFFFFF
                Next
                If grid.TextMatrix(grid.Row, 9) = "" Then
                    txtjumlah.Value = CDbl(CLng(txtjumlah.Value) - grid.TextMatrix(grid.Row, 10))
                Else
                    'txtjumlah.Value = CDbl(CLng(txtjumlah.Value) - grid.TextMatrix(grid.Row, 6))
                    txtjumlah.Value = CDbl(grid.TextMatrix(grid.Row, 9) + CLng(txtjumlah.Value))
                End If
                txtselisih.Value = txtnilai.Value - txtjumlah.Value
            End If
End Sub

Private Sub sortnumber()
    With Adodc1.Recordset
        .MoveFirst
        Do While Not .EOF
            grid.Col = 0
            grid.TextMatrix(grid.Row, 0) = grid.Row
            If grid.Row = .RecordCount Then Exit Sub
            grid.Row = grid.Row + 1
            .MoveNext
        Loop
    End With
End Sub

Private Sub txtjumlah_Change()
    If txtnilai.Value = txtjumlah.Value Then
        lblstatus = "Status : Balance"
        lblstatus.BackColor = &H80000011
    Else
        lblstatus = "Status : UnBalance"
        lblstatus.BackColor = vbRed
    End If
End Sub

Private Sub txtnilai_Change()
    If txtnilai.Value <> "0.00" Then
        txtselisih.Value = txtnilai.Value - txtjumlah.Value
    End If
    txtjumlah_Change
End Sub

Function getreconsil() As String    '2016060001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yyyymm")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(noreconsil)as kr from am_reconsil where noreconsil like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getreconsil = strnumber + "0001"
    Else
        kode = CLng(Mid(RST!kr, 7, 4)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "000" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 4) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getreconsil = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getreconsil = strnumber + "0001"
End Function
