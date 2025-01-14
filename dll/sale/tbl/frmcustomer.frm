VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmcustomer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Customer"
   ClientHeight    =   7035
   ClientLeft      =   5625
   ClientTop       =   4875
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleMode       =   0  'User
   ScaleWidth      =   5938.029
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Non Aktif"
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
      Left            =   4170
      TabIndex        =   39
      Top             =   5775
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ya"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Toko"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3270
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pabrik"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   3270
      Width           =   855
   End
   Begin VB.TextBox txtarea 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtNamaNPWP 
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
      Left            =   1680
      MaxLength       =   255
      TabIndex        =   13
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtAlamatNPWP 
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
      Height          =   525
      Left            =   1680
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   3975
   End
   Begin TDBNumber6Ctl.TDBNumber txterm 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmcustomer.frx":0000
      Caption         =   "frmcustomer.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcustomer.frx":0085
      Keys            =   "frmcustomer.frx":00A3
      Spin            =   "frmcustomer.frx":00ED
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtkota 
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
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtKodePos 
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
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtKontak 
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
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtFax 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtTelp 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtAlamat 
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
      Height          =   525
      Left            =   1680
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox txtNama 
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
      Left            =   1680
      MaxLength       =   255
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtKode 
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   6510
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
      MICON           =   "frmcustomer.frx":0115
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
      Left            =   3840
      TabIndex        =   17
      Top             =   6510
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
      MICON           =   "frmcustomer.frx":042F
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
      Left            =   2880
      TabIndex        =   16
      Top             =   6510
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
      MICON           =   "frmcustomer.frx":0749
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBMask6Ctl.TDBMask txtnonpwp 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   5400
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   503
      Caption         =   "frmcustomer.frx":0A63
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmcustomer.frx":0AC8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "9  9  9  9    9  9  9  9    9  9  9  9    9  9  9  9"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   0
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "                                                    "
      Value           =   ""
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   32
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Area"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmcustomer.frx":0B0A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox checkstatus 
      Height          =   270
      Left            =   3750
      TabIndex        =   35
      Top             =   135
      Width           =   210
      _Version        =   851970
      _ExtentX        =   370
      _ExtentY        =   476
      _StockProps     =   79
      BackColor       =   -2147483644
      UseVisualStyle  =   -1  'True
   End
   Begin TDBNumber6Ctl.TDBNumber txtlimit 
      Height          =   285
      Left            =   1680
      TabIndex        =   37
      Top             =   5760
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   503
      Calculator      =   "frmcustomer.frx":0E24
      Caption         =   "frmcustomer.frx":0E44
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcustomer.frx":0EB0
      Keys            =   "frmcustomer.frx":0ECE
      Spin            =   "frmcustomer.frx":0F10
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
      ValueVT         =   1945108485
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -15
      TabIndex        =   41
      Top             =   6405
      Width           =   5895
   End
   Begin VB.Label lbllimit 
      Caption         =   "Limit piutang non aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   1665
      TabIndex        =   40
      Top             =   6120
      Width           =   1665
   End
   Begin VB.Label Label15 
      Caption         =   "Limit"
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
      Left            =   225
      TabIndex        =   38
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NON AKTIF "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3990
      TabIndex        =   36
      Top             =   105
      Width           =   1680
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Penambahan Customer hanya dilakukan di kantor Pusat."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   34
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label lblarea 
      BackColor       =   &H80000005&
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
      TabIndex        =   33
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label11 
      Caption         =   "No Npwp"
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
      TabIndex        =   31
      Top             =   5430
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Nama Npwp"
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
      TabIndex        =   30
      Top             =   4470
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Alamat Npwp"
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
      TabIndex        =   29
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Kawasan berikat"
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
      TabIndex        =   28
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "T e r m                                       Hari"
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
      TabIndex        =   27
      Top             =   3630
      Width           =   2655
   End
   Begin VB.Label Label14 
      Caption         =   "K o t a"
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
      TabIndex        =   26
      Top             =   1470
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Pabrik / Toko"
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
      TabIndex        =   25
      Top             =   3270
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Contact Person"
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
      TabIndex        =   24
      Top             =   2910
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Faxsimile"
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
      TabIndex        =   23
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Telepone"
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
      TabIndex        =   22
      Top             =   2190
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat Customer"
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
      Top             =   870
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Customer"
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
      TabIndex        =   20
      Top             =   510
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Customer"
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
      TabIndex        =   19
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "frmcustomer"
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

Dim str1, str16, str99 As String
Dim strcust, strtype, strid As String
Dim int2 As Integer

Private Sub Check2_Click()
    If Check2.Value = Checked Then Check2.Caption = "Aktif": lbllimit = "Limit piutang aktif."
    If Check2.Value = Unchecked Then Check2.Caption = "Non Aktif": lbllimit = "Limit piutang non aktif."
End Sub

Private Sub checkstatus_Click()
    If checkstatus.Value = xtpChecked Then
        lblstatus.Caption = " Aktif ": lblstatus.ForeColor = &HC000&
    Else
        lblstatus.Caption = " Non Aktif ": lblstatus.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdadd_Click()
    If MsgBox("Penambahan Customer baru hanya dilakukan di Kantor Pusat." & vbCrLf & "Lanjukan ?", vbYesNo + vbExclamation, "Add Customer") = vbNo Then Exit Sub
    If Len(Trim(txtKode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtKode.SetFocus
        Exit Sub
    End If
    
    If txtKode = "" Or txtNama = "" Or txtAlamat = "" Or txtkota = "" Or txtarea = "" Or txtKodePos = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtKode = Trim(txtKode)
    
    int2 = 0
    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        history
        txtKode = str16
        int2 = 1
        
        GoTo jump99
        Exit Sub
    End If
    OBJ.Close
    
jump99:
    
    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        history
        txtKode = str16
        int2 = 1
        
        GoTo jump98
        Exit Sub
    End If
    OBJ.Close
    
jump98:
    
    OBJ.Open dsn
    SQL = "INSERT INTO AM_Customer"
    SQL = SQL + "(KodeCust"
    SQL = SQL + ",NamaCust"
    SQL = SQL + ",AlamatCust"
    SQL = SQL + ",kota"
    SQL = SQL + ",TelpCust"
    SQL = SQL + ",FaxCust"
    SQL = SQL + ",contactPerson"
    SQL = SQL + ",KodePos"
    SQL = SQL + ",termcust"
    SQL = SQL + ",limit"
    SQL = SQL + ",Namanpwp"
    SQL = SQL + ",alamatnpwp"
    SQL = SQL + ",Nonpwp"
    SQL = SQL + ",kodearea"
    SQL = SQL + ",kodeacgl"
    SQL = SQL + ",status"
    SQL = SQL + ",berikat"
    SQL = SQL + ",flaglimit"
    SQL = SQL + ",identry"
    SQL = SQL + ",dateupdate"
    SQL = SQL + ",idupdate"
    SQL = SQL + ",Dateentry)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + "('" & txtKode & "'"
    SQL = SQL + ", '" & txtNama & "'"
    SQL = SQL + ", '" & txtAlamat & "'"
    SQL = SQL + ", '" & txtkota & "'"
    SQL = SQL + ", '" & txtTelp & "'"
    SQL = SQL + ", '" & txtFax & "'"
    SQL = SQL + ", '" & txtKontak & "'"
    SQL = SQL + ", '" & txtKodePos & "'"
    SQL = SQL + ",convert(money, '" & txterm & "')"
    SQL = SQL + ",convert(money, '" & txtlimit & "')"
    SQL = SQL + ", '" & txtNamaNPWP & "'"
    SQL = SQL + ", '" & txtAlamatNPWP & "'"
    SQL = SQL + ", '" & txtnonpwp & "'"
    SQL = SQL + ", '" & txtarea & "'"
    SQL = SQL + ", '0'"
    If checkstatus.Value = xtpChecked Then SQL = SQL + ", '0'"
    If checkstatus.Value = xtpUnchecked Then SQL = SQL + ", '1'"
    If Check1.Value = 0 Then SQL = SQL + ", '0'"
    If Check1.Value = 1 Then SQL = SQL + ", '1'"
    If Check2.Value = 0 Then SQL = SQL + ", '0'"
    If Check2.Value = 1 Then SQL = SQL + ", '1'"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ", Convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", ''"
    SQL = SQL + ", Convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select c_type,c_id,ac_cust from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strcust = RST!ac_cust
        strtype = RST!c_type
        strid = RST!c_id
        
        OBJ1.Open dsn1
        SQL1 = "select top 1 noac from gl_masterac where noac like '" & strcust & "%' order by noac desc"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then strcust = RST1!noac
        OBJ1.Close
        
        strcust = strcust + 1
        
        OBJ1.Open dsn
        SQL1 = "select noacc from am_autoaccust where kodecomp = '" & strid & "' and kodecust ='" & txtKode & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "insert into am_autoaccust ("
            SQL1 = SQL1 + "kodecomp, "
            SQL1 = SQL1 + "noacc, "
            SQL1 = SQL1 + "kodecust)"
        
            SQL1 = SQL1 + " values('" & strid & "',"
            SQL1 = SQL1 + "'" & strcust & "',"
            SQL1 = SQL1 + "'" & txtKode & "')"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn1
            SQL1 = "insert into gl_masterac"
            SQL1 = SQL1 + "(noac"
            SQL1 = SQL1 + ",nmac"
            SQL1 = SQL1 + ",typeac"
            SQL1 = SQL1 + ",jenisac1"
            SQL1 = SQL1 + ",jenisac2"
            SQL1 = SQL1 + ",jenisac3"
            SQL1 = SQL1 + ",jenisac4"
            SQL1 = SQL1 + ",jenisac5"
            SQL1 = SQL1 + ",jenisac6"
            SQL1 = SQL1 + ",jenisac7"
            SQL1 = SQL1 + ",jenisac8"
            SQL1 = SQL1 + ",jenisac9"
            SQL1 = SQL1 + ",jenisac10"
            SQL1 = SQL1 + ",flag"
            SQL1 = SQL1 + ",idupdate"
            SQL1 = SQL1 + ",dateupdate"
            SQL1 = SQL1 + ",identry"
            SQL1 = SQL1 + ",Dateentry)"
            
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + "('" & strcust & "'"
            SQL1 = SQL1 + ", '" & Mid(txtNama + " (" + txtKode + ")", 1, 40) & "'"
            SQL1 = SQL1 + ", 'AS'"
            SQL1 = SQL1 + ", '" & strtype & "'"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", '" & txtKode & "'"
            SQL1 = SQL1 + ", '0'"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "')"
            SQL1 = SQL1 + ", '" & kuser & "'"
            SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
            
            SQL1 = "insert into gl_chacct"
            SQL1 = SQL1 + "(kdcomp"
            SQL1 = SQL1 + ",noac"
            SQL1 = SQL1 + ",typeac"
            SQL1 = SQL1 + ",balancedb"
            SQL1 = SQL1 + ",balancecr"
            SQL1 = SQL1 + ",begindb"
            SQL1 = SQL1 + ",begincr"
            SQL1 = SQL1 + ",periode01"
            SQL1 = SQL1 + ",periode02"
            SQL1 = SQL1 + ",periode03"
            SQL1 = SQL1 + ",periode04"
            SQL1 = SQL1 + ",periode05"
            SQL1 = SQL1 + ",periode06"
            SQL1 = SQL1 + ",periode07"
            SQL1 = SQL1 + ",periode08"
            SQL1 = SQL1 + ",periode09"
            SQL1 = SQL1 + ",periode10"
            SQL1 = SQL1 + ",periode11"
            SQL1 = SQL1 + ",periode12"
            SQL1 = SQL1 + ",periode13"
            SQL1 = SQL1 + ",last01"
            SQL1 = SQL1 + ",last02"
            SQL1 = SQL1 + ",last03"
            SQL1 = SQL1 + ",last04"
            SQL1 = SQL1 + ",last05"
            SQL1 = SQL1 + ",last06"
            SQL1 = SQL1 + ",last07"
            SQL1 = SQL1 + ",last08"
            SQL1 = SQL1 + ",last09"
            SQL1 = SQL1 + ",last10"
            SQL1 = SQL1 + ",last11"
            SQL1 = SQL1 + ",last12"
            SQL1 = SQL1 + ",last13"
            SQL1 = SQL1 + ",temp01"
            SQL1 = SQL1 + ",temp02"
            SQL1 = SQL1 + ",temp03"
            SQL1 = SQL1 + ",temp04"
            SQL1 = SQL1 + ",temp05"
            SQL1 = SQL1 + ",temp06"
            SQL1 = SQL1 + ",temp07"
            SQL1 = SQL1 + ",temp08"
            SQL1 = SQL1 + ",temp09"
            SQL1 = SQL1 + ",temp10"
            SQL1 = SQL1 + ",temp11"
            SQL1 = SQL1 + ",temp12"
            SQL1 = SQL1 + ",temp13"
            SQL1 = SQL1 + ",budget01"
            SQL1 = SQL1 + ",budget02"
            SQL1 = SQL1 + ",budget03"
            SQL1 = SQL1 + ",budget04"
            SQL1 = SQL1 + ",budget05"
            SQL1 = SQL1 + ",budget06"
            SQL1 = SQL1 + ",budget07"
            SQL1 = SQL1 + ",budget08"
            SQL1 = SQL1 + ",budget09"
            SQL1 = SQL1 + ",budget10"
            SQL1 = SQL1 + ",budget11"
            SQL1 = SQL1 + ",budget12"
            SQL1 = SQL1 + ",budget13)"
            
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + "('" & strid & "'"
            SQL1 = SQL1 + ", '" & strcust & "'"
            SQL1 = SQL1 + ", 'AS'"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0')"
            SQL1 = SQL1 + ", convert(money,'0'))"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
    End If
    OBJ.Close
        
    If int2 = 1 Then
        MsgBox "Data already exist, data was saved with next number " & txtKode & vbCrLf & _
        "Click OK To Continue ...", vbExclamation, "Warning"
    Else
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtKode = ""
    txtNama = ""
    txtAlamat = ""
    txtkota = ""
    txtTelp = ""
    txtFax = ""
    txtKontak = ""
    txtKodePos = ""
    txtNamaNPWP = ""
    txtAlamatNPWP = ""
    txtnonpwp = ""
    txtarea = ""
    lblarea = ""
    txterm = 0
    txtlimit = "0.00"
    Check1.Value = 0
    Check2.Value = 0
    Option1.Value = False
    Option2.Value = False
    txtKode.SetFocus
    checkstatus.Value = xtpUnchecked: lblstatus = " Non Aktif ": lblstatus.ForeColor = &HFF&
    history
    txtKode = str16
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdsearch_Click()
    carisql1 = "select kode, nama from am_area"
    namatabel = "Area"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea = hasil
    cariarea
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtTelp.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    history
    txtKode = str16
End Sub

Private Sub history()
    OBJ1.Open dsn
    SQL1 = "select id2,kode2 from am_branch"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        If RST1!id2 = "1" Then str1 = RST1!kode2 Else str1 = "0"
    Else
        str1 = "0"
    End If
    
    SQL1 = "select top 1 kodecust from am_customer where kodecust like 'C-" & str1 & "%' order by kodecust desc"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        str99 = RST1!kodecust
    Else
        str99 = "C-" & str1 & "0000"
    End If
    OBJ1.Close
    
    str99 = Right(str99, 4)
    str99 = str99 + 1
    If Len(str99) = 1 Then str16 = "C-" & str1 & "000" & str99
    If Len(str99) = 2 Then str16 = "C-" & str1 & "00" & str99
    If Len(str99) = 3 Then str16 = "C-" & str1 & "0" & str99
    If Len(str99) = 4 Then str16 = "C-" & str1 & str99
    If Len(str99) = 5 Then str16 = ""
End Sub

Private Sub Option1_Click()
    txtKodePos = "1"
End Sub

Private Sub Option2_Click()
    txtKodePos = "2"
End Sub

Private Sub txtarea_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtTelp.SetFocus
End Sub

Private Sub txtarea_LostFocus()
     cariarea
End Sub

Private Sub cariarea()
    If txtarea = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_area where kode = '" & txtarea & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblarea = RST!nama
        
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Information"
    txtarea = ""
    lblarea = ""
    txtarea.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKontak.SetFocus
End Sub

Private Sub txtKode_GotFocus()
    Call Blok(txtKode)
End Sub

Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNama.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtkode_LostFocus()
    If txtKode = "" Then Exit Sub
    If txtKode.SelLength <> 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtNama = RST!namacust
        txtAlamat = RST!alamatcust
        txtkota = RST!kota
        txtTelp = RST!telpcust
        txtFax = RST!faxcust
        txtKontak = RST!contactperson
        txterm = RST!termcust
        txtNamaNPWP = RST!namanpwp
        txtAlamatNPWP = RST!alamatnpwp
        txtnonpwp = RST!nonpwp
        txtarea = RST!kodearea
        If RST!limit = 0 Then Check1.Value = 0
        If RST!limit = 1 Then Check1.Value = 1
        
        SQL = "select * from am_area where kode = '" & txtarea & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblarea = RST!nama
        
        MsgBox "Data already exist.", vbInformation, "Information"
        txtKode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    txtNama = ""
    txtAlamat = ""
    txtkota = ""
    txtTelp = ""
    txtFax = ""
    txtKontak = ""
    txtKodePos = ""
    Option1.Value = False
    Option2.Value = False
    txtNamaNPWP = ""
    txtAlamatNPWP = ""
    txtnonpwp = ""
    txtarea = ""
    lblarea = ""
    txterm = 0
    Check1.Value = 0
    txtNama.SetFocus
End Sub

Private Sub txtkota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtarea.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamat.SetFocus
End Sub

Private Sub txtNamaNPWP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatNPWP.SetFocus
End Sub

Private Sub txtTelp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFax.SetFocus
End Sub
