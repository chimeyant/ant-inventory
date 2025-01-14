VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmintranlist 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7830
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
   Icon            =   "frmintranlist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "UnBalance"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Unposted/Balance"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Posted/Balance"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   3000
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Adjustment"
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
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Non Adjustment"
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
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   600
      Top             =   3600
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
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
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
      Format          =   51052547
      CurrentDate     =   37694
   End
   Begin TDBText6Ctl.TDBText txtcom1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmintranlist.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmintranlist.frx":23DE
      Key             =   "frmintranlist.frx":23FC
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
   Begin TDBText6Ctl.TDBText txtkodetran1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmintranlist.frx":2438
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmintranlist.frx":24A4
      Key             =   "frmintranlist.frx":24C2
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
      MaxLength       =   2
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
   Begin TDBText6Ctl.TDBText txtnotran1 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Caption         =   "frmintranlist.frx":24FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmintranlist.frx":256A
      Key             =   "frmintranlist.frx":2588
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
      MaxLength       =   15
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   1560
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
      Format          =   50921475
      CurrentDate     =   37694
   End
   Begin TDBText6Ctl.TDBText txtkodetran2 
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   2280
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmintranlist.frx":25C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmintranlist.frx":2630
      Key             =   "frmintranlist.frx":264E
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
      MaxLength       =   2
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
   Begin TDBText6Ctl.TDBText txtnotran2 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Caption         =   "frmintranlist.frx":268A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmintranlist.frx":26F6
      Key             =   "frmintranlist.frx":2714
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
      MaxLength       =   15
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Company"
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
      MICON           =   "frmintranlist.frx":2750
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Preview"
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
      MICON           =   "frmintranlist.frx":2A6A
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
      Left            =   6600
      TabIndex        =   11
      Top             =   3360
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
      MICON           =   "frmintranlist.frx":2D84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox cmbdaerah 
      Height          =   285
      Left            =   2040
      TabIndex        =   26
      Top             =   1920
      Width           =   855
      VariousPropertyBits=   746608667
      MaxLength       =   1
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1508;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   3
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Kode Daerah"
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
      Left            =   360
      TabIndex        =   25
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Bank In"
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
      TabIndex        =   20
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "List"
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
      TabIndex        =   19
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label8 
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
      Left            =   4200
      TabIndex        =   18
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "To Kode Transaksi"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   2310
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "To No. Transaksi"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   2670
      Width           =   1335
   End
   Begin VB.Label Label11 
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
      Left            =   360
      TabIndex        =   15
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "From Kode Transaksi"
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
      Left            =   360
      TabIndex        =   14
      Top             =   2310
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "From No. Transaksi"
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
      Left            =   360
      TabIndex        =   13
      Top             =   2670
      Width           =   1695
   End
   Begin VB.Label lblcom1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmintranlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2, str3, str4, str5, str6, str7, str8 As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtcom1 = hasil
    caricom1
    hasil = ""
End Sub

Private Sub cmdview_Click()
    If txtcom1 = "" Then Exit Sub
    If (txtkodetran1 = "" And txtkodetran2 <> "") Or (txtkodetran2 = "" And txtkodetran1 <> "") Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    If (txtnotran1 = "" And txtnotran2 <> "") Or (txtnotran2 = "" And txtnotran1 <> "") Or cmbdaerah = "" Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Check1.Value = 0 And Check2.Value = 0 Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    If txtkodetran2 < txtkodetran1 Then
        MsgBox "To Kode Can Not Smaller Then From Kode.", vbExclamation, "Warning"
        txtkodetran2 = ""
        txtkodetran2.SetFocus
        Exit Sub
    End If
    
    If txtnotran2 < txtnotran1 Then
        MsgBox "To No. Can Not Smaller Then From No.", vbExclamation, "Warning"
        txtnotran2 = ""
        txtnotran2.SetFocus
        Exit Sub
    End If
    
    str1 = txtkodetran1
    str2 = txtkodetran2
    str3 = txtnotran1
    str4 = txtnotran2
    str7 = "I"
    
    If txtkodetran1 = "" And txtkodetran2 = "" Then
        str1 = "0"
        str2 = "z"
    End If
    
    If txtnotran1 = "" And txtnotran2 = "" Then
        str3 = "0"
        str4 = "z"
    End If
    
    If Check1.Value = 1 Then
        str5 = "adjoke"
    Else
        str5 = "adjxoke"
    End If
    
    If Check2.Value = 1 Then
        str6 = "nonadjoke"
    Else
        str6 = "nonadjxoke"
    End If
    
    If Option1.Value Then str8 = "P"
    If Option2.Value Then str8 = "B"
    If Option3.Value Then str8 = "U"
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(gl_tranlist)"
    crystal.ReportFileName = AppPath & "\reports\gl\ledger\tranlist.rpt"
    crystal.ParameterFields(0) = "@com1;" + txtcom1 + ";true"
    crystal.ParameterFields(1) = "@tgl1;" + Format(date1, "yyyyMMdd") + ";true"
    crystal.ParameterFields(2) = "@tgl2;" + Format(date2, "yyyyMMdd") + ";true"
    crystal.ParameterFields(3) = "@kode1;" + str1 + ";true"
    crystal.ParameterFields(4) = "@kode2;" + str2 + ";true"
    crystal.ParameterFields(5) = "@no1;" + str3 + ";true"
    crystal.ParameterFields(6) = "@no2;" + str4 + ";true"
    crystal.ParameterFields(7) = "@pilih1;" + str5 + ";true"
    crystal.ParameterFields(8) = "@pilih2;" + str6 + ";true"
    crystal.ParameterFields(9) = "@pilih3;" + str7 + ";true"
    crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
    crystal.ParameterFields(11) = "@post1;" + str8 + ";true"
    crystal.ParameterFields(12) = "@post2;;true"
    crystal.ParameterFields(13) = "@kode3;" + cmbdaerah + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    
    cmbdaerah.Clear
    cmbdaerah.ColumnCount = 2
    cmbdaerah.ListWidth = "4 cm"
    cmbdaerah.ColumnWidths = "1 cm; 3 cm"
    cmbdaerah.AddItem "1"
    cmbdaerah.AddItem "2"
    cmbdaerah.AddItem "3"
    cmbdaerah.AddItem "4"
    cmbdaerah.List(0, 1) = "Pabrik"
    cmbdaerah.List(1, 1) = "Jakarta"
    cmbdaerah.List(2, 1) = "Surabaya"
    cmbdaerah.List(3, 1) = "Semarang"
End Sub

Private Sub txtcom1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtcom1_LostFocus()
    caricom1
End Sub

Private Sub caricom1()
    If txtcom1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtcom1 & " Not Found.", vbExclamation, "Warning"
        txtcom1 = ""
        lblcom1 = ""
        date1 = Date
        date2 = Date
        txtcom1.SetFocus
    Else
        lblcom1 = RST!nmcompscr
        date1 = RST!tglawal
        date2 = RST!tglakhir
    End If
    OBJ.Close
End Sub

Private Sub txtkodetran1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodetran2.SetFocus
End Sub

Private Sub txtkodetran1_LostFocus()
    carikodetran1
End Sub

Private Sub carikodetran1()
    If txtkodetran1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdtrx = '" & txtkodetran1 & "' and kdcomp = '" & txtcom1 & "' and flagprint = 'I' and flagadjustment = '" & cmbdaerah & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Kode Transaction " & txtkodetran1 & " Not Found.", vbExclamation, "Warning"
        txtkodetran1 = ""
        txtkodetran1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkodetran2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnotran1.SetFocus
End Sub

Private Sub txtkodetran2_LostFocus()
    carikodetran2
End Sub

Private Sub carikodetran2()
    If txtkodetran2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdtrx = '" & txtkodetran2 & "' and kdcomp = '" & txtcom1 & "' and flagprint = 'I' and flagadjustment = '" & cmbdaerah & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Kode Transaction " & txtkodetran2 & " Not Found.", vbExclamation, "Warning"
        txtkodetran2 = ""
        txtkodetran2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtnotran1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnotran2.SetFocus
End Sub

Private Sub txtnotran2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub
