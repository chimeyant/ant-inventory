VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmbukubesar 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7095
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
   Icon            =   "frmbukubesar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "with Report Code"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtcode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Rekapitulasi Buku Besar"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Buku Besar"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Ganti Halaman Per Account"
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
      TabIndex        =   14
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Total Per Tanggal Transaksi"
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
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Total Per Account"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Currency, Rates dan Base Ditampilkan"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtacc2 
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
      Left            =   4920
      MaxLength       =   15
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   3120
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
      Format          =   134414339
      CurrentDate     =   37728
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
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
      Format          =   134414339
      CurrentDate     =   37728
   End
   Begin VB.TextBox txtarea1 
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
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtarea2 
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
      Left            =   4920
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   1200
      Top             =   6000
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
   Begin TDBNumber6Ctl.TDBNumber txtpanjang 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmbukubesar.frx":2372
      Caption         =   "frmbukubesar.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmbukubesar.frx":23E7
      Keys            =   "frmbukubesar.frx":2405
      Spin            =   "frmbukubesar.frx":2447
      AlignHorizontal =   2
      AlignVertical   =   2
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
      MaxValue        =   10
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   10
      MaxValueVT      =   1937178629
      MinValueVT      =   1397948421
   End
   Begin TDBNumber6Ctl.TDBNumber txtperiode 
      Height          =   285
      Left            =   4920
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmbukubesar.frx":246F
      Caption         =   "frmbukubesar.frx":248F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmbukubesar.frx":24E4
      Keys            =   "frmbukubesar.frx":2502
      Spin            =   "frmbukubesar.frx":2544
      AlignHorizontal =   2
      AlignVertical   =   2
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
      MaxValue        =   13
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1638405
      Value           =   1
      MaxValueVT      =   1937178629
      MinValueVT      =   1397948421
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4200
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
      MICON           =   "frmbukubesar.frx":256C
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
      Left            =   5880
      TabIndex        =   17
      Top             =   4200
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
      MICON           =   "frmbukubesar.frx":2886
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
      Left            =   360
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Account"
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
      MICON           =   "frmbukubesar.frx":2BA0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch4 
      Height          =   285
      Left            =   3600
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Account"
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
      MICON           =   "frmbukubesar.frx":2EBA
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
      Left            =   360
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Company"
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
      MICON           =   "frmbukubesar.frx":31D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   3600
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Company"
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
      MICON           =   "frmbukubesar.frx":34EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Report Code"
      ENAB            =   0   'False
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
      MICON           =   "frmbukubesar.frx":3808
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Buku"
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
      TabIndex        =   22
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Besar"
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
      TabIndex        =   21
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label9 
      Caption         =   "Panjang Acc.                          (max 10)"
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
      Left            =   360
      TabIndex        =   20
      Top             =   1710
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
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
      Left            =   3840
      TabIndex        =   19
      Top             =   3150
      Width           =   855
   End
   Begin VB.Label Label6 
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
      TabIndex        =   18
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmbukubesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2, str3, str5, str6, str7, str8 As String
Dim PL, PJ As Boolean

Function nmcomp()
    nmcomp = " Konsolidasi "
    If txtarea1 = txtarea2 Then
        OBJ.Open dsn
        SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then nmcomp = RST!nmcompprn
        OBJ.Close
    End If
End Function

Function period()
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    period = "Per " & Format(RST!tglakhir, "dd MMMM yyyy")
    OBJ.Close
End Function

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        date1 = Date
        date2 = Date
        txtperiode = Month(date1)
        txtarea1.SetFocus
    Else
        date1 = RST!tglawal
        date2 = RST!tglakhir
        txtperiode = RST!periode
        
        format_coa = RST!formatac
    End If
    OBJ.Close
End Sub

Private Sub cariarea2()
    If txtarea2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea2 & " Not Found.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select Form_no, description from gl_rforms"
    namatabel = "Buku Besar"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtcode = hasil
    hasil = ""
    hasil1 = ""
    txtarea1.SetFocus
End Sub

Private Sub cmdview_Click()
    If txtarea1 = "" Or txtarea2 = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtarea2 < txtarea1 Then
        MsgBox "To Company Can Not Smaller Then From Company.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
        Exit Sub
    End If
    
    If Option3.Value And txtcode = "" Then
        MsgBox "Data entry not complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtacc1 = "" Then txtacc1 = "0"
    If txtacc2 = "" Then txtacc2 = "z"
    If txtacc1 <> "" Then txtacc1 = x_original(txtacc1)
    If txtacc2 <> "" Then txtacc2 = x_original(txtacc2)
    
    If txtacc2 < txtacc1 Then
        MsgBox "To Account Can Not Smaller Then From Account.", vbExclamation, "Warning"
        txtacc2 = ""
        txtacc2.SetFocus
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    str7 = RST!periode
    OBJ.Close
    
    str3 = Str(Check1.Value)
    str5 = Str(Check3.Value)
    str8 = Str(Check4.Value)
    str6 = Str(txtperiode)
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    
    If Not Option3.Value Then
        If Check5.Value = 0 Then
            If txtarea1 = txtarea2 And txtpanjang <> 10 Then
                str2 = Str(txtpanjang)
                crystal.DataFiles(0) = "Proc(gl_ledger1)"
                If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger1.rpt"
                If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger01.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(2) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(3) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(4) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(7) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(8) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(10) = "@panjang;" + str2 + ";true"
                crystal.ParameterFields(11) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 = txtarea2 And txtpanjang = 10 Then
                If PL = True Then
                    PL = False
                    'MsgBox "L"
                    crystal.DataFiles(0) = "Proc(gl_ledger2b)" '*Rekap Buku besar biaya "L"
                    If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger2b.rpt" '*Buku besar biaya "L"
                    If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger02b.rpt" '*Rekap Buku besar biaya
                ElseIf PJ = True Then
                    PJ = False
                    'MsgBox "P"
                    crystal.DataFiles(0) = "Proc(gl_ledger2p)" '"P" (tester)
                    If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger2a.rpt"
                    If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger02a.rpt"
                Else
                    crystal.DataFiles(0) = "Proc(gl_ledger2)" '"P & L"
                    crystal.DataFiles(0) = "Proc(gl_ledger2_devtes)" '"P & L" (tester)
        '-----------MsgBox txtarea1 & ";" & txtacc1 & ";" & txtacc2 & ";" & str3 & ";" & str5 & ";" & str6 & ";" & str7 & ";" & nmuser & ";" & str2 & ";" & str8 & ";" & Format(date1, "yyyyMMdd")
                    If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger2.rpt"
                    If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger02.rpt"
                End If
                
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(2) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(3) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(4) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(7) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(8) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(10) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 <> txtarea2 And txtpanjang <> 10 Then
                str2 = Str(txtpanjang)
                crystal.DataFiles(0) = "Proc(gl_ledger3)"
                If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger3.rpt"
                If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger03.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
                crystal.ParameterFields(2) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(3) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(4) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(6) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(7) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(8) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(9) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(11) = "@panjang;" + str2 + ";true"
                crystal.ParameterFields(12) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 <> txtarea2 And txtpanjang = 10 Then
                crystal.DataFiles(0) = "Proc(gl_ledger4)"
                If Option1.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger4.rpt"
                If Option2.Value Then crystal.ReportFileName = AppPath & "\reports\gl\report\ledger04.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
                crystal.ParameterFields(2) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(3) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(4) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(6) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(7) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(8) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(9) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(11) = "@pilih4;" + str8 + ";true"
            End If
        Else
            If txtarea1 = txtarea2 And txtpanjang <> 10 Then
                str2 = Str(txtpanjang)
                crystal.DataFiles(0) = "Proc(gl_ledger5)"
                crystal.ReportFileName = App.Path & "\report\ledger5.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(2) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(3) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(4) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(7) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(8) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(10) = "@panjang;" + str2 + ";true"
                crystal.ParameterFields(11) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 = txtarea2 And txtpanjang = 10 Then
                crystal.DataFiles(0) = "Proc(gl_ledger6)"
                crystal.ReportFileName = App.Path & "\report\ledger6.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(2) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(3) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(4) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(7) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(8) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(10) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 <> txtarea2 And txtpanjang <> 10 Then
                str2 = Str(txtpanjang)
                crystal.DataFiles(0) = "Proc(gl_ledger7)"
                crystal.ReportFileName = App.Path & "\report\ledger7.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
                crystal.ParameterFields(2) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(3) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(4) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(6) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(7) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(8) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(9) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(11) = "@panjang;" + str2 + ";true"
                crystal.ParameterFields(12) = "@pilih4;" + str8 + ";true"
            ElseIf txtarea1 <> txtarea2 And txtpanjang = 10 Then
                crystal.DataFiles(0) = "Proc(gl_ledger8)"
                crystal.ReportFileName = App.Path & "\report\ledger8.rpt"
                crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
                crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
                crystal.ParameterFields(2) = "@acc1;" + txtacc1 + ";true"
                crystal.ParameterFields(3) = "@acc2;" + txtacc2 + ";true"
                crystal.ParameterFields(4) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
                crystal.ParameterFields(5) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
                crystal.ParameterFields(6) = "@pilih1;" + str3 + ";true"
                crystal.ParameterFields(7) = "@pilih3;" + str5 + ";true"
                crystal.ParameterFields(8) = "@periode;" + str6 + ";true"
                crystal.ParameterFields(9) = "@periode1;" + str7 + ";true"
                crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
                crystal.ParameterFields(11) = "@pilih4;" + str8 + ";true"
            End If
        End If
    Else
        crystal.ReportFileName = App.Path & "\report\ledger9.rpt"
        crystal.DataFiles(0) = "Proc(gl_ledger9)"
        crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@acc1;" + txtacc1 + ";true"
        crystal.ParameterFields(2) = "@acc2;" + txtacc2 + ";true"
        crystal.ParameterFields(3) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
        crystal.ParameterFields(4) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
        crystal.ParameterFields(5) = "@pilih1;" + str3 + ";true"
        crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
        crystal.ParameterFields(7) = "@periode;" + str6 + ";true"
        crystal.ParameterFields(8) = "@periode1;" + str7 + ";true"
        crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(10) = "@pilih4;" + str8 + ";true"
        crystal.ParameterFields(11) = "@form;" + txtcode + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    If txtacc1 = "0" Then txtacc1 = ""
    If txtacc2 = "z" Then txtacc2 = ""
    If txtacc1 <> "" Then txtacc1 = original(txtacc1)
    If txtacc2 <> "" Then txtacc2 = original(txtacc2)
End Sub

Private Sub date1_Change()
    txtperiode = Month(date1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
    'L  (F2)
        PL = True
        cmdview_Click
    ElseIf KeyCode = 112 Then
    'P  (F1)
        PJ = True
        cmdview_Click
    End If
End Sub

Private Sub Form_Load()
   
    
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub Option1_Click()
    Check5.Enabled = True
    
    Check1.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check1.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    
    txtcode = ""
    txtcode.Enabled = False
    cmdsearch.Enabled = False
End Sub

Private Sub Option2_Click()
    Check5.Value = 0
    Check5.Enabled = False
    
    Check1.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check1.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    
    txtcode = ""
    txtcode.Enabled = False
    cmdsearch.Enabled = False
End Sub

Private Sub Option3_Click()
    Check1.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    Check5.Enabled = False
    Check1.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    
    txtcode = ""
    txtcode.Enabled = True
    cmdsearch.Enabled = True
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea2.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub txtarea2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc1.SetFocus
End Sub

Private Sub txtarea2_LostFocus()
    cariarea2
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    hasil = ""
    hasil1 = ""
    txtarea2.SetFocus
    cariarea1
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "'"
    If txtarea1 = txtarea2 Then
        namatabel = "Company Account "
    Else
        namatabel = "Company Account  "
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    hasil = ""
    hasil1 = ""
    txtacc2.SetFocus
End Sub

Private Sub cmdsearch4_Click()
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "'"
    If txtarea1 = txtarea2 Then
        namatabel = "Company Account "
    Else
        namatabel = "Company Account  "
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Private Sub cariacc1()
    If txtacc1 = "" Then Exit Sub
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtacc1) & "'"
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "' and b.noac = '" & x_original(txtacc1) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc1 & " Not Found.", vbExclamation, "Warning"
        txtacc1 = ""
        txtacc1.SetFocus
    Else
        If txtarea1 = txtarea2 Then txtacc1 = original(RST!noac)
    End If
    OBJ.Close
End Sub

Private Sub cariacc2()
    If txtacc2 = "" Then Exit Sub
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    OBJ.Open dsn
    'sql = "select * from gl_masterac where noac = '" & x_original(txtacc2) & "'"
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "' and b.noac = '" & x_original(txtacc2) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc2 & " Not Found.", vbExclamation, "Warning"
        txtacc2 = ""
        txtacc2.SetFocus
    Else
        If txtarea1 = txtarea2 Then txtacc2 = original(RST!noac)
    End If
    OBJ.Close
End Sub

Private Sub txtacc1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc2.SetFocus
End Sub

Private Sub txtacc1_LostFocus()
    cariacc1
End Sub

Private Sub txtacc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtacc2_LostFocus()
    cariacc2
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea1.SetFocus
End Sub

Private Sub txtcode_LostFocus()
    If txtcode = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_rforms where form_no = '" & txtcode & "' and report_type = '4'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Report Code " & txtcode & " Not Found.", vbExclamation, "Warning"
        txtcode = ""
        txtcode.SetFocus
    End If
    OBJ.Close
End Sub
