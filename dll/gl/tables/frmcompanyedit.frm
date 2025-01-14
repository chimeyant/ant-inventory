VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcompanyedit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Edit"
   ClientHeight    =   5850
   ClientLeft      =   5715
   ClientTop       =   5565
   ClientWidth     =   7680
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
   Icon            =   "frmcompanyedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   4800
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
      Format          =   178651139
      CurrentDate     =   37694
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   4800
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
      Format          =   178651139
      CurrentDate     =   37694
   End
   Begin VB.TextBox txtlayout 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   7
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtcontact 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox txtelp 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   5
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox txtkota 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtalamat 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   40
      TabIndex        =   3
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox txtscreen 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   40
      TabIndex        =   2
      Top             =   2280
      Width           =   5535
   End
   Begin VB.TextBox txtprint 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1920
      Width           =   5535
   End
   Begin TDBText6Ctl.TDBText txtkode 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   503
      Caption         =   "frmcompanyedit.frx":2372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcompanyedit.frx":23DE
      Key             =   "frmcompanyedit.frx":23FC
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
   Begin VB.TextBox txtype 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode"
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
      MICON           =   "frmcompanyedit.frx":2438
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   4440
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtperiode"
      BuddyDispid     =   196617
      OrigLeft        =   3480
      OrigTop         =   3480
      OrigRight       =   3720
      OrigBottom      =   3735
      Max             =   13
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmcompanyedit.frx":2752
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdelete 
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmcompanyedit.frx":2A6C
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
      Left            =   5520
      TabIndex        =   13
      Top             =   5280
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
      MICON           =   "frmcompanyedit.frx":2D86
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
      Left            =   6480
      TabIndex        =   14
      Top             =   5280
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
      MICON           =   "frmcompanyedit.frx":30A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdreplace 
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Automatically Add Company Account"
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
      MICON           =   "frmcompanyedit.frx":33BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label txtperiode 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   27
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
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
      TabIndex        =   30
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Updating"
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
      TabIndex        =   29
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblnama 
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
      Left            =   1800
      TabIndex        =   26
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      Caption         =   "Tanggal Periode                                                 S/D"
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
      Top             =   4830
      Width           =   4815
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      Caption         =   "Periode On Proses"
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
      TabIndex        =   24
      Top             =   4470
      Width           =   1575
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      Caption         =   "Layout Account"
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
      TabIndex        =   23
      Top             =   4110
      Width           =   1575
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3750
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Telepon"
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
      TabIndex        =   21
      Top             =   3390
      Width           =   1575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Kota"
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
      TabIndex        =   20
      Top             =   3030
      Width           =   1575
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Alamat "
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
      TabIndex        =   19
      Top             =   2670
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Nama On Screen"
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
      Top             =   2310
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Nama On Print"
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
      TabIndex        =   17
      Top             =   1950
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Company Type"
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
      TabIndex        =   16
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmcompanyedit"
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

Dim str1, str2, str11, str12 As String

Private Sub addchacct()
    OBJ1.Open dsn
    SQL1 = "select * from gl_chacct where kdcomp = '" & txtkode & "' and noac = '" & str1 & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close
        Exit Sub
    End If
    
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
    SQL1 = SQL1 + "('" & txtkode & "'"
    SQL1 = SQL1 + ", '" & str1 & "'"
    SQL1 = SQL1 + ", '" & str2 & "'"
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
    OBJ1.Close
End Sub

Private Sub cmdclear_Click()
    txtkode = ""
    txtype = ""
    txtprint = ""
    txtscreen = ""
    txtalamat = ""
    txtkota = ""
    txtelp = ""
    txtcontact = ""
    txtlayout = ""
    txtperiode = 1
    lblnama = ""
    date1 = Date
    date2 = Date
    UpDown1.Enabled = True
    date1.Enabled = True
    date2.Enabled = True
    txtkode.Enabled = True
    cmdsearch.Enabled = True
    txtkode.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub cmdelete_Click()
    If txtkode = "" Or txtype = "" Or txtprint = "" _
    Or txtscreen = "" Or txtlayout = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_chacct WHERE kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_transaksi WHERE kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "DELETE FROM gl_company WHERE Kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE FROM gl_chacct WHERE kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
    Exit Sub
    
jump:
    MsgBox "Can Not Delete, Record Still In Use.", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdreplace_Click()
    If txtkode = "" Or txtype = "" Or txtprint = "" _
    Or txtscreen = "" Or txtlayout = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Add Company Account ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_chacct WHERE kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jumpc
    
    SQL = "select * from gl_masterac where jenisac1 = '" & txtype & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str1 = RST!noac
        str2 = RST!typeac
        
        addchacct
        
        RST.MoveNext
    Loop
    SQL = "select * from gl_masterac where jenisac2 = '" & txtype & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str1 = RST!noac
        str2 = RST!typeac
        
        addchacct
        
        RST.MoveNext
    Loop
    SQL = "select * from gl_masterac where jenisac3 = '" & txtype & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str1 = RST!noac
        str2 = RST!typeac
        
        addchacct
        
        RST.MoveNext
    Loop
    SQL = "select * from gl_masterac where jenisac4 = '" & txtype & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str1 = RST!noac
        str2 = RST!typeac
        
        addchacct
        
        RST.MoveNext
    Loop
    SQL = "select * from gl_masterac where jenisac5 = '" & txtype & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        str1 = RST!noac
        str2 = RST!typeac
        
        addchacct
        
        RST.MoveNext
    Loop
    
    MsgBox "Company Account Added, Click OK To Continue ...", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
    Exit Sub
    
jumpc:
    MsgBox "Can Not Delete, There is already Company Account.", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    txtKode_LostFocus
    hasil = ""
End Sub

Private Sub cmdupdate_click()
    If txtkode = "" Or txtype = "" Or txtprint = "" _
    Or txtscreen = "" Or txtlayout = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub

    OBJ.Open dsn
    SQL = "UPDATE gl_company SET "
    SQL = SQL + "Nmcompprn = '" & txtprint & "',"
    SQL = SQL + "Nmcompscr = '" & txtscreen & "',"
    SQL = SQL + "alcomp = '" & txtalamat & "',"
    SQL = SQL + "kotacomp = '" & txtkota & "',"
    SQL = SQL + "tlpcomp = '" & txtelp & "',"
    SQL = SQL + "formatac = '" & txtlayout & "',"
    SQL = SQL + "contactperson = '" & txtcontact & "',"
    SQL = SQL + "tglawal = convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "tglakhir = convert(datetime,'" & tanggal2 & "'),"
    SQL = SQL + "periode = '" & txtperiode & "',"
    SQL = SQL + "idUpdate = '" & kuser & "',"
    SQL = SQL + "DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + "WHERE Kdcomp =  '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    txtperiode = 1
End Sub

Private Sub txtalamat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkota.SetFocus
End Sub

Private Sub txtcontact_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtlayout.SetFocus
End Sub

Private Sub txtelp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtcontact.SetFocus
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtprint.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    If txtkode = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtype = RST!kdtype
        txtprint = RST!nmcompprn
        txtscreen = RST!nmcompscr
        txtalamat = RST!alcomp
        txtkota = RST!kotacomp
        txtelp = RST!tlpcomp
        txtcontact = RST!contactperson
        txtlayout = RST!formatac
        txtperiode = RST!periode
        date1 = RST!tglawal
        date2 = RST!tglakhir
        
        SQL = "select * from gl_comptype where kdtype = '" & txtype & "'"
        Set RST = OBJ.Execute(SQL)
        lblnama = txtype & " - " & RST!nmtype
        
        SQL = "select * from gl_chacct where kdcomp = '" & txtkode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            UpDown1.Enabled = False
            date1.Enabled = False
            date2.Enabled = False
        End If
        
        txtkode.Enabled = False
        cmdsearch.Enabled = False
        
        txtprint.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Company " & txtkode & " Not Found.", vbInformation, "Information"
    txtkode = ""
    txtkode.SetFocus
End Sub

Private Sub txtkota_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtelp.SetFocus
End Sub

Private Sub txtlayout_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = 120 Or KeyAscii = 45 Or KeyAscii = vbKeyBack Or KeyAscii = 32 Or KeyAscii = 46) Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtprint_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtscreen.SetFocus
End Sub

Private Sub txtprint_KeyUp(KeyCode As Integer, Shift As Integer)
    txtscreen = txtprint
End Sub

Private Sub txtscreen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtalamat.SetFocus
End Sub
