VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmcustomeredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Customer"
   ClientHeight    =   7005
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
   ScaleHeight     =   7005
   ScaleMode       =   0  'User
   ScaleWidth      =   5938.03
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
      TabIndex        =   47
      Top             =   5775
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   240
      TabIndex        =   40
      Top             =   1800
      Visible         =   0   'False
      Width           =   5415
      Begin VB.OptionButton Option7 
         Caption         =   "Ubah Limit Piutang Customer, Aktif/Non Aktifkan Limit"
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
         Left            =   120
         TabIndex        =   46
         Top             =   1455
         Width           =   4260
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Ubah Status Customer (Aktif/Non Aktif)."
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
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   3615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Ubah Nama-NPWP, Alamat-NPWP, dan No-NPWP."
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
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Ubah AlamatCustomer, Kota, KodeArea, Telephone, Faxsimile, ContactPerson, Pabrik/Toko, Term, dan KawasanBerikat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   5175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ubah Nama Customer."
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin Chameleon.chameleonButton cmdch1 
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   1815
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "frmcustomeredit.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdch7 
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   1815
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "frmcustomeredit.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
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
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   3975
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
      MaxLength       =   50
      TabIndex        =   13
      Top             =   4440
      Width           =   3975
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
   Begin VB.TextBox txtkode 
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
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
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
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   6495
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
      MICON           =   "frmcustomeredit.frx":0634
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
      TabIndex        =   23
      Top             =   6495
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
      MICON           =   "frmcustomeredit.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Calculator      =   "frmcustomeredit.frx":0C68
      Caption         =   "frmcustomeredit.frx":0C88
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcustomeredit.frx":0CED
      Keys            =   "frmcustomeredit.frx":0D0B
      Spin            =   "frmcustomeredit.frx":0D55
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
   Begin TDBMask6Ctl.TDBMask txtnonpwp 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   5400
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   503
      Caption         =   "frmcustomeredit.frx":0D7D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmcustomeredit.frx":0DE2
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
   Begin Chameleon.chameleonButton cmdsearch0 
      Height          =   285
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Customer"
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
      MICON           =   "frmcustomeredit.frx":0E24
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
      Left            =   240
      TabIndex        =   37
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
      MICON           =   "frmcustomeredit.frx":113E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdch8 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   6495
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
      MICON           =   "frmcustomeredit.frx":1458
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdch9 
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   6495
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
      MICON           =   "frmcustomeredit.frx":1772
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox checkstatus 
      Height          =   270
      Left            =   3720
      TabIndex        =   42
      Top             =   150
      Visible         =   0   'False
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
      TabIndex        =   44
      Top             =   5760
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   503
      Calculator      =   "frmcustomeredit.frx":1A8C
      Caption         =   "frmcustomeredit.frx":1AAC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcustomeredit.frx":1B18
      Keys            =   "frmcustomeredit.frx":1B36
      Spin            =   "frmcustomeredit.frx":1B78
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
      ValueVT         =   361168901
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
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
      Left            =   1680
      TabIndex        =   49
      Top             =   6105
      Width           =   1665
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -15
      TabIndex        =   48
      Top             =   6360
      Width           =   5895
   End
   Begin VB.Label Label1 
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
      TabIndex        =   45
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NON AKTIF "
      BeginProperty Font 
         Name            =   "Clarendon Blk BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3960
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   1680
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
      TabIndex        =   39
      Top             =   3990
      Width           =   1335
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
      TabIndex        =   38
      Top             =   1800
      Width           =   3375
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
      TabIndex        =   35
      Top             =   4830
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
      TabIndex        =   34
      Top             =   4470
      Width           =   1335
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
      TabIndex        =   33
      Top             =   5430
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
      TabIndex        =   32
      Top             =   3630
      Width           =   2775
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
      TabIndex        =   31
      Top             =   1470
      Width           =   1215
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
      TabIndex        =   30
      Top             =   510
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
      TabIndex        =   29
      Top             =   870
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
      TabIndex        =   28
      Top             =   2190
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
      TabIndex        =   27
      Top             =   2550
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
      TabIndex        =   26
      Top             =   2910
      Width           =   1335
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
End
Attribute VB_Name = "frmcustomeredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2 As String

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub caricustomer()
    If txtKode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtNama = RST!namacust
        txtAlamat = RST!alamatcust
        txtkota = RST!kota
        txtTelp = RST!telpcust
        txterm = RST!termcust
        txtFax = RST!faxcust
        txtKontak = RST!contactperson
        txtKodePos = RST!kodepos
        If txtKodePos = "1" Then Option1.Value = True Else Option2.Value = True
        txtNamaNPWP = RST!namanpwp
        txtAlamatNPWP = RST!alamatnpwp
        txtnonpwp = RST!nonpwp
        txtarea = RST!kodearea
        If RST!berikat = 0 Then Check1.Value = 0
        If RST!berikat = 1 Then Check1.Value = 1
        If RST!Status = 0 Then lblstatus.Caption = " Aktif ": lblstatus.ForeColor = &HC000&: checkstatus.Value = xtpChecked
        If RST!Status = 1 Then lblstatus.Caption = " Non Aktif ": lblstatus.ForeColor = &HFF&: checkstatus.Value = xtpUnchecked
        txtlimit = RST!limit
        If RST!flaglimit = 0 Then Check2.Value = 0
        If RST!flaglimit = 1 Then Check2.Value = 1
        checkstatus.Visible = True
        lblstatus.Visible = True
        SQL = "select * from am_area where kode = '" & txtarea & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblarea = RST!nama
        
        OBJ.Close
        txtKode.Enabled = False
        cmdsearch0.Enabled = False
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then Check2.Caption = "Aktif": lbllimit = "Limit piutang aktif."
    If Check2.Value = 0 Then Check2.Caption = "Non Aktif": lbllimit = "Limit piutang non aktif."
End Sub

Private Sub checkstatus_Click()
    If checkstatus.Value = xtpChecked Then
        lblstatus.Caption = " Aktif ": lblstatus.ForeColor = &HC000&
    Else
        lblstatus.Caption = " Non Aktif ": lblstatus.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdch1_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdch7_Click()
    OBJ.Open dsn
    SQL = "select id2,kode2 from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str1 = RST!kode2 Else str1 = "0"
    OBJ.Close
    
    If Mid(txtKode, 3, 1) = str1 Then
        MsgBox "User tidak bisa mengupdate Data customer dari cabang lain.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Option3.Value Then
        If txtKode = "" Or txtNama = "" Then
            MsgBox "Data entry not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "select * from AM_invhdr WHERE kodecust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ.Close
            MsgBox "Can not update, data in use.", vbInformation, "Information"
            Exit Sub
        End If
        
        SQL = "select * from AM_autoaccust WHERE kodecust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str2 = RST!noacc
            OBJ.Close
            
            OBJ.Open dsn1
            SQL = "select * from gl_transaksi WHERE noactrx = '" & str2 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                OBJ.Close
                MsgBox "Can not update, data in use.", vbInformation, "Information"
                Exit Sub
            End If
            OBJ.Close
            OBJ.Open dsn
        End If
        
        SQL = "UPDATE AM_Customer SET "
        SQL = SQL + "NamaCust = '" & txtNama & "'"
        SQL = SQL + ",idUpdate = '" & kuser & "'"
        SQL = SQL + ",DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
        SQL = SQL + " WHERE KodeCust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    ElseIf Option4.Value Then
        If txtAlamat = "" Or txtarea = "" Or txtKodePos = "" Then
            MsgBox "Data entry not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "UPDATE AM_Customer SET "
        SQL = SQL + "AlamatCust = '" & txtAlamat & "'"
        SQL = SQL + ",kota = '" & txtkota & "'"
        SQL = SQL + ",kodearea = '" & txtarea & "'"
        SQL = SQL + ",TelpCust = '" & txtTelp & "'"
        SQL = SQL + ",FaxCust = '" & txtFax & "'"
        SQL = SQL + ",ContactPerson = '" & txtKontak & "'"
        SQL = SQL + ",kodepos = '" & txtKodePos & "'"
        SQL = SQL + ",termcust = convert(money,'" & txterm & "')"
        If Check1.Value = 0 Then SQL = SQL + ",berikat = '0'"
        If Check1.Value = 1 Then SQL = SQL + ",berikat = '1'"
        SQL = SQL + ",idUpdate = '" & kuser & "'"
        SQL = SQL + ",DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
        SQL = SQL + " WHERE KodeCust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    ElseIf Option5.Value Then
        OBJ.Open dsn
        SQL = "UPDATE AM_Customer SET "
        SQL = SQL + "namanpwp = '" & txtNamaNPWP & "'"
        SQL = SQL + ",alamatnpwp = '" & txtAlamatNPWP & "'"
        SQL = SQL + ",nonpwp = '" & txtnonpwp & "'"
        SQL = SQL + ",idUpdate = '" & kuser & "'"
        SQL = SQL + ",DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
        SQL = SQL + " WHERE KodeCust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    ElseIf Option6.Value Then
        OBJ.Open dsn
        SQL = "UPDATE AM_Customer SET "
        If checkstatus.Value = xtpChecked Then
            SQL = SQL + "status = '0'"
        Else
            SQL = SQL + "status = '1'"
        End If
        SQL = SQL + " Where KodeCust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    ElseIf Option7.Value Then
        OBJ.Open dsn
        SQL = "UPDATE am_Customer SET limit='" & txtlimit.Value & "'"
        If Check2.Value = 0 Then SQL = SQL + ",flaglimit = '0'"
        If Check2.Value = 1 Then SQL = SQL + ",flaglimit = '1'"
        SQL = SQL + " Where KodeCust = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    
    MsgBox "Data updated, click ok to continue ...", vbInformation, "Information"
    Frame1.Visible = False
    cmdclear_Click
End Sub

Private Sub cmdch8_Click()
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    Frame1.Visible = True
    Option3.SetFocus
End Sub

Private Sub cmdch9_Click()
    If txtKode = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from AM_sohdr WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo gakbisahapus
    
    SQL = "select * from AM_soapp WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo gakbisahapus

    SQL = "select * from AM_sjhdr WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo gakbisahapus
    
    SQL = "select * from AM_sjapp WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo gakbisahapus

    SQL = "select * from AM_invhdr WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo gakbisahapus
    
    SQL = "select * from AM_autoaccust WHERE kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!noacc
        
        OBJ.Close
        
        OBJ.Open dsn1
        SQL = "select * from gl_transaksi WHERE noactrx = '" & str2 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            GoTo gakbisahapus
        Else
            OBJ.Close
            OBJ.Open dsn
        End If
    End If

    SQL = "DELETE FROM AM_customer WHERE Kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE FROM AM_autoaccust WHERE Kodecust = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    OBJ.Open dsn1
    SQL = "DELETE FROM gl_masterac WHERE noac = '" & str2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE FROM gl_chacct WHERE noac = '" & str2 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data deleted, click ok to continue ...", , "Collector"
    cmdclear_Click
    
    Exit Sub
    
gakbisahapus:
    OBJ.Close
    MsgBox "Can not delete, data in use.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtKode.Enabled = True
    cmdsearch0.Enabled = True
    txtKode = ""
    txtNama = ""
    txtAlamat = ""
    txtkota = ""
    txtTelp = ""
    txtFax = ""
    txtKontak = ""
    txtNamaNPWP = ""
    txtAlamatNPWP = ""
    txtnonpwp = ""
    txtKodePos = ""
    txtlimit = "0.00"
    Option1.Value = False
    Option2.Value = False
    txtarea = ""
    lblarea = ""
    txterm = 0
    Check1.Value = 0
    Check2.Value = 0
    checkstatus.Visible = False
    lblstatus.Visible = False
    txtKode.SetFocus
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kode, nama from am_area"
    namatabel = "Area"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea = hasil
    lblarea = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtTelp.SetFocus
End Sub

Private Sub cmdsearch0_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch0_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    caricustomer
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtNama.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='62' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdch8.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='63' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdch9.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
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
    If txtarea = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_area where kode = '" & txtarea & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblarea = RST!nama
    Else
        MsgBox "Data not found.", vbInformation, "Information"
        txtarea = ""
        lblarea = ""
        txtarea.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKontak.SetFocus
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub txtkode_LostFocus()
    caricustomer
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
