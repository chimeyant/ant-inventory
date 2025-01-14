VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmaccountedit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Account"
   ClientHeight    =   5745
   ClientLeft      =   5715
   ClientTop       =   5565
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmaccountedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "As Header"
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
      Left            =   3120
      TabIndex        =   34
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "No. Account diatas sebagai Rugi Laba s/d Tahun Lalu"
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   4215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "No. Account diatas sebagai Rugi Laba s/d Bulan Lalu"
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
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   4215
   End
   Begin VB.TextBox lbltype1 
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
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox lbltype5 
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
      MaxLength       =   2
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox lbltype4 
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
      MaxLength       =   2
      TabIndex        =   6
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox lbltype3 
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
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox lbltype2 
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
      MaxLength       =   2
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox cmbtype 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmaccountedit.frx":2372
      Left            =   1800
      List            =   "frmaccountedit.frx":2374
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtnama 
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
      Top             =   1560
      Width           =   5535
   End
   Begin TDBText6Ctl.TDBText txtkode 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      Caption         =   "frmaccountedit.frx":2376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaccountedit.frx":23E2
      Key             =   "frmaccountedit.frx":2400
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
      MaxLength       =   10
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
      TabIndex        =   28
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Comp. Type 1"
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
      MICON           =   "frmaccountedit.frx":243C
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
      Left            =   360
      TabIndex        =   29
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Comp. Type 2"
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
      MICON           =   "frmaccountedit.frx":2756
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
      TabIndex        =   30
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Comp. Type 3"
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
      MICON           =   "frmaccountedit.frx":2A70
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
      Left            =   360
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Comp. Type 4"
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
      MICON           =   "frmaccountedit.frx":2D8A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch5 
      Height          =   285
      Left            =   360
      TabIndex        =   32
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Comp. Type 5"
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
      MICON           =   "frmaccountedit.frx":30A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   5160
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
      MICON           =   "frmaccountedit.frx":33BE
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
      TabIndex        =   11
      Top             =   5160
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
      MICON           =   "frmaccountedit.frx":36D8
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
      TabIndex        =   12
      Top             =   5160
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
      MICON           =   "frmaccountedit.frx":39F2
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
      TabIndex        =   13
      Top             =   5160
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
      MICON           =   "frmaccountedit.frx":3D0C
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
      TabIndex        =   33
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No. Account"
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
      MICON           =   "frmaccountedit.frx":4026
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Master Account"
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
      TabIndex        =   26
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label7 
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
      TabIndex        =   25
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lbltypeacc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblnama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblnama5 
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
      Left            =   2400
      TabIndex        =   22
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label lblnama4 
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
      Left            =   2400
      TabIndex        =   21
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label lblnama3 
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
      Left            =   2400
      TabIndex        =   20
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label lblnama2 
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
      Left            =   2400
      TabIndex        =   19
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Jenis Perusahaan yang memakai No.Account diatas."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2400
      Width           =   4395
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   3840
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblnama1 
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
      Left            =   2400
      TabIndex        =   16
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Nama Account"
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
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Type Account"
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
      Top             =   1950
      Width           =   1395
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmaccountedit"
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

Dim str1 As String

Private Sub cmbtype_Click()
    Option1.Enabled = True
    Option2.Enabled = True
    Option1.Value = False
    Option2.Value = False
    Option1.Enabled = False
    Option2.Enabled = False
    
    Select Case cmbtype.text
    Case "Assets"
        lbltype = "AS"
    Case "Liability"
        lbltype = "LI"
    Case "Capital"
        lbltype = "CA"
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        OBJ.Open dsn
        SQL = "select * from gl_accrl"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!rl_ptd <> "" Then Option2.Enabled = False
            If RST!rl_ytd <> "" Then Option1.Enabled = False
        End If
        OBJ.Close
    Case "Income"
        lbltype = "IN"
    Case "Expenses"
        lbltype = "EX"
    Case "Income Summary"
        lbltype = "IS"
    End Select
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    txtKode = ""
    txtNama = ""
    cmbtype = ""
    lbltype = ""
    lblnama1 = ""
    lblnama2 = ""
    lblnama3 = ""
    lblnama4 = ""
    lblnama5 = ""
    lbltype1 = ""
    lbltype2 = ""
    lbltype3 = ""
    lbltype4 = ""
    lbltype5 = ""
    Option1.Enabled = True
    Option2.Enabled = True
    Option1.Value = False
    Option2.Value = False
    Option1.Enabled = False
    Option2.Enabled = False
    txtKode.Enabled = True
    cmdsearch.Enabled = True
    cmdsearch1.Enabled = True
    cmdsearch2.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    cmdsearch5.Enabled = True
    cmbtype.Enabled = True
    Check1.Value = 0
    txtKode.SetFocus
    cmbtype.Clear
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdelete_Click()
    OBJ.Open dsn
    SQL = "select ac_supp from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!ac_supp = Mid(txtKode, 1, 2) Then
            MsgBox "Can not delete account. (account for supplier)", vbInformation, "Information"
            cmdclear_Click
            OBJ.Close
            Exit Sub
        End If
    End If
    OBJ.Close
    
    If txtKode = "" Or cmbtype = "" Or txtNama = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_accrl where rl_ptd = '" & txtKode & "' or rl_ytd = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!rl_ptd = txtKode Then
            If MsgBox("Delete this account for R/L s/d Bulan Lalu ?", vbQuestion + vbYesNo, "Question") = vbNo Then
                OBJ.Close
                Exit Sub
            End If
        End If
        If RST!rl_ytd = txtKode Then
            If MsgBox("Delete this account for R/L s/d Tahun Lalu ?", vbQuestion + vbYesNo, "Question") = vbNo Then
                OBJ.Close
                Exit Sub
            End If
        End If
        SQL = "select * from gl_transaksi where flag = 'P'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then GoTo jump
    End If
    
    SQL = "select * from gl_transaksi WHERE noactrx = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_dforms WHERE acc_no1 = '" & txtKode & "' or acc_no2 = '" & txtKode & "' or acc_no3 = '" & txtKode & "' or acc_no4 = '" & txtKode & "' or acc_no5 = '" & txtKode & "' or acc_no6 = '" & txtKode & "' or acc_no7 = '" & txtKode & "' or acc_no8 = '" & txtKode & "' or acc_no9 = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_aktiva WHERE ac_aktiva = '" & txtKode & "' or ac_susut = '" & txtKode & "' or ac_biaya = '" & txtKode & "' or ac_lawan = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump
    
    SQL = "select * from gl_accrl where rl_ptd = '" & txtKode & "' or rl_ytd = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If Option1.Enabled = True Then
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ytd = ''"
            Set RST = OBJ.Execute(SQL)
        End If
        If Option2.Enabled = True Then
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ptd = ''"
            Set RST = OBJ.Execute(SQL)
        End If
    End If
    
    SQL = "DELETE FROM gl_masterac WHERE noac = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "DELETE FROM gl_chacct WHERE noac = '" & txtKode & "'"
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

Private Sub cmdsearch_Click()
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Master Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    txtNama = hasil1
    txtKode_LostFocus
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    lbltype1 = hasil
    lblnama1 = hasil1
    str1 = hasil
    hasil = ""
    hasil1 = ""
    caritype (1)
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    lbltype2 = hasil
    lblnama2 = hasil1
    str1 = hasil
    hasil = ""
    hasil1 = ""
    caritype (2)
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    lbltype3 = hasil
    lblnama3 = hasil1
    str1 = hasil
    hasil = ""
    hasil1 = ""
    caritype (3)
End Sub

Private Sub cmdsearch4_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    lbltype4 = hasil
    lblnama4 = hasil1
    str1 = hasil
    hasil = ""
    hasil1 = ""
    caritype (4)
End Sub

Private Sub cmdsearch5_Click()
    carisql1 = "select kdtype, nmtype from gl_comptype"
    namatabel = "Company Type"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    lbltype5 = hasil
    lblnama5 = hasil1
    str1 = hasil
    hasil = ""
    hasil1 = ""
    caritype (5)
End Sub

Private Sub cmdupdate_click()
    If txtKode = "" Or cmbtype = "" Or txtNama = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If lbltype1 = "" And lbltype2 = "" And lbltype3 = "" And _
    lbltype4 = "" And lbltype5 = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If Option1.Value = True Then
        If MsgBox("Set this account for R/L s/d Tahun Lalu ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    End If
    If Option2.Value = True Then
        If MsgBox("Set this account for R/L s/d Bulan Lalu ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_accrl where rl_ptd = '" & txtKode & "' or rl_ytd = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "select * from gl_transaksi where flag = 'P'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then GoTo jump1
    End If
  
    If lbltype <> lbltypeacc Then
        SQL = "select * from gl_transaksi WHERE noactrx = '" & txtKode & "' and flag = 'P'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then GoTo jump1
        
        SQL = "select * from gl_dforms WHERE acc_no1 = '" & txtKode & "' or acc_no2 = '" & txtKode & "' or acc_no3 = '" & txtKode & "' or acc_no4 = '" & txtKode & "' or acc_no5 = '" & txtKode & "' or acc_no6 = '" & txtKode & "' or acc_no7 = '" & txtKode & "' or acc_no8 = '" & txtKode & "' or  acc_no9 = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then GoTo jump1
    
        SQL = "select * from gl_aktiva WHERE ac_aktiva = '" & txtKode & "' or ac_susut = '" & txtKode & "' or ac_biaya = '" & txtKode & "' or ac_lawan = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then GoTo jump1
    End If
    
    If Check1.Value = 1 Then
        SQL = "select * from gl_transaksi WHERE noactrx = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Cannot set account as header, account already in use.", vbInformation, "Information"
            OBJ.Close
            cmdclear_Click
            
            Exit Sub
        End If
    End If
    
    If lbltype1 <> "" Then
        If cmdsearch1.Enabled = True Then
            SQL = "select * from gl_company where kdtype = '" & lbltype1 & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                str1 = RST!kdcomp
                addchacct
                
                RST.MoveNext
            Loop
        End If
    End If
    If lbltype2 <> "" Then
        If cmdsearch2.Enabled = True Then
            SQL = "select * from gl_company where kdtype = '" & lbltype2 & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                str1 = RST!kdcomp
                addchacct
                
                RST.MoveNext
            Loop
        End If
    End If
    If lbltype3 <> "" Then
        If cmdsearch3.Enabled = True Then
            SQL = "select * from gl_company where kdtype = '" & lbltype3 & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                str1 = RST!kdcomp
                addchacct
                
                RST.MoveNext
            Loop
        End If
    End If
    If lbltype4 <> "" Then
        If cmdsearch4.Enabled = True Then
            SQL = "select * from gl_company where kdtype = '" & lbltype4 & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                str1 = RST!kdcomp
                addchacct
                    
                RST.MoveNext
            Loop
        End If
    End If
    If lbltype5 <> "" Then
        If cmdsearch5.Enabled = True Then
            SQL = "select * from gl_company where kdtype = '" & lbltype5 & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                str1 = RST!kdcomp
                addchacct
                
                RST.MoveNext
            Loop
        End If
    End If
    
    SQL = "UPDATE gl_masterac SET "
    SQL = SQL + "Nmac = '" & txtNama & "',"
    SQL = SQL + "flag = '" & Check1.Value & "',"
    SQL = SQL + "typeac = '" & lbltype & "',"
    SQL = SQL + "jenisac1 = '" & lbltype1 & "',"
    SQL = SQL + "jenisac2 = '" & lbltype2 & "',"
    SQL = SQL + "jenisac3 = '" & lbltype3 & "',"
    SQL = SQL + "jenisac4 = '" & lbltype4 & "',"
    SQL = SQL + "jenisac5 = '" & lbltype5 & "',"
    SQL = SQL + "idUpdate = '" & kuser & "',"
    SQL = SQL + "DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + "WHERE noac =  '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "UPDATE gl_chacct SET "
    SQL = SQL + "typeac = '" & lbltype & "'"
    SQL = SQL + "WHERE noac =  '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "select * from gl_accrl where rl_ptd = '" & txtKode & "' or rl_ytd = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If Option1.Value = False And Option1.Enabled = True Then
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ytd = ''"
            Set RST = OBJ.Execute(SQL)
        End If
        If Option2.Value = False And Option2.Enabled = True Then
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ptd = ''"
            Set RST = OBJ.Execute(SQL)
        End If
    End If
    If Option1.Value = True Then
        SQL = "select * from gl_accrl"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            SQL = "insert into gl_accrl"
            SQL = SQL + "(rl_ptd"
            SQL = SQL + ",rl_ytd)"
        
            SQL = SQL + "VALUES"
            SQL = SQL + "(''"
            SQL = SQL + ", '" & txtKode & "')"
            Set RST = OBJ.Execute(SQL)
        Else
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ytd = '" & txtKode & "'"
            Set RST = OBJ.Execute(SQL)
        End If
    End If
    If Option2.Value = True Then
        SQL = "select * from gl_accrl"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            SQL = "insert into gl_accrl"
            SQL = SQL + "(rl_ytd"
            SQL = SQL + ",rl_ptd)"
        
            SQL = SQL + "VALUES"
            SQL = SQL + "(''"
            SQL = SQL + ", '" & txtKode & "')"
            Set RST = OBJ.Execute(SQL)
        Else
            SQL = "UPDATE gl_accrl SET "
            SQL = SQL + "rl_ptd = '" & txtKode & "'"
            Set RST = OBJ.Execute(SQL)
        End If
    End If
    OBJ.Close
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    Exit Sub
    
jump1:
    MsgBox "Can Not Update, Record Still In Use.", vbInformation, "Information"
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub addchacct()
    OBJ1.Open dsn
    SQL1 = "select * from gl_chacct where kdcomp = '" & str1 & "' and noac = '" & txtKode & "'"
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
    SQL1 = SQL1 + "('" & str1 & "'"
    SQL1 = SQL1 + ", '" & txtKode & "'"
    SQL1 = SQL1 + ", '" & lbltype & "'"
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Option1_DblClick()
    If Option1.Enabled = True And Option1.Value = True And Option2.Enabled = False Then
        Option1.Value = False
    End If
End Sub

Private Sub Option2_DblClick()
    If Option2.Enabled = True And Option2.Value = True And Option1.Enabled = False Then
        Option2.Value = False
    End If
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    If txtKode = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        cmbtype.Clear
        
        cmbtype.AddItem "Assets"
        cmbtype.AddItem "Liability"
        cmbtype.AddItem "Capital"
        cmbtype.AddItem "Income"
        cmbtype.AddItem "Expenses"
        cmbtype.AddItem "Income Summary"
        
        txtNama = RST!nmac
        lbltype = RST!typeac
        Check1.Value = RST!flag
        lblnama = RST!nmac
        lbltypeacc = RST!typeac
        
        If lbltype <> "IS" Then
            OBJ1.Open dsn
            SQL1 = "select * from gl_masterac where typeac = 'IS'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then cmbtype.RemoveItem 5
            OBJ1.Close
        End If
        
        lbltype1 = RST!jenisac1
        lbltype2 = RST!jenisac2
        lbltype3 = RST!jenisac3
        lbltype4 = RST!jenisac4
        lbltype5 = RST!jenisac5
        
        Select Case lbltype
        Case "AS"
            cmbtype = "Assets"
        Case "LI"
            cmbtype = "Liability"
        Case "CA"
            cmbtype = "Capital"
            
            Option1.Enabled = True
            Option2.Enabled = True
            
            SQL = "select * from gl_accrl"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If RST!rl_ptd <> "" Then Option2.Enabled = False
                If RST!rl_ytd <> "" Then Option1.Enabled = False
            End If
        Case "IN"
            cmbtype = "Income"
        Case "EX"
            cmbtype = "Expenses"
        Case "IS"
            cmbtype = "Income Summary"
        End Select
                
        If lbltype1 <> "" Then
            SQL = "select * from gl_comptype where kdtype = '" & lbltype1 & "'"
            Set RST = OBJ.Execute(SQL)
            lblnama1 = RST!nmtype
            
            cmdsearch1.Enabled = False
        End If
        If lbltype2 <> "" Then
            SQL = "select * from gl_comptype where kdtype = '" & lbltype2 & "'"
            Set RST = OBJ.Execute(SQL)
            lblnama2 = RST!nmtype
            
            cmdsearch2.Enabled = False
        End If
        If lbltype3 <> "" Then
            SQL = "select * from gl_comptype where kdtype = '" & lbltype3 & "'"
            Set RST = OBJ.Execute(SQL)
            lblnama3 = RST!nmtype
            
            cmdsearch3.Enabled = False
        End If
        If lbltype4 <> "" Then
            SQL = "select * from gl_comptype where kdtype = '" & lbltype4 & "'"
            Set RST = OBJ.Execute(SQL)
            lblnama4 = RST!nmtype
            
            cmdsearch4.Enabled = False
        End If
        If lbltype5 <> "" Then
            SQL = "select * from gl_comptype where kdtype = '" & lbltype5 & "'"
            Set RST = OBJ.Execute(SQL)
            lblnama5 = RST!nmtype
            
            cmdsearch5.Enabled = False
        End If
        
        SQL = "select * from gl_accrl where rl_ptd = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            Option2.Enabled = True
            Option2.Value = True
        End If
        
        SQL = "select * from gl_accrl where rl_ytd = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            Option1.Enabled = True
            Option1.Value = True
        End If
        
        SQL = "select * from gl_chacct where noac = '" & txtKode & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!periode01 <> 0 Or RST!periode02 <> 0 Or RST!periode03 <> 0 Or _
            RST!periode04 <> 0 Or RST!periode05 <> 0 Or RST!periode06 <> 0 Or _
            RST!periode07 <> 0 Or RST!periode08 <> 0 Or RST!periode09 <> 0 Or _
            RST!periode10 <> 0 Or RST!periode11 <> 0 Or RST!periode12 <> 0 Or _
            RST!periode13 <> 0 Or _
            RST!last01 <> 0 Or RST!last02 <> 0 Or RST!last03 <> 0 Or _
            RST!last04 <> 0 Or RST!last05 <> 0 Or RST!last06 <> 0 Or _
            RST!last07 <> 0 Or RST!last08 <> 0 Or RST!last09 <> 0 Or _
            RST!last10 <> 0 Or RST!last11 <> 0 Or RST!last12 <> 0 Or _
            RST!last13 <> 0 Or RST!balancedb <> 0 Or RST!balancecr <> 0 Or _
            RST!begindb <> 0 Or RST!begincr <> 0 Then cmbtype.Enabled = False
        End If
    
        cmdsearch.Enabled = False
        txtKode.Enabled = False
        txtNama.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Account " & txtKode & " Not Found.", vbInformation, "Information"
    txtKode = ""
    txtKode.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmbtype.Enabled = True Then cmbtype.SetFocus
End Sub

Private Sub caritype(ByVal teks As Integer)
    If lbltype1 <> "" And teks <> 1 Then
        If lbltype1 = str1 Then
            MsgBox "Type Already Exist", vbInformation, "Information"
            hapus (teks)
            Exit Sub
        End If
    End If
    If lbltype2 <> "" And teks <> 2 Then
        If lbltype2 = str1 Then
            MsgBox "Type Already Exist", vbInformation, "Information"
            hapus (teks)
            Exit Sub
        End If
    End If
    If lbltype3 <> "" And teks <> 3 Then
        If lbltype3 = str1 Then
            MsgBox "Type Already Exist", vbInformation, "Information"
            hapus (teks)
            Exit Sub
        End If
    End If
    If lbltype4 <> "" And teks <> 4 Then
        If lbltype4 = str1 Then
            MsgBox "Type Already Exist", vbInformation, "Information"
            hapus (teks)
            Exit Sub
        End If
    End If
    If lbltype5 <> "" And teks <> 5 Then
        If lbltype5 = str1 Then
            MsgBox "Type Already Exist", vbInformation, "Information"
            hapus (teks)
            Exit Sub
        End If
    End If
End Sub

Private Sub hapus(ByVal no As Integer)
    If no = 1 Then
        lbltype1 = ""
        lblnama1 = ""
        str1 = ""
        lbltype1.SetFocus
    ElseIf no = 2 Then
        lbltype2 = ""
        lblnama2 = ""
        str1 = ""
        lbltype2.SetFocus
    ElseIf no = 3 Then
        lbltype3 = ""
        lblnama3 = ""
        str1 = ""
        lbltype3.SetFocus
    ElseIf no = 4 Then
        lbltype4 = ""
        lblnama4 = ""
        str1 = ""
        lbltype4.SetFocus
    ElseIf no = 5 Then
        lbltype5 = ""
        lblnama5 = ""
        str1 = ""
        lbltype5.SetFocus
    End If
End Sub

Private Sub lbltype1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then lbltype2.SetFocus
End Sub

Private Sub lbltype1_LostFocus()
    If lbltype1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_comptype where kdtype = '" & lbltype1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama1 = RST!nmtype
        str1 = RST!kdtype
    Else
        MsgBox "Type Company" & lbltype1 & " Not Found.", vbInformation, "Information"
        lbltype1 = ""
        lbltype1.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    caritype (1)
End Sub

Private Sub lbltype2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then lbltype3.SetFocus
End Sub

Private Sub lbltype2_LostFocus()
    If lbltype2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_comptype where kdtype = '" & lbltype2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama2 = RST!nmtype
        str1 = RST!kdtype
    Else
        MsgBox "Type Company" & lbltype2 & " Not Found.", vbInformation, "Information"
        lbltype2 = ""
        lbltype2.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    caritype (2)
End Sub

Private Sub lbltype3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then lbltype4.SetFocus
End Sub

Private Sub lbltype3_LostFocus()
    If lbltype3 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_comptype where kdtype = '" & lbltype3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama3 = RST!nmtype
        str1 = RST!kdtype
    Else
        MsgBox "Type Company" & lbltype3 & " Not Found.", vbInformation, "Information"
        lbltype3 = ""
        lbltype3.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    caritype (3)
End Sub

Private Sub lbltype4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then lbltype5.SetFocus
End Sub

Private Sub lbltype4_LostFocus()
    If lbltype4 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_comptype where kdtype = '" & lbltype4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama4 = RST!nmtype
        str1 = RST!kdtype
    Else
        MsgBox "Type Company" & lbltype4 & " Not Found.", vbInformation, "Information"
        lbltype4 = ""
        lbltype4.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    caritype (4)
End Sub

Private Sub lbltype5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdupdate.SetFocus
End Sub

Private Sub lbltype5_LostFocus()
    If lbltype5 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_comptype where kdtype = '" & lbltype5 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama5 = RST!nmtype
        str1 = RST!kdtype
    Else
        MsgBox "Type Company" & lbltype5 & " Not Found.", vbInformation, "Information"
        lbltype5 = ""
        lbltype5.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    caritype (5)
End Sub
