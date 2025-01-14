VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmreportdetail 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmreportdetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtacc1 
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtacc2 
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtacc3 
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtacc4 
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtacc5 
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtacc6 
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
      Height          =   315
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtacc7 
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
      Height          =   315
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtacc8 
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
      Height          =   315
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtacc9 
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
      Height          =   315
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox cmblineno 
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
      ItemData        =   "frmreportdetail.frx":2372
      Left            =   1440
      List            =   "frmreportdetail.frx":2427
      TabIndex        =   0
      Top             =   1800
      Width           =   615
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   4200
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
      MICON           =   "frmreportdetail.frx":250E
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
      Left            =   3480
      TabIndex        =   11
      Top             =   4200
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
      MICON           =   "frmreportdetail.frx":2828
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
      Left            =   4440
      TabIndex        =   12
      Top             =   4200
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
      MICON           =   "frmreportdetail.frx":2B42
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
      Left            =   5400
      TabIndex        =   13
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
      MICON           =   "frmreportdetail.frx":2E5C
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
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #1"
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
      MICON           =   "frmreportdetail.frx":3176
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
      Left            =   240
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #2"
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
      MICON           =   "frmreportdetail.frx":3490
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
      Left            =   240
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #3"
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
      MICON           =   "frmreportdetail.frx":37AA
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
      Left            =   240
      TabIndex        =   31
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #4"
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
      MICON           =   "frmreportdetail.frx":3AC4
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
      Left            =   240
      TabIndex        =   32
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #5"
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
      MICON           =   "frmreportdetail.frx":3DDE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch6 
      Height          =   285
      Left            =   3600
      TabIndex        =   33
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #6"
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
      MICON           =   "frmreportdetail.frx":40F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch7 
      Height          =   285
      Left            =   3600
      TabIndex        =   34
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #7"
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
      MICON           =   "frmreportdetail.frx":4412
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch8 
      Height          =   285
      Left            =   3600
      TabIndex        =   35
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #8"
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
      MICON           =   "frmreportdetail.frx":472C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch9 
      Height          =   285
      Left            =   3600
      TabIndex        =   36
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #9"
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
      MICON           =   "frmreportdetail.frx":4A46
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
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Type Group"
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
      TabIndex        =   26
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label lbltype2 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Top             =   1350
      Width           =   4215
   End
   Begin VB.Label lbltype1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Top             =   630
      Width           =   4215
   End
   Begin VB.Label lblcode1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lbldesc1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   390
      Width           =   4215
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      Caption         =   "Line No. #"
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
      Top             =   1830
      Width           =   1095
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Group No."
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
      Top             =   870
      Width           =   855
   End
   Begin VB.Label lblgroup1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Top             =   870
      Width           =   975
   End
   Begin VB.Label lbldesc2 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   1110
      Width           =   4215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Report Code"
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
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Type Report"
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
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   -120
      TabIndex        =   27
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmreportdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmblineno_Click()
    If cmblineno = "" Then Exit Sub
    cariline
End Sub

Private Sub cariline()
    hapusdetail
    If cmblineno >= 1 And cmblineno <= 50 Then
        txtacc2.Enabled = True
        txtacc3.Enabled = True
        txtacc4.Enabled = True
        txtacc5.Enabled = True
        txtacc6.Enabled = True
        txtacc7.Enabled = True
        txtacc8.Enabled = True
        txtacc9.Enabled = True
        cmdsearch2.Enabled = True
        cmdsearch3.Enabled = True
        cmdsearch4.Enabled = True
        cmdsearch5.Enabled = True
        cmdsearch6.Enabled = True
        cmdsearch7.Enabled = True
        cmdsearch8.Enabled = True
        cmdsearch9.Enabled = True
    Else
        txtacc2.Enabled = False
        txtacc3.Enabled = False
        txtacc4.Enabled = False
        txtacc5.Enabled = False
        txtacc6.Enabled = False
        txtacc7.Enabled = False
        txtacc8.Enabled = False
        txtacc9.Enabled = False
        cmdsearch2.Enabled = False
        cmdsearch3.Enabled = False
        cmdsearch4.Enabled = False
        cmdsearch5.Enabled = False
        cmdsearch6.Enabled = False
        cmdsearch7.Enabled = False
        cmdsearch8.Enabled = False
        cmdsearch9.Enabled = False
    End If
End Sub

Private Sub cmblineno_KeyPress(KeyAscii As Integer)
    If Len(cmblineno) = 2 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = vbKeyBack Or KeyAscii = 45) Then KeyAscii = 0
End Sub

Private Sub cmblineno_LostFocus()
    If cmblineno = "" Then Exit Sub
    If Not ((cmblineno >= -9 And cmblineno <= -1) Or (cmblineno >= 1 And cmblineno <= 50)) Then
        cmblineno = ""
        cmblineno.SetFocus
        Exit Sub
    End If
    cariline
End Sub

Private Sub cmdclear_Click()
    If cmblineno.Enabled = False Then Exit Sub
    cmblineno.Enabled = True
    cmblineno = ""
    cmblineno.SetFocus
    hapusdetail
    enable_detail
End Sub

Private Sub hapusdetail()
    txtacc1 = ""
    txtacc2 = ""
    txtacc3 = ""
    txtacc4 = ""
    txtacc5 = ""
    txtacc6 = ""
    txtacc7 = ""
    txtacc8 = ""
    txtacc9 = ""
End Sub

Private Sub enable_detail()
    txtacc2.Enabled = True
    txtacc3.Enabled = True
    txtacc4.Enabled = True
    txtacc5.Enabled = True
    txtacc6.Enabled = True
    txtacc7.Enabled = True
    txtacc8.Enabled = True
    txtacc9.Enabled = True
    cmdsearch2.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    cmdsearch5.Enabled = True
    cmdsearch6.Enabled = True
    cmdsearch7.Enabled = True
    cmdsearch8.Enabled = True
    cmdsearch9.Enabled = True
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdelete_Click()
    If cmblineno = "" Or txtacc1 = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    frmreport.grid2.Row = setup1
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 1) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 2) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 3) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 4) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 5) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 6) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 7) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 8) = ""
    frmreport.grid2.TextMatrix(frmreport.grid2.Row, 9) = ""
    Do While True
        If frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 1) = "" Then
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 1) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 2) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 3) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 4) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 5) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 6) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 7) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 8) = ""
            frmreport.grid2.TextMatrix(frmreport.grid2.Row, 9) = ""
            Exit Do
        End If
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 0)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 1) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 1)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 2) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 2)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 3) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 3)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 4) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 4)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 5) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 5)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 6) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 6)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 7) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 7)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 8) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 8)
        frmreport.grid2.TextMatrix(frmreport.grid2.Row, 9) = frmreport.grid2.TextMatrix(frmreport.grid2.Row + 1, 9)
        frmreport.grid2.Row = frmreport.grid2.Row + 1
    Loop
    frmreport.grid2.Rows = frmreport.grid2.Rows - 1
    frmreport.lbltotgroup = "Total Group : " & frmreport.grid2.Rows - 2
    
    For z = setup1 To 99
        For y = 0 To 10
            myarray(setup3, y, z) = myarray(setup3, y, z + 1)
        Next y
    Next z
    
    MsgBox "Data Detail Group Report Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch2_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch3_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc3 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch4_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc4 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch5_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc5 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch6_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch6_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc6 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch7_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch7_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc7 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch8_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch8_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc8 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch9_Click()
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch9_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc9 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdSave_Click()
    If cmblineno = "" Or (txtacc1 = "" And txtacc2 = "" And txtacc3 = "" And txtacc4 = "" And txtacc5 = "" And txtacc6 = "" And txtacc7 = "" And txtacc8 = "" And txtacc9 = "") Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    frmreport.grid2.Row = 1
    Do While cmblineno.Enabled = True
        If frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = "" Then Exit Do
        
        If frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = cmblineno Then
            MsgBox "Can't Add, Detail Group Report " & cmblineno & " Already Exsist.", vbInformation, "Information"
            cmdclear_Click
            Exit Sub
        End If
        frmreport.grid2.Row = frmreport.grid2.Row + 1
    Loop
    
    If cmblineno.Enabled = True Then
        frmreport.grid2.Row = 1
        Do While True
            If frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = "" Then
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 0) = cmblineno
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 1) = txtacc1
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 2) = txtacc2
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 3) = txtacc3
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 4) = txtacc4
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 5) = txtacc5
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 6) = txtacc6
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 7) = txtacc7
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 8) = txtacc8
                frmreport.grid2.TextMatrix(frmreport.grid2.Row, 9) = txtacc9
                
                myarray(setup3, 0, frmreport.grid2.Row) = cmblineno
                myarray(setup3, 1, frmreport.grid2.Row) = x_original(txtacc1)
                myarray(setup3, 2, frmreport.grid2.Row) = x_original(txtacc2)
                myarray(setup3, 3, frmreport.grid2.Row) = x_original(txtacc3)
                myarray(setup3, 4, frmreport.grid2.Row) = x_original(txtacc4)
                myarray(setup3, 5, frmreport.grid2.Row) = x_original(txtacc5)
                myarray(setup3, 6, frmreport.grid2.Row) = x_original(txtacc6)
                myarray(setup3, 7, frmreport.grid2.Row) = x_original(txtacc7)
                myarray(setup3, 8, frmreport.grid2.Row) = x_original(txtacc8)
                myarray(setup3, 9, frmreport.grid2.Row) = x_original(txtacc9)
                myarray(setup3, 10, frmreport.grid2.Row) = frmreport.grid1.TextMatrix(setup3, 0)
                
                frmreport.grid2.Rows = frmreport.grid2.Rows + 1
                Exit Do
            End If
            frmreport.grid2.Row = frmreport.grid2.Row + 1
        Loop
        frmreport.lbltotdetail = "Total Detail : " & frmreport.grid2.Rows - 2
        MsgBox "Detail Group Report Is Added, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
    Else
        frmreport.grid2.TextMatrix(setup1, 1) = txtacc1
        frmreport.grid2.TextMatrix(setup1, 2) = txtacc2
        frmreport.grid2.TextMatrix(setup1, 3) = txtacc3
        frmreport.grid2.TextMatrix(setup1, 4) = txtacc4
        frmreport.grid2.TextMatrix(setup1, 5) = txtacc5
        frmreport.grid2.TextMatrix(setup1, 6) = txtacc6
        frmreport.grid2.TextMatrix(setup1, 7) = txtacc7
        frmreport.grid2.TextMatrix(setup1, 8) = txtacc8
        frmreport.grid2.TextMatrix(setup1, 9) = txtacc9
        
        myarray(setup3, 1, setup1) = x_original(txtacc1)
        myarray(setup3, 2, setup1) = x_original(txtacc2)
        myarray(setup3, 3, setup1) = x_original(txtacc3)
        myarray(setup3, 4, setup1) = x_original(txtacc4)
        myarray(setup3, 5, setup1) = x_original(txtacc5)
        myarray(setup3, 6, setup1) = x_original(txtacc6)
        myarray(setup3, 7, setup1) = x_original(txtacc7)
        myarray(setup3, 8, setup1) = x_original(txtacc8)
        myarray(setup3, 9, setup1) = x_original(txtacc9)
        myarray(setup3, 10, setup1) = frmreport.grid1.TextMatrix(setup3, 0)
        
        MsgBox "Detail Group Report Is Updated, Click OK To Continue ...", vbInformation, "Information"
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    lblcode1 = frmreport.txtreportcode
    lbldesc1 = frmreport.txtdesc1
    If setup2 = 1 Then
        lbltype1 = "Balance Sheet"
    ElseIf setup2 = 2 Then
        lbltype1 = "Income Statement"
    ElseIf setup2 = 3 Then
        lbltype1 = "Cash Flow"
    End If
    lblgroup1 = frmreport.grid1.TextMatrix(setup3, 0)
    lbldesc2 = frmreport.grid1.TextMatrix(setup3, 2)
    lbltype2 = frmreport.grid1.TextMatrix(setup3, 7)
    If setup1 <> "" Then
        cmblineno = frmreport.grid2.TextMatrix(setup1, 0)
        txtacc1 = frmreport.grid2.TextMatrix(setup1, 1)
        txtacc2 = frmreport.grid2.TextMatrix(setup1, 2)
        txtacc3 = frmreport.grid2.TextMatrix(setup1, 3)
        txtacc4 = frmreport.grid2.TextMatrix(setup1, 4)
        txtacc5 = frmreport.grid2.TextMatrix(setup1, 5)
        txtacc6 = frmreport.grid2.TextMatrix(setup1, 6)
        txtacc7 = frmreport.grid2.TextMatrix(setup1, 7)
        txtacc8 = frmreport.grid2.TextMatrix(setup1, 8)
        txtacc9 = frmreport.grid2.TextMatrix(setup1, 9)
        cmblineno.Enabled = False
        If cmblineno >= 1 And cmblineno <= 50 Then
            txtacc2.Enabled = True
            txtacc3.Enabled = True
            txtacc4.Enabled = True
            txtacc5.Enabled = True
            txtacc6.Enabled = True
            txtacc7.Enabled = True
            txtacc8.Enabled = True
            txtacc9.Enabled = True
            cmdsearch2.Enabled = True
            cmdsearch3.Enabled = True
            cmdsearch4.Enabled = True
            cmdsearch5.Enabled = True
            cmdsearch6.Enabled = True
            cmdsearch7.Enabled = True
            cmdsearch8.Enabled = True
            cmdsearch9.Enabled = True
        Else
            txtacc2.Enabled = False
            txtacc3.Enabled = False
            txtacc4.Enabled = False
            txtacc5.Enabled = False
            txtacc6.Enabled = False
            txtacc7.Enabled = False
            txtacc8.Enabled = False
            txtacc9.Enabled = False
            cmdsearch2.Enabled = False
            cmdsearch3.Enabled = False
            cmdsearch4.Enabled = False
            cmdsearch5.Enabled = False
            cmdsearch6.Enabled = False
            cmdsearch7.Enabled = False
            cmdsearch8.Enabled = False
            cmdsearch9.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    setup1 = ""
    frmreport.SSTab1.Tab = 0
    frmreport.SSTab1.Tab = 1
End Sub

Private Sub txtacc1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtacc2.Enabled = True Then txtacc2.SetFocus
    If KeyAscii = 13 And txtacc2.Enabled = False Then cmdsave.SetFocus
End Sub

Private Sub txtacc1_LostFocus()
    If txtacc1 = "" Or cmblineno = "" Then Exit Sub
    OBJ.Open dsn
    cektype
    If cmblineno >= -9 And cmblineno <= -1 Then
        SQL = "select * from gl_masterac where noac like '" & x_original(txtacc1) & "%' and (typeac = '" & setup5 & "')"
    Else
        SQL = "select * from gl_masterac where noac = '" & x_original(txtacc1) & "' and (typeac = '" & setup5 & "')"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc1 & " Not Found.", vbInformation, "Information"
        txtacc1 = ""
        txtacc1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc3.SetFocus
End Sub

Private Sub txtacc2_LostFocus()
    If txtacc2 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc2) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc2 & " Not Found.", vbInformation, "Information"
        txtacc2 = ""
        txtacc2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc4.SetFocus
End Sub

Private Sub txtacc3_LostFocus()
    If txtacc3 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc3) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc3 & " Not Found.", vbInformation, "Information"
        txtacc3 = ""
        txtacc3.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc5.SetFocus
End Sub

Private Sub txtacc4_LostFocus()
    If txtacc4 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc4) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc4 & " Not Found.", vbInformation, "Information"
        txtacc4 = ""
        txtacc4.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc6.SetFocus
End Sub

Private Sub txtacc5_LostFocus()
    If txtacc5 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc5) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc5 & " Not Found.", vbInformation, "Information"
        txtacc5 = ""
        txtacc5.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc7.SetFocus
End Sub

Private Sub txtacc6_LostFocus()
    If txtacc6 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc6) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc6 & " Not Found.", vbInformation, "Information"
        txtacc6 = ""
        txtacc6.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtacc8.Enabled = True Then txtacc8.SetFocus
    If KeyAscii = 13 And txtacc8.Enabled = False Then cmdsave.SetFocus
End Sub

Private Sub txtacc7_LostFocus()
    If txtacc7 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc7) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc7 & " Not Found.", vbInformation, "Information"
        txtacc7 = ""
        txtacc7.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc9.SetFocus
End Sub

Private Sub txtacc8_LostFocus()
    If txtacc8 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc8) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc8 & " Not Found.", vbInformation, "Information"
        txtacc8 = ""
        txtacc8.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtacc9_LostFocus()
    If txtacc9 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc9) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc9 & " Not Found.", vbInformation, "Information"
        txtacc9 = ""
        txtacc9.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cektype()
    If setup2 = 1 Then
        If setup4 = "1" Then
            setup5 = "AS"
        ElseIf setup4 = "2" Then
            setup5 = "LI"
        ElseIf setup4 = "3" Then
            setup5 = "CA' or typeac = 'IS"
        ElseIf setup4 = "4" Then
            setup5 = "IS' or typeac = 'CA"
        End If
    ElseIf setup2 = 3 Then
        If setup4 = "1" Then
            setup5 = "AS"
        ElseIf setup4 = "2" Then
            setup5 = "LI"
        ElseIf setup4 = "3" Then
            setup5 = "CA"
        ElseIf setup4 = "4" Then
            setup5 = "IN"
        ElseIf setup4 = "5" Then
            setup5 = "EX"
        End If
    Else
        If setup4 = "1" Then
            setup5 = "IN"
        ElseIf setup4 = "2" Then
            setup5 = "EX"
        End If
    End If
End Sub
