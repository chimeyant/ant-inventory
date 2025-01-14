VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmitem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Barang Jadi"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   1680
   ClientWidth     =   7575
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
   Icon            =   "frmitem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "help"
      Height          =   3975
      Left            =   2760
      TabIndex        =   39
      Top             =   1800
      Visible         =   0   'False
      Width           =   7335
      Begin Chameleon.chameleonButton cmdclosehelp 
         Height          =   375
         Left            =   6360
         TabIndex        =   40
         Top             =   3410
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
         MICON           =   "frmitem.frx":2372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Akan menghapus rule level 4 saja, juga akan menghapus Item Master yang mempunyai level 1, 2, 3, dan 4 yang sama."
         Height          =   975
         Left            =   4200
         TabIndex        =   51
         Top             =   2400
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jika Level 4 sudah ada maka akan mengupdate keterangan"
         Height          =   975
         Left            =   1200
         TabIndex        =   50
         Top             =   2400
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Akan menghapus rule level 3 saja, juga akan menghapus Item Master yang mempunyai level 1, 2, dan 3 yang sama."
         Height          =   975
         Left            =   4200
         TabIndex        =   49
         Top             =   1440
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Akan menghapus rule level 2, 3, dan 4 juga akan menghapus Item Master yang mempunyai level 1 dan 2 yang sama."
         Height          =   975
         Left            =   4200
         TabIndex        =   48
         Top             =   480
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jika Level 3 sudah ada maka akan mengupdate keterangan"
         Height          =   975
         Left            =   1200
         TabIndex        =   47
         Top             =   1440
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jika Level 2 sudah ada maka akan mengupdate keterangan"
         Height          =   975
         Left            =   1200
         TabIndex        =   46
         Top             =   480
         Width           =   3000
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Level4"
         Height          =   975
         Left            =   120
         TabIndex        =   45
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Level3"
         Height          =   975
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Level2"
         Height          =   975
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   200
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox txt2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   200
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txt3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   200
      TabIndex        =   7
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txt4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   200
      TabIndex        =   11
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtnama 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   15
      Top             =   2400
      Width           =   5895
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   4560
      MaxLength       =   60
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtpo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   225
      Left            =   5520
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmitem.frx":268C
      Caption         =   "frmitem.frx":26AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmitem.frx":2718
      Keys            =   "frmitem.frx":2736
      Spin            =   "frmitem.frx":2778
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
      EditMode        =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtkode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
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
      Left            =   7200
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   1920
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
      Left            =   6480
      Picture         =   "frmitem.frx":27A0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
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
      Left            =   6840
      Picture         =   "frmitem.frx":2A82
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   5400
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
      MICON           =   "frmitem.frx":2DD0
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
      Left            =   5520
      TabIndex        =   23
      Top             =   5400
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
      MICON           =   "frmitem.frx":30EA
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
      Left            =   2520
      TabIndex        =   21
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save Barang Jadi"
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
      MICON           =   "frmitem.frx":3404
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
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Category"
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
      MICON           =   "frmitem.frx":371E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmd2 
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Top             =   600
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":3A38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmd3 
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":3D52
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmd4 
      Height          =   285
      Left            =   6360
      TabIndex        =   12
      Top             =   1320
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":406C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel2 
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Top             =   600
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":4386
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel3 
      Height          =   285
      Left            =   6720
      TabIndex        =   9
      Top             =   960
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":46A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel4 
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Top             =   1320
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmitem.frx":49BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update Nama"
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
      MICON           =   "frmitem.frx":4CD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdhelp 
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "help"
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
      MICON           =   "frmitem.frx":4FEE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadddetail 
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add Detail"
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
      MICON           =   "frmitem.frx":5308
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox l4 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   1320
      Width           =   855
      VariousPropertyBits=   746608667
      MaxLength       =   2
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
   Begin MSForms.ComboBox l3 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   855
      VariousPropertyBits=   746608667
      MaxLength       =   3
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
   Begin MSForms.ComboBox l2 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   855
      VariousPropertyBits=   746608667
      MaxLength       =   2
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
   Begin MSForms.ComboBox l1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
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
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level 4                       [2 char]"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   1350
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level 3                       [3 char]"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   990
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level 2                       [2 char]"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level 1"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   270
      Width           =   975
   End
   Begin VB.Label lblpo 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   2760
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label21 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5160
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Item"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   -120
      TabIndex        =   37
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow As String
Dim i As Integer
Dim tblprod, gcol1 As Boolean

Private Sub cmd2_Click()
    If MsgBox("Save rule Level 2 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub

    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 2 sudah ada, proses ini akan mengUPDATE keterangan level 2." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1 dan 2 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1 dan 2 sudah dipakai di Sales Order." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1 dan 2 sudah dipakai di Surat Jalan." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If

        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1 dan 2 sudah dipakai di Mutasi." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "update am_itemcode set ket = '" & txt2 & "' where lev = '2' and kode = '" & l1 & l2 & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    Else
        SQL = "insert into am_itemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"

        SQL = SQL + " values"
        SQL = SQL + " ('2'"
        SQL = SQL + " ,'" & l1 & l2 & "'"
        SQL = SQL + " ,'" & txt2 & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If

    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmd3_Click()
    If MsgBox("Save rule Level 3 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Or Len(Trim(l3)) < 3 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 3 sudah ada, proses ini akan mengUPDATE keterangan level 3." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2 dan 3 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2 dan 3 sudah dipakai di Sales Order." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2 dan 3 sudah dipakai di Surat Jalan." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If

        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2 dan 3 sudah dipakai di Mutasi." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "update am_itemcode set ket = '" & txt3 & "' where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    Else
        SQL = "insert into am_itemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('3'"
        SQL = SQL + " ,'" & l1 & l2 & l3 & "'"
        SQL = SQL + " ,'" & txt3 & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmd4_Click()
    If MsgBox("Save rule Level 4 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Or l4 = "" Or txt4 = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Or Len(Trim(l3)) < 3 Or Len(Trim(l4)) < 2 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    If l1 = "L" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l4 & "'"
    If l1 = "K" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "W" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "R" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 4 sudah ada, proses ini akan mengUPDATE keterangan level 4." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2, 3 dan 4 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & l3 & l4 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2, 3 dan 4 sudah dipakai di Sales Order." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2, 3 dan 4 sudah dipakai di Surat Jalan." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        
        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Rule Level 1, 2, 3 dan 4 sudah dipakai di Mutasi." & vbCrLf & _
            "click OK untuk LANJUTKAN update atau CANCEL untuk BATAL.", vbExclamation + vbOKCancel, "Information") = vbCancel Then
                OBJ.Close
                Exit Sub
            End If
        End If
        OBJ.Close
        
        OBJ.Open dsn
        If l1 = "L" Then SQL = "update am_itemcode set ket = '" & txt4 & "' where lev = '4' and kode = '" & l1 & l4 & "'"
        If l1 = "K" Then SQL = "update am_itemcode set ket = '" & txt4 & "' where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
        If l1 = "W" Then SQL = "update am_itemcode set ket = '" & txt4 & "' where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
        If l1 = "R" Then SQL = "update am_itemcode set ket = '" & txt4 & "' where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        hapusemua
    Else
        SQL = "insert into am_itemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('4'"
        If l1 = "L" Then SQL = SQL + " ,'" & l1 & l4 & "'"
        If l1 = "K" Then SQL = SQL + " ,'" & l1 & l2 & l4 & "'"
        If l1 = "W" Then SQL = SQL + " ,'" & l1 & l2 & l4 & "'"
        If l1 = "R" Then SQL = SQL + " ,'" & l1 & l2 & l4 & "'"
        SQL = SQL + " ,'" & txt4 & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdadd_Click()
    If txtKode = "" Or txtpo = "" Or l1 = "" Or l2 = "" Or l3 = "" Or l4 = "" Then
        MsgBox "Data entry not complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data entry not complete", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    txtKode = Trim(txtKode)
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 5) = "0.00" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    If l1 = "L" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l4 & "'"
    If l1 = "K" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "W" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "R" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 4, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_itemmst where kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Data already exist.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "insert into am_itemmst"
    SQL = SQL + " (kodebarang,"
    SQL = SQL + " namabarang,"
    SQL = SQL + " jenisbarang,"
    SQL = SQL + " kodeproduk,"
    SQL = SQL + " identry,"
    SQL = SQL + " dateentry,"
    SQL = SQL + " idupdate,"
    SQL = SQL + " dateupdate)"
    
    SQL = SQL + " values"
    SQL = SQL + " ('" & txtKode & "'"
    SQL = SQL + " ,'" & txtNama & "'"
    SQL = SQL + " ,'1'"
    SQL = SQL + " ,'" & txtpo & "'"
    SQL = SQL + " ,'" & kuser & "'"
    SQL = SQL + " ,convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + " ,' '"
    SQL = SQL + " ,convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "insert into am_itemdtl"
        SQL = SQL + " (kodebarang,"
        SQL = SQL + " namabarang,"
        SQL = SQL + " level_,"
        SQL = SQL + " kodesatuan,"
        SQL = SQL + " pricesale,"
        SQL = SQL + " hpprata2,"
        SQL = SQL + " konversi)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('" & txtKode & "'"
        SQL = SQL + " ,'" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 0), "general number") & "')"
        SQL = SQL + " ,'" & grid.TextMatrix(grid.Row, 1) & "'"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        SQL = SQL + " ,convert(money,'0')"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdadddetail_Click()
    If txtKode = "" Or txtpo = "" Or l1 = "" Or l2 = "" Or l3 = "" Or l4 = "" Then
        MsgBox "Data entry not complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data entry not complete", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    txtKode = Trim(txtKode)
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 5) = "0.00" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    If l1 = "L" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l4 & "'"
    If l1 = "K" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "W" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "R" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 4, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "delete from am_itemdtl where kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "insert into am_itemdtl"
        SQL = SQL + " (kodebarang,"
        SQL = SQL + " namabarang,"
        SQL = SQL + " level_,"
        SQL = SQL + " kodesatuan,"
        SQL = SQL + " pricesale,"
        SQL = SQL + " hpprata2,"
        SQL = SQL + " konversi)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('" & txtKode & "'"
        SQL = SQL + " ,'" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 0), "general number") & "')"
        SQL = SQL + " ,'" & grid.TextMatrix(grid.Row, 1) & "'"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "')"
        SQL = SQL + " ,convert(money,'0')"
        SQL = SQL + " ,convert(money,'" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    l1.Clear
    l1.ColumnCount = 2
    l1.ListWidth = "6 cm"
    l1.ColumnWidths = "2 cm; 4 cm"
    i = 0
    
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,1,1)'kode',b.ket from am_itemcode b where b.lev='1' order by substring(b.kode,1,1)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l1.AddItem RST!kode
        l1.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    l2.Clear
    l3.Clear
    l4.Clear
    l1 = ""
    l2 = ""
    l3 = ""
    l4 = ""
    txt1 = ""
    txt2 = ""
    txt3 = ""
    txt4 = ""
    
    txtKode = ""
    txtKode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdclosehelp_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdel2_Click()
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Proses ini akan mengHAPUS Rule Level 2, 3, dan Level 4." & vbCrLf & _
        "Proses ini juga akan mengHAPUS Item Master/Detail yang bersangkutan." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 1 dan 2 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 1 dan 2 sudah dipakai di Sales Order." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"

            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 1 dan 2 sudah dipakai di Surat Jalan." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 1 dan 2 sudah dipakai di Mutasi." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("HAPUS Rule Level 2, 3 dan 4 ?", vbQuestion + vbYesNo, "Question") = vbYes Then
            SQL = "delete from am_itemcode where lev = '2' and kode like '" & l1 & l2 & "%'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "delete from am_itemcode where lev = '3' and kode like '" & l1 & l2 & "%'"
            Set RST = OBJ.Execute(SQL)
            
            If l1 = "L" Then SQL = "select * from am_itemcode where lev = '4' and kode like '" & l1 & "%'"
            If l1 = "K" Then SQL = "delete from am_itemcode where lev = '4' and kode like '" & l1 & l2 & "%'"
            If l1 = "W" Then SQL = "delete from am_itemcode where lev = '4' and kode like '" & l1 & l2 & "%'"
            If l1 = "R" Then SQL = "delete from am_itemcode where lev = '4' and kode like '" & l1 & l2 & "%'"
            Set RST = OBJ.Execute(SQL)
            
            MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
        End If
        
        SQL = "delete from am_itemmst where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_itemdtl where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Master and Detail Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdel3_Click()
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Or Len(Trim(l3)) < 3 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Proses ini akan mengHAPUS Rule Level 3." & vbCrLf & _
        "Proses ini juga akan mengHAPUS Item Master/Detail yang bersangkutan." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 3 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 3 sudah dipakai di Sales Order." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"

            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 3 sudah dipakai di Surat Jalan." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 3 sudah dipakai di Mutasi." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("HAPUS Rule Level 3 ?", vbQuestion + vbYesNo, "Question") = vbYes Then
            SQL = "delete from am_itemcode where lev = '3' and kode like '" & l1 & l2 & l3 & "%'"
            Set RST = OBJ.Execute(SQL)
            
            MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
        End If
        
        SQL = "delete from am_itemmst where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_itemdtl where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Master and Detail Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdel4_Click()
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Or l4 = "" Or txt4 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 2 Or Len(Trim(l3)) < 3 Or Len(Trim(l4)) < 2 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_itemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    If l1 = "L" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l4 & "'"
    If l1 = "K" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "W" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    If l1 = "R" Then SQL = "select * from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Proses ini akan mengHAPUS Rule Level 4." & vbCrLf & _
        "Proses ini juga akan mengHAPUS Item Master/Detail yang bersangkutan." & vbCrLf & _
        "(Program sudah memberi PERINGATAN kepada User) Lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_soapp where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 4 sudah dipakai di Sales Order (app)." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_solin where kodebarang like '" & l1 & l2 & l3 & l4 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 4 sudah dipakai di Sales Order." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"

            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_sjapp where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 4 sudah dipakai di Surat Jalan." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_bpblin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Rule Level 4 sudah dipakai di Mutasi." & vbCrLf & _
            "Delete di BATAL kan !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("HAPUS Rule Level 4 ?", vbQuestion + vbYesNo, "Question") = vbYes Then
            If l1 = "L" Then SQL = "delete from am_itemcode where lev = '4' and kode = '" & l1 & l4 & "'"
            If l1 = "K" Then SQL = "delete from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
            If l1 = "W" Then SQL = "delete from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
            If l1 = "R" Then SQL = "delete from am_itemcode where lev = '4' and kode = '" & l1 & l2 & l4 & "'"
            Set RST = OBJ.Execute(SQL)
            
            MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
        End If
                
        SQL = "delete from am_itemmst where kodebarang = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_itemdtl where kodebarang = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Master and Detail Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdhelp_Click()
    Frame1.Visible = True
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodeproduk, namaproduk from am_produk"
    namatabel = "Product"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtpo = hasil
    lblpo = hasil1
    caristtp
    hasil = ""
    hasil1 = ""
    grid.SetFocus
End Sub

Private Sub cmdupdate_click()
    If txtKode = "" Or txtpo = "" Or l1 = "" Or l2 = "" Or l3 = "" Or l4 = "" Then
        MsgBox "Data entry not complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_itemmst where kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found, update aborted.", vbExclamation, "Information"
        OBJ.Close
        
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "UPDATE am_itemmst SET "
    SQL = SQL + "Namabarang = '" & txtNama & "'"
    SQL = SQL + "WHERE Kodebarang =  '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "update am_itemdtl set"
        SQL = SQL + " namabarang='" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL = SQL + " where kodebarang = '" & txtKode & "' and level_ =" & grid.TextMatrix(grid.Row, 0)
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from AM_soapp WHERE Kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_sjapp WHERE Kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_bpblin WHERE Kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "UPDATE am_itemmst SET "
    SQL = SQL + "kodeproduk = '" & txtpo & "'"
    SQL = SQL + "WHERE Kodebarang =  '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data updated, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
    
    Exit Sub
    
jump1:
    OBJ.Close
    MsgBox "Name updated, but can not update Category, data in use.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    If tblprod = True Then
        tblprod = False
        If hasil = "" Then Exit Sub
        txtpo = hasil
        lblpo = hasil1
        caristtp
        hasil = ""
        hasil1 = ""
        grid.SetFocus
        Exit Sub
    ElseIf gcol1 = True Then
        gcol1 = False
        If hasil = "" Then Exit Sub
    Select Case grid.Col
        Case 1
            grid.Row = 1
            Do While True
                If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                If grid.TextMatrix(grid.Row, 1) = hasil And posrow <> grid.Row Then
                    MsgBox "Kode Satuan already exist.", vbInformation, "Information"
                    hasil = ""
                    grid.Row = posrow
                    grid.Col = 0
                    Set grid.CellPicture = blank
                    grid.SetFocus
                    Exit Sub
                End If
                grid.Row = grid.Row + 1
            Loop
            
            grid.Row = posrow
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            
            OBJ.Open dsn
            SQL = "select namasatuan from am_unit where kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = txtNama
                grid.TextMatrix(grid.Row, 3) = "0.00"
                grid.TextMatrix(grid.Row, 4) = "0.00"
                grid.TextMatrix(grid.Row, 5) = "0.00"
                If grid.Row = 1 Then grid.TextMatrix(grid.Row, 5) = "1.00"
                
                lblitem = "    Nama Satuan : " & RST!namasatuan
                
                If grid.Row <> 1 Then grid.TextMatrix(grid.Row, 0) = grid.Row - 1
                If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                grid.SetFocus
                grid.Col = 2
            Else
                MsgBox "Satuan Not Found", vbExclamation, "Warning"
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = ""
                grid.TextMatrix(grid.Row, 5) = ""
            End If
            OBJ.Close
            Exit Sub
    End Select
    Exit Sub
    End If
    
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='31' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
    '
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='32' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        cmdupdate.Enabled = False
    '        cmdadddetail.Enabled = False
    '    End If
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='374' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        cmd2.Enabled = False
    '        cmd3.Enabled = False
    '        cmd4.Enabled = False
    '        cmdel2.Enabled = False
    '        cmdel3.Enabled = False
    '        cmdel4.Enabled = False
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    grid.TextMatrix(0, 0) = "Level"
    grid.TextMatrix(0, 1) = "Kode Satuan"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Price"
    grid.TextMatrix(0, 4) = "HPP"
    grid.TextMatrix(0, 5) = "Konversi"
    grid.TextMatrix(1, 0) = 0
    grid.ColWidth(0) = 550
    grid.ColWidth(1) = 1250
    grid.ColWidth(2) = 2250
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1500
    
    grid.RowHeightMin = 300
    
    setup2 = "add"
    
    l1.Clear
    l1.ColumnCount = 2
    l1.ListWidth = "6 cm"
    l1.ColumnWidths = "2 cm; 4 cm"
    i = 0
    
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,1,1)'kode',b.ket from am_itemcode b where b.lev='1' order by substring(b.kode,1,1)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l1.AddItem RST!kode
        l1.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "select * from am_unit where kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblitem = "    Nama Satuan : " & RST!namasatuan
        OBJ.Close
    End If
    If txtKode = "" Or txtpo = "" Then Exit Sub
    posrow = grid.Row
    Select Case grid.Col
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 7 Then
                MsgBox "Maximum 5 Level.", vbExclamation, "Warning"
                grid.Col = 0
                Exit Sub
            End If
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 60
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft + 30
            txtket.Top = grid.Top + grid.CellTop + 30
            txtket.MaxLength = 3
            txtket.Visible = True
            txtket.SetFocus
        Case 2
            If grid.TextMatrix(grid.Row, 1) = "" Or txtket.Visible = True Then Exit Sub
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 60
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft + 30
            txtket.Top = grid.Top + grid.CellTop + 30
            txtket.MaxLength = 30
            txtket.Visible = True
            txtket.SetFocus
        Case 3, 4, 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.Row = 1 And grid.Col = 5 Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtKode = "" Or txtpo = "" Then Exit Sub
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
        If grid.Rows - 1 = 7 Then Exit Sub
        If txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtket.Width = grid.ColWidth(grid.Col) - 60
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft + 30
        txtket.Top = grid.Top + grid.CellTop + 30
        txtket.MaxLength = 3
        txtket.Visible = True
        txtket.SetFocus
    Case 2
        If grid.TextMatrix(grid.Row, 1) = "" Or txtket.Visible = True Then Exit Sub
            
        posrow = grid.Row
        txtket.Width = grid.ColWidth(grid.Col) - 60
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft + 30
        txtket.Top = grid.Top + grid.CellTop + 30
        txtket.MaxLength = 30
        txtket.Visible = True
        txtket.SetFocus
    Case 3, 4, 5
        If grid.TextMatrix(grid.Row, 1) = "" Or txtnilai.Visible = True Then Exit Sub
        If grid.Row = 1 And grid.Col = 5 Then Exit Sub
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid.Col
        Case 1
            grid.Row = 1
            Do While True
                If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                If grid.TextMatrix(grid.Row, 1) = hasil And posrow <> grid.Row Then
                    MsgBox "Kode Satuan already exist.", vbInformation, "Information"
                    hasil = ""
                    grid.Row = posrow
                    grid.Col = 0
                    Set grid.CellPicture = blank
                    grid.SetFocus
                    Exit Sub
                End If
                grid.Row = grid.Row + 1
            Loop

            grid.Row = posrow
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            
            OBJ.Open dsn
            SQL = "select namasatuan from am_unit where kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = txtNama
                grid.TextMatrix(grid.Row, 3) = "0.00"
                grid.TextMatrix(grid.Row, 4) = "0.00"
                grid.TextMatrix(grid.Row, 5) = "0.00"
                If grid.Row = 1 Then grid.TextMatrix(grid.Row, 5) = "1.00"
                
                lblitem = "    Nama Satuan : " & RST!namasatuan
                
                If grid.Row <> 1 Then grid.TextMatrix(grid.Row, 0) = grid.Row - 1
                If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                grid.SetFocus
                grid.Col = 2
            Else
                MsgBox "Satuan Not Found", vbExclamation, "Warning"
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = ""
                grid.TextMatrix(grid.Row, 5) = ""
            End If
            OBJ.Close
    End Select
End Sub

Private Sub grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 And Y > 300 And X >= 0 And X <= 550 Then
    '    If grid.Rows - 1 = 7 Then
    '        frmainmenu.smnuinsert.Visible = False
    '    Else
    '        frmainmenu.smnuinsert.Visible = True
    '    End If
        
    '    frmainmenu.smnuremove.Caption = "Remove Level " & (Y \ 300) - 1
    '    frmainmenu.smnuinsert.Caption = "Insert Level " & (Y \ 300) - 1
    '    z = Y \ 300
    '    txtket.Visible = False
    '    txtnilai.Visible = False
    '    If z = 7 Then Exit Sub
    '    PopupMenu frmainmenu.mnupop
    'End If
End Sub

Private Sub l1_Change()
    l2.Clear
    l2.ColumnCount = 2
    l2.ListWidth = "6 cm"
    l2.ColumnWidths = "2 cm; 4 cm"
    l3.Clear
    l4.Clear
    txt2 = ""
    txt3 = ""
    txt4 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,2,2)'kode',b.ket from am_itemcode b where b.lev='2' and substring(b.kode,1,1)='" & l1 & "' order by substring(b.kode,2,2)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l2.AddItem RST!kode
        l2.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtKode = l1 & l2 & l3 & l4
End Sub

Private Sub l1_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_itemcode b where b.lev='1' and substring(b.kode,1,1)='" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt1 = RST!ket
    Else
        txt1 = ""
    End If
    OBJ.Close
End Sub

Private Sub l1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub l1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l2_Change()
    l3.Clear
    l3.ColumnCount = 2
    l3.ListWidth = "6 cm"
    l3.ColumnWidths = "2 cm; 4 cm"
    l4.Clear
    txt3 = ""
    txt4 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,4,3)'kode',b.ket from am_itemcode b where b.lev='3' and substring(b.kode,1,3)='" & l1 & l2 & "' order by substring(b.kode,4,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l3.AddItem RST!kode
        l3.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtKode = l1 & l2 & l3 & l4
End Sub

Private Sub l2_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_itemcode b where b.lev='2' and substring(b.kode,1,3)='" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt2 = RST!ket
    Else
        txt2 = ""
    End If
    OBJ.Close
End Sub

Private Sub l2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub l3_Change()
    l4.Clear
    l4.ColumnCount = 2
    l4.ListWidth = "6 cm"
    l4.ColumnWidths = "2 cm; 4 cm"
    txt4 = ""
    
    If l1 <> "" Then
        i = 0
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select distinct substring(b.kode,2,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' order by substring(b.kode,2,2)"
        If l1 = "K" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        If l1 = "W" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        If l1 = "R" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            l4.AddItem RST!kode
            l4.List(i, 1) = RST!ket
            i = i + 1
            RST.MoveNext
        Loop
        OBJ.Close
    End If
    
    txtKode = l1 & l2 & l3 & l4
End Sub

Private Sub l3_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_itemcode b where b.lev='3' and substring(b.kode,1,6)='" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt3 = RST!ket
    Else
        txt3 = ""
    End If
    OBJ.Close
End Sub

Private Sub l3_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub l4_Change()
    txtKode = l1 & l2 & l3 & l4
End Sub

Private Sub l4_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l4 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt4 = RST!ket
        Else
            txt4 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l4_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 1
                grid.Row = 1
                Do While True
                    If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                    If grid.TextMatrix(grid.Row, 1) = txtket And posrow <> grid.Row Then
                        MsgBox "Kode Satuan already exist.", vbInformation, "Information"
                        txtket = ""
                        grid.Row = posrow
                        grid.Col = 0
                        Set grid.CellPicture = blank
                        grid.SetFocus
                        Exit Sub
                    End If
                    grid.Row = grid.Row + 1
                Loop
                
                grid.Row = posrow
                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False
                
                OBJ.Open dsn
                SQL = "select namasatuan from am_unit where kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    grid.TextMatrix(grid.Row, 2) = txtNama
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    grid.TextMatrix(grid.Row, 4) = "0.00"
                    grid.TextMatrix(grid.Row, 5) = "0.00"
                    If grid.Row = 1 Then grid.TextMatrix(grid.Row, 5) = "1.00"
                    
                    lblitem = "    Nama Satuan : " & RST!namasatuan
                    
                    OBJ.Close
                    
                    If grid.Row <> 1 Then grid.TextMatrix(grid.Row, 0) = grid.Row - 1
                    grid.SetFocus
                    grid.Col = 2
    
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    
                    carisql1 = "select kodesatuan, namasatuan, init from am_unit"
                    namatabel = "Satuan"
                        
                    frmsearch.Show vbModal
                End If
                grid.Col = 1
            Case 2
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                
                grid.TextMatrix(posrow, 2) = txtket
                txtket = ""
                grid.SetFocus
                grid.Row = posrow
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        Select Case grid.Col
        Case 1
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtkode_Change()
    hapusemua
    
    OBJ.Open dsn
    SQL = "select * from am_itemmst where kodebarang = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtNama = RST!namabarang
        txtpo = RST!kodeproduk
        
       ' SQL = "select namaproduk from am_produk where kodeproduk = '" & txtpo & "'"
       ' Set RST = OBJ.Execute(SQL)
       ' If Not RST.EOF Then lblpo = RST!namaproduk
        
        grid.Row = 1
        SQL = "select * from am_itemdtl where kodebarang = '" & txtKode & "' order by level_"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            grid.TextMatrix(grid.Row, 0) = RST!level_
            grid.TextMatrix(grid.Row, 1) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = Format(RST!pricesale, "###,##0.00")
            grid.TextMatrix(grid.Row, 4) = "0.00"
            grid.TextMatrix(grid.Row, 5) = Format(RST!konversi, "###,##0.00")
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNama.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtpo.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtpo_LostFocus()
    caristtp
End Sub

Private Sub caristtp()
    If txtpo = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_produk where kodeproduk = '" & txtpo & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblpo = RST!namaproduk
    Else
        MsgBox "Category Not Found.", vbExclamation, "Warning"
        lblpo = ""
        
        cmdsearch2_Click
        If hasil <> "" Then
            txtpo = hasil
            SQL = "select * from am_produk where kodeproduk = '" & txtpo & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then lblpo = RST!namaproduk
            hasil = ""
            hasil1 = ""
            grid.SetFocus
        End If
    End If
    OBJ.Close
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    
    grid.Col = 0
    Set grid.CellPicture = blank
    lblitem = "    Nama Satuan : "
End Sub

Private Sub hapusemua()
    lblpo = ""
    txtpo = ""
    txtNama = ""
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 550
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 2000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1500
    lblitem = "    Nama Satuan : "
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
