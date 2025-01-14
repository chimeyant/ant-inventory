VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpenerimaanimport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   34
      Top             =   4680
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
      Left            =   5640
      Picture         =   "frmpenerimaanimport.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   33
      Top             =   4680
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
      Left            =   5400
      Picture         =   "frmpenerimaanimport.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtnobukti 
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
      Left            =   4800
      MaxLength       =   17
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtpo 
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtcurr 
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtsup 
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
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   4
      Top             =   480
      Width           =   5415
   End
   Begin VB.TextBox txtket 
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
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1560
      Width           =   5415
   End
   Begin VB.TextBox ket2 
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
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1920
      Width           =   5415
   End
   Begin VB.TextBox ket3 
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
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2460
      Width           =   5415
   End
   Begin VB.TextBox ket4 
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
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2190
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Overwrite (change)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2775
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         ItemData        =   "frmpenerimaanimport.frx":0630
         Left            =   120
         List            =   "frmpenerimaanimport.frx":0632
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpenerimaanimport.frx":0634
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Get File (*_terima.TRS) ~ on c:\"
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
      MICON           =   "frmpenerimaanimport.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdimport 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Import *_terima.trs"
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
      MICON           =   "frmpenerimaanimport.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdover 
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Overwrite"
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
      MICON           =   "frmpenerimaanimport.frx":0F82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   3240
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Format          =   106823683
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   3120
      TabIndex        =   13
      Top             =   2880
      Width           =   7215
      _ExtentX        =   12726
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
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "frmpenerimaanimport.frx":129C
      Caption         =   "frmpenerimaanimport.frx":12BC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanimport.frx":1328
      Keys            =   "frmpenerimaanimport.frx":1346
      Spin            =   "frmpenerimaanimport.frx":1388
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
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   285
      Left            =   8280
      TabIndex        =   8
      Top             =   1200
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmpenerimaanimport.frx":13B0
      Caption         =   "frmpenerimaanimport.frx":13D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaanimport.frx":143C
      Keys            =   "frmpenerimaanimport.frx":145A
      Spin            =   "frmpenerimaanimport.frx":149C
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
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   10
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
   Begin Chameleon.chameleonButton cmdclearr 
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpenerimaanimport.frx":14C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblket1 
      Height          =   255
      Left            =   3120
      TabIndex        =   36
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label lblket2 
      Height          =   255
      Left            =   3120
      TabIndex        =   35
      Top             =   5130
      Width           =   3495
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal LPB"
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
      Left            =   7080
      TabIndex        =   31
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Nomor P.O."
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
      Left            =   3240
      TabIndex        =   30
      Top             =   870
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Nomor LPB"
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
      Left            =   3240
      TabIndex        =   29
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Currency"
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
      Left            =   3240
      TabIndex        =   28
      Top             =   1230
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      Left            =   8280
      TabIndex        =   27
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier"
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
      Left            =   3240
      TabIndex        =   26
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Keterangan LPB"
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
      Left            =   3240
      TabIndex        =   25
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "PPn (%)"
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
      Left            =   7080
      TabIndex        =   24
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Keterangan PO"
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
      Left            =   3240
      TabIndex        =   23
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Import Data must execute on Server or Workstation that SQL Server is running."
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
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ready to Import !!"
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
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmpenerimaanimport"
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

Dim i, j, m As Integer

Private Sub cmdclear_Click()
    Label2 = Dir$("c:\*_terima.trs")
    If Label2 <> "" Then
        Label2 = "c:\" & Label2
        Label3 = "File Found." & vbCrLf & Label2
    Else
        Label3 = "File Not Found."
    End If
    cmdclear.Enabled = False
End Sub

Private Sub cmdclearr_Click()
    hapusgrid
    
    lblket1 = ""
    lblket2 = ""
    txtnobukti = ""
    Label6 = ""
    txtsup = ""
    txtpo = ""
    txtcurr = ""
    txtnilaikurs = 0
    txtppn = 0
    txtket = ""
    ket2 = ""
    ket3 = ""
    ket4 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdimport_Click()
On Error Resume Next
    If Label2 = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Label3 = "  please wait a moment ..."
    
    If MsgBox("Please make sure file exsist and valid." & vbCrLf & "Are you sure want to continue import file " & Label2 & " ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    Label3 = "Result :"
    
    '===============revisi update/delete penerimaan
    OBJ.Open dsn
    SQL = "SELECT distinct nobeli FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[revisi$] where flag1='0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "delete FROM am_belirev where nobeli='" & RST!nobeli & "' and flag1='0' and flag2='0'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        RST.MoveNext
        DoEvents
    Loop
    
    SQL = "SELECT nobeli,tglbeli,nopo,ref1,ref2,kodesupp,kodecur,nilaikurs,kodebarang,qty,price,kodesatuan,lineitem,flag1,flag2,ppn,keterangan,keterangan2,keterangan3,keterangan4 FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[revisi$]"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        date1 = RST!tglbeli
        
        OBJ1.Open dsn
        SQL1 = "INSERT INTO AM_belirev"
        SQL1 = SQL1 + " (noBeli"
        SQL1 = SQL1 + ", TglBeli"
        SQL1 = SQL1 + ", nopo"
        SQL1 = SQL1 + ", ref1"
        SQL1 = SQL1 + ", ref2"
        SQL1 = SQL1 + ", kodesupp"
        SQL1 = SQL1 + ", kodecur"
        SQL1 = SQL1 + ", nilaikurs"
        SQL1 = SQL1 + ", Kodebarang"
        SQL1 = SQL1 + ", qty"
        SQL1 = SQL1 + ", Price"
        SQL1 = SQL1 + ", kodesatuan"
        SQL1 = SQL1 + ", keterangan"
        SQL1 = SQL1 + ", keterangan2"
        SQL1 = SQL1 + ", keterangan3"
        SQL1 = SQL1 + ", keterangan4"
        SQL1 = SQL1 + ", ppn"
        SQL1 = SQL1 + ", lineitem"
        SQL1 = SQL1 + ", flag1"   '0 untuk update 1 untuk delete
        SQL1 = SQL1 + ", flag2)"
    
        SQL1 = SQL1 + " VALUES"
        SQL1 = SQL1 + " ('" & RST!nobeli & "'"
        SQL1 = SQL1 + ",Convert(dateTime, '" & tanggal1 & "')"
        SQL1 = SQL1 + ", '" & RST!nopo & "'"
        SQL1 = SQL1 + ", '" & RST!ref1 & "'"
        SQL1 = SQL1 + ", '" & RST!ref2 & "'"
        SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
        SQL1 = SQL1 + ", '" & RST!kodecur & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
        SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
        SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan2 & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan3 & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan4 & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!ppn & "')"
        SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
        SQL1 = SQL1 + ", '" & RST!flag1 & "'"
        SQL1 = SQL1 + ", '" & RST!flag2 & "')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
    
        RST.MoveNext
        DoEvents
    Loop
    
    SQL = "SELECT distinct nobeli FROM am_belirev where flag1='1' and flag2='0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * FROM am_beliapp where nobeli = '" & RST!nobeli & "' and flag1='1' and flag2='0'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            SQL1 = "delete FROM am_beliapp where nobeli = '" & RST!nobeli & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            
            SQL1 = "update AM_belirev set flag2 = '1' where nobeli = '" & RST!nobeli & "' and flag1 = '1'"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
    
        RST.MoveNext
        DoEvents
    Loop
    OBJ.Close
    '========================================
    
    '===============penerimaan
    OBJ.Open dsn
    SQL = "SELECT nobeli,tglbeli,nopo,ref1,ref2,kodesupp,kodecur,nilaikurs,kodebarang,qty,price,kodesatuan,lineitem,flag1,flag2,ppn,keterangan,keterangan2,keterangan3,keterangan4 FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[terima$]"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_beliapp where nobeli = '" & RST!nobeli & "' and kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_beliapp"
            SQL1 = SQL1 + " (NoBeli"
            SQL1 = SQL1 + ", TglBeli"
            SQL1 = SQL1 + ", nopo"
            SQL1 = SQL1 + ", ref1"
            SQL1 = SQL1 + ", ref2"
            SQL1 = SQL1 + ", kodesupp"
            SQL1 = SQL1 + ", kodecur"
            SQL1 = SQL1 + ", nilaikurs"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", Price"
            SQL1 = SQL1 + ", kodesatuan"
            SQL1 = SQL1 + ", keterangan"
            SQL1 = SQL1 + ", keterangan2"
            SQL1 = SQL1 + ", keterangan3"
            SQL1 = SQL1 + ", keterangan4"
            SQL1 = SQL1 + ", ppn"
            SQL1 = SQL1 + ", lineitem"
            SQL1 = SQL1 + ", flag1"
            SQL1 = SQL1 + ", flag2)"
        
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + " ('" & RST!nobeli & "'"
            SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglbeli) & "/" & Day(RST!tglbeli) & "/" & Year(RST!tglbeli) & "')"
            SQL1 = SQL1 + ", '" & RST!nopo & "'"
            SQL1 = SQL1 + ", '" & RST!ref1 & "'"
            SQL1 = SQL1 + ", '" & RST!ref2 & "'"
            SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
            SQL1 = SQL1 + ", '" & RST!kodecur & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
            SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
            SQL1 = SQL1 + ", '" & RST!keterangan & "'"
            SQL1 = SQL1 + ", '" & RST!keterangan2 & "'"
            SQL1 = SQL1 + ", '" & RST!keterangan3 & "'"
            SQL1 = SQL1 + ", '" & RST!keterangan4 & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!ppn & "')"
            SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
            SQL1 = SQL1 + ", '0'"
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        RST.MoveNext
    Loop
    
    SQL = "select count(nobeli)'totalbaris' from AM_beliapp where flag1 = '0'"
    Set RST = OBJ.Execute(SQL)
    i = RST!totalbaris
    Label3 = Label3 + vbCrLf + "   Import PENERIMAAN, " & i & " records affected."
    OBJ.Close
    
    '===============retur
    OBJ.Open dsn
    SQL = "SELECT noretur,tglretur,nobeli,nopo,kodesupp,kodecur,nilaikurs,keterangan,kodebarang,qty,price,kodesatuan,lineitem,flag1,flag2 FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[retur$]"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_beliretur where noretur = '" & RST!noretur & "' and kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_beliretur"
            SQL1 = SQL1 + " (Noretur"
            SQL1 = SQL1 + ", Tglretur"
            SQL1 = SQL1 + ", nobeli"
            SQL1 = SQL1 + ", nopo"
            SQL1 = SQL1 + ", kodesupp"
            SQL1 = SQL1 + ", kodecur"
            SQL1 = SQL1 + ", nilaikurs"
            SQL1 = SQL1 + ", keterangan"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", Price"
            SQL1 = SQL1 + ", kodesatuan"
            SQL1 = SQL1 + ", lineitem"
            SQL1 = SQL1 + ", flag1"
            SQL1 = SQL1 + ", flag2)"
        
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + " ('" & RST!noretur & "'"
            SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglretur) & "/" & Day(RST!tglretur) & "/" & Year(RST!tglretur) & "')"
            SQL1 = SQL1 + ", '" & RST!nobeli & "'"
            SQL1 = SQL1 + ", '" & RST!nopo & "'"
            SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
            SQL1 = SQL1 + ", '" & RST!kodecur & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
            SQL1 = SQL1 + ", '" & RST!keterangan & "'"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
            SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
            SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
            SQL1 = SQL1 + ", '0'"
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        RST.MoveNext
    Loop
    SQL = "select count(noretur)'totalbaris' from AM_beliretur where flag1 = '0'"
    Set RST = OBJ.Execute(SQL)
    m = RST!totalbaris
    Label3 = Label3 + vbCrLf + "   Import RETUR, " & m & " records affected."
    
    '===============retur temporary
    SQL = "SELECT noretur,nobeli,kodebarang,qty,qtyuse FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[returtemp$]"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_belireturtemp where noretur = '" & RST!noretur & "' and kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_belireturtemp"
            SQL1 = SQL1 + " (Noretur"
            SQL1 = SQL1 + ", nobeli"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", qtyuse)"
        
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + " ('" & RST!noretur & "'"
            SQL1 = SQL1 + ", '" & RST!nobeli & "'"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qtyuse & "'))"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
    '===============supplier
    OBJ.Open dsn
    SQL = "Select count(kodesupp)'jum' from am_supplier"
    Set RST = OBJ.Execute(SQL)
    j = RST!jum

    SQL = "SELECT kodesupp,namasupp,alamatsupp1,alamatsupp2,telpsupp,faxsupp,contactperson,Category,Wp FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[supplier$]"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_supplier where kodesupp = '" & RST!kodesupp & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_supplier"
            SQL1 = SQL1 + " (kodesupp"
            SQL1 = SQL1 + ", namasupp"
            SQL1 = SQL1 + ", alamatsupp1"
            SQL1 = SQL1 + ", alamatsupp2"
            SQL1 = SQL1 + ", telpsupp"
            SQL1 = SQL1 + ", faxsupp"
            SQL1 = SQL1 + ", category"
            SQL1 = SQL1 + ", wp"
            SQL1 = SQL1 + ", contactperson)"
        
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + " ('" & RST!kodesupp & "'"
            SQL1 = SQL1 + ", '" & RST!namasupp & "'"
            SQL1 = SQL1 + ", '" & RST!alamatsupp1 & "'"
            SQL1 = SQL1 + ", '" & RST!alamatsupp2 & "'"
            SQL1 = SQL1 + ", '" & RST!telpsupp & "'"
            SQL1 = SQL1 + ", '" & RST!faxsupp & "'"
            SQL1 = SQL1 + ", '" & RST!Category & "'"
            SQL1 = SQL1 + ", '" & RST!wp & "'"
            SQL1 = SQL1 + ", '" & RST!contactperson & "')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
    '===============unit
    j = 0
    OBJ.Open dsn
    SQL = "SELECT kodesatuan,namasatuan,initial FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[unit$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_apunit where kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_apunit"
                SQL1 = SQL1 + " (KodeSatuan"
                SQL1 = SQL1 + ", NamaSatuan"
                SQL1 = SQL1 + ", initial)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", '" & RST!namasatuan & "'"
                SQL1 = SQL1 + ", '" & RST!initial & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                j = j + 1
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    '===============item
    j = 0
    OBJ.Open dsn
    SQL = "SELECT kodebarang,namabarang,kodesatuan,kodeproduk,kodesatuanmutasi FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[item$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_apitemmst where kodebarang = '" & RST!kodebarang & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_apitemmst"
                SQL1 = SQL1 + " (Kodebarang"
                SQL1 = SQL1 + ", namabarang"
                SQL1 = SQL1 + ", kodesatuan"
                SQL1 = SQL1 + ", kodesatuanmutasi"
                SQL1 = SQL1 + ", kodeproduk)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodebarang & "'"
                SQL1 = SQL1 + ", '" & RST!namabarang & "'"
                SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", '" & RST!kodesatuanmutasi & "'"
                SQL1 = SQL1 + ", '" & RST!kodeproduk & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                j = j + 1
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    '===============item rule
    j = 0
    OBJ.Open dsn
    SQL = "SELECT lev,kode,ket FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[rule$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_apitemcode where lev = '" & RST!lev & "' and kode = '" & RST!kode & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_apitemcode"
                SQL1 = SQL1 + " (lev"
                SQL1 = SQL1 + ", kode"
                SQL1 = SQL1 + ", ket)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!lev & "'"
                SQL1 = SQL1 + ", '" & RST!kode & "'"
                SQL1 = SQL1 + ", '" & RST!ket & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                j = j + 1
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    '===============kode PO
    OBJ.Open dsn
    SQL = "SELECT kode1,kode2,kode3 FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[divisi$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        OBJ1.Open dsn
        SQL1 = "delete from AM_kode"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
            
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "INSERT INTO AM_kode"
            SQL1 = SQL1 + " (Kode1"
            SQL1 = SQL1 + ", kode2"
            SQL1 = SQL1 + ", kode3)"
            
            SQL1 = SQL1 + "VALUES"
            SQL1 = SQL1 + " ('" & RST!kode1 & "'"
            SQL1 = SQL1 + ", '" & RST!kode2 & "'"
            SQL1 = SQL1 + ", '" & RST!kode3 & "')"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    
    Kill Label2
    
    OBJ.Open dsn
    SQL = "update AM_beliapp set flag1 = '1' where flag1 = '0'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "update AM_beliretur set flag1 = '1' where flag1 = '0'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Import Complete.", vbInformation, "Information"
    cmdimport.Enabled = False
    
    List1.Clear
    OBJ.Open dsn
    SQL = "select distinct nobeli from AM_belirev where flag1 = '0' and flag2 = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nobeli
        
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdover_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtpo = "" Or txtcurr = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtppn > 0 And txtppn < 10 Then
        MsgBox "PPn Value must 10.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If MsgBox("Are you sure want to overwrite (Penerimaan Barang)?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from AM_apopnfil where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "To overwrite, user must unconfirm " & txtnobukti & " (BPB)", vbExclamation, "Information"

        OBJ.Close
        Exit Sub
    End If
    
    SQL = "delete from AM_beliapp where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT * FROM am_belirev WHERE nobeli = '" & txtnobukti & "' and flag1 = '0' and flag2 = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "INSERT INTO AM_beliapp"
        SQL1 = SQL1 + " (NoBeli"
        SQL1 = SQL1 + ", TglBeli"
        SQL1 = SQL1 + ", nopo"
        SQL1 = SQL1 + ", ref1"
        SQL1 = SQL1 + ", ref2"
        SQL1 = SQL1 + ", kodesupp"
        SQL1 = SQL1 + ", kodecur"
        SQL1 = SQL1 + ", nilaikurs"
        SQL1 = SQL1 + ", Kodebarang"
        SQL1 = SQL1 + ", qty"
        SQL1 = SQL1 + ", Price"
        SQL1 = SQL1 + ", kodesatuan"
        SQL1 = SQL1 + ", keterangan"
        SQL1 = SQL1 + ", keterangan2"
        SQL1 = SQL1 + ", keterangan3"
        SQL1 = SQL1 + ", keterangan4"
        SQL1 = SQL1 + ", ppn"
        SQL1 = SQL1 + ", lineitem"
        SQL1 = SQL1 + ", flag1"
        SQL1 = SQL1 + ", flag2)"

        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + " ('" & RST!nobeli & "'"
        SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglbeli) & "/" & Day(RST!tglbeli) & "/" & Year(RST!tglbeli) & "')"
        SQL1 = SQL1 + ", '" & RST!nopo & "'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
        SQL1 = SQL1 + ", '" & RST!kodecur & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
        SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
        SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan2 & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan3 & "'"
        SQL1 = SQL1 + ", '" & RST!keterangan4 & "'"
        SQL1 = SQL1 + ",Convert (Money, '" & RST!ppn & "')"
        SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
        SQL1 = SQL1 + ", '1'"
        SQL1 = SQL1 + ", '0')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    List1.RemoveItem (i)
    
    OBJ.Open dsn
    SQL = "update AM_belirev set flag2 = '1' where nobeli = '" & txtnobukti & "' and flag1 = '0'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Overwrite Complete.", vbInformation, "Information"
    cmdclearr_Click
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='94' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "K/Sat."
    grid.TextMatrix(0, 3) = "Qty"
    grid.TextMatrix(0, 4) = "Price"
    grid.TextMatrix(0, 5) = "Jumlah"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 1500
    
    grid.RowHeightMin = 300
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    OBJ1.Open dsn
    SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket1 = "Nama Barang : " & RST1!namabarang
    If RST1.EOF Then lblket1 = "Nama Barang : "

    SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 2) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket2 = "Nama Satuan : " & RST1!namasatuan
    If RST1.EOF Then lblket2 = "Nama Satuan : "
    OBJ1.Close
End Sub

Private Sub ket2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket4.SetFocus
    KeyAscii = 0
End Sub

Private Sub ket4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket3.SetFocus
    KeyAscii = 0
End Sub

Private Sub List1_DblClick()
    If List1.text = "" Then Exit Sub

    hapusgrid
    txtnobukti = ""
    Label6 = ""
    txtsup = ""
    txtpo = ""
    txtcurr = ""
    txtnilaikurs = 0
    txtppn = 0
    txtket = ""
    ket2 = ""
    ket3 = ""
    ket4 = ""
    lblket1 = ""
    lblket2 = ""
    
    txtnobukti = List1.text
    i = List1.ListIndex

    OBJ.Open dsn
    SQL = "select distinct * from am_belirev where nobeli = '" & txtnobukti & "' and flag1='0'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbeli
        Label6 = Format(RST!tglbeli, "dd MMMM yyyy")
        txtpo = RST!nopo
        txtcurr = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtppn = RST!ppn
        txtket = RST!keterangan
        ket2 = RST!keterangan2
        ket3 = RST!keterangan3
        ket4 = RST!keterangan4
        
        OBJ1.Open dsn
        SQL1 = "select * from am_supplier where kodesupp = '" & RST!kodesupp & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then txtsup = RST1!namasupp Else txtsup = ""
        OBJ1.Close
        
        grid.Row = 1
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.Col = 2
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 2) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 4) = Format(RST!Price, "###,###,##0.0000")
            grid.TextMatrix(grid.Row, 5) = Format(RST!qty * RST!Price, "###,###,##0.0000")
            
            grid.Col = 0
            Set grid.CellPicture = uncheck.Picture

            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub txtcurr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ket2.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtppn.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsup.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtcurr.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtppn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtket.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtpo.SetFocus
    KeyAscii = 0
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Private Sub hapusgrid()
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
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 1500
End Sub
