VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmpayaredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Payment"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpayaredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3750
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   240
         TabIndex        =   56
         Top             =   960
         Width           =   8655
      End
      Begin VB.TextBox txtcari1 
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
         Left            =   2040
         TabIndex        =   40
         Top             =   1725
         Width           =   1575
      End
      Begin VB.TextBox txtcari 
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
         Left            =   2040
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin Chameleon.chameleonButton cmdcari 
         Height          =   285
         Left            =   3840
         TabIndex        =   39
         Top             =   1320
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Browse Cek or Giro"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmpayaredit.frx":2372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdcari1 
         Height          =   285
         Left            =   3840
         TabIndex        =   41
         Top             =   1725
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Browse Faktur"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmpayaredit.frx":268C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdclose1 
         Height          =   375
         Left            =   7200
         TabIndex        =   54
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Close search by ..."
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmpayaredit.frx":29A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         Caption         =   "User just type the number and press enter on text box or just click the command button to browsing No. Giro/Cek or No. Faktur."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   55
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label Label9 
         Caption         =   "for Search by No. Faktur, user must set character "">"" on the grid header to noapply not nobkt. Just click on grid header."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   53
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label8 
         Caption         =   "User can search Payment by No. Giro/Cek or by No. Faktur."
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
         TabIndex        =   44
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Search by No. Faktur"
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
         TabIndex        =   43
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Search by No. Cek/Giro"
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
         TabIndex        =   42
         Top             =   1350
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   300
      Left            =   8280
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   975
         VariousPropertyBits=   612386843
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1720;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin Chameleon.chameleonButton cmdswitch 
      Height          =   390
      Left            =   7920
      TabIndex        =   37
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "search by ..."
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayaredit.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtketerangan 
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
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1560
      Width           =   6255
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
      Left            =   7800
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
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
      Left            =   7560
      Picture         =   "frmpayaredit.frx":2FDA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
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
      Left            =   7320
      Picture         =   "frmpayaredit.frx":3390
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtkurs 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   6360
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Format          =   143851521
      CurrentDate     =   38515
   End
   Begin VB.TextBox txtkodecol 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin TDBText6Ctl.TDBText txtbukti 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmpayaredit.frx":3746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":37B2
      Key             =   "frmpayaredit.frx":37D0
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Format          =   143851523
      CurrentDate     =   37421
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":380C
      Caption         =   "frmpayaredit.frx":382C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":3898
      Keys            =   "frmpayaredit.frx":38B6
      Spin            =   "frmpayaredit.frx":38F8
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Bukti"
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
      MICON           =   "frmpayaredit.frx":3920
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
      Left            =   8280
      TabIndex        =   16
      Top             =   6120
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayaredit.frx":3C3A
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
      Left            =   7320
      TabIndex        =   15
      Top             =   6120
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayaredit.frx":3F54
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel 
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   6120
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayaredit.frx":426E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Collector"
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
      MICON           =   "frmpayaredit.frx":4588
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   6120
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpayaredit.frx":48A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmpayaredit.frx":4BBC
      Caption         =   "frmpayaredit.frx":4BDC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":4C48
      Keys            =   "frmpayaredit.frx":4C66
      Spin            =   "frmpayaredit.frx":4CA8
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "frmpayaredit.frx":4CD0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":4D3C
      Key             =   "frmpayaredit.frx":4D5A
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
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
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
   Begin MSComCtl2.DTPicker date3 
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   143851523
      CurrentDate     =   38767
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":4D96
      Caption         =   "frmpayaredit.frx":4DB6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":4E22
      Keys            =   "frmpayaredit.frx":4E40
      Spin            =   "frmpayaredit.frx":4E82
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   9
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
      _Band(0).Cols   =   9
   End
   Begin TDBNumber6Ctl.TDBNumber txtsisa 
      Height          =   255
      Left            =   6360
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":4EAA
      Caption         =   "frmpayaredit.frx":4ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":4F36
      Keys            =   "frmpayaredit.frx":4F54
      Spin            =   "frmpayaredit.frx":4F96
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai2 
      Height          =   255
      Left            =   7320
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":4FBE
      Caption         =   "frmpayaredit.frx":4FDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":504A
      Keys            =   "frmpayaredit.frx":5068
      Spin            =   "frmpayaredit.frx":50AA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai3 
      Height          =   255
      Left            =   7320
      TabIndex        =   49
      Top             =   240
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":50D2
      Caption         =   "frmpayaredit.frx":50F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":515E
      Keys            =   "frmpayaredit.frx":517C
      Spin            =   "frmpayaredit.frx":51BE
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai4 
      Height          =   255
      Left            =   7320
      TabIndex        =   50
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":51E6
      Caption         =   "frmpayaredit.frx":5206
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":5272
      Keys            =   "frmpayaredit.frx":5290
      Spin            =   "frmpayaredit.frx":52D2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai5 
      Height          =   255
      Left            =   6840
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      Calculator      =   "frmpayaredit.frx":52FA
      Caption         =   "frmpayaredit.frx":531A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpayaredit.frx":5386
      Keys            =   "frmpayaredit.frx":53A4
      Spin            =   "frmpayaredit.frx":53E6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSForms.Label lblflag 
      Height          =   735
      Left            =   7920
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      BackColor       =   12640511
      Caption         =   "Pembayaran dengan base currency"
      Size            =   "2355;1296"
      SpecialEffect   =   6
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblbayar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Bayar Apply : 0.00"
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
      TabIndex        =   45
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Bayar : 0.00"
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
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label lblsisa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Sisa : 0.00"
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
      TabIndex        =   32
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label lblapply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Apply : 0.00"
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
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   27
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
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
      Left            =   3480
      TabIndex        =   26
      Top             =   510
      Width           =   855
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblnamacol 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      TabIndex        =   23
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
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
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Bukti"
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
      Left            =   3225
      TabIndex        =   18
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblsup 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      TabIndex        =   17
      Top             =   825
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   -120
      TabIndex        =   35
      Top             =   6030
      Width           =   12135
   End
End
Attribute VB_Name = "frmpayaredit"
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

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim str2, posrow As String
Dim i As Integer

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Function tanggalgrid1()
    If Grid1.TextMatrix(Grid1.Row, 7) = "" Then
        tanggalgrid1 = "01/01/1900"
    Else
        tanggalgrid1 = Month(Grid1.TextMatrix(Grid1.Row, 7)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 7)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 7))
    End If
End Function

Function tanggalgrid2()
    If Grid1.TextMatrix(Grid1.Row, 8) = "" Then
        tanggalgrid2 = "01/01/1900"
    Else
        tanggalgrid2 = Month(Grid1.TextMatrix(Grid1.Row, 8)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 8)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 8))
    End If
End Function

Private Sub cmbtype_Click()
    If cmbtype = "" Then Exit Sub
    
    Grid1.Row = 1
    Do While cmbtype = "Tunai"
        If Grid1.Row = Grid1.Rows - 1 Then Exit Do
        
        If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then
            cmbtype = ""
            cmbtype.Visible = False
            Exit Sub
        End If
        
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Row = posrow
    
    Grid1.SetFocus
    Grid1.TextMatrix(Grid1.Row, 1) = cmbtype
    Grid1.TextMatrix(Grid1.Row, 6) = "0.00"
    cmbtype = ""
    Frame2.Visible = False
    
    Grid1.Col = 0
    Set Grid1.CellPicture = uncheck.Picture
                        
    If Grid1.Row = (Grid1.Rows - 1) Then Grid1.Rows = Grid1.Rows + 1
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then cmbtype_LostFocus
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    Frame2.Visible = False
End Sub

Private Sub cmdcari_Click()
    carisql1 = "select nobkt,nogiro from AM_cashsub where nogiro <> ''"
    namatabel = "Cek/Giro"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    hasil = ""
    Frame1.Visible = False
    
    Cariar
End Sub

Private Sub cmdcari1_Click()
    carisql1 = "select nobkt,noapply from AM_cashlin where kodebayar = 'PM'"
    namatabel = "Faktur"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari1_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    txtcari1 = hasil1
    Frame1.Visible = False
    hasil = ""
    hasil1 = ""
    
    Cariar
End Sub

Private Sub cmdclose1_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If grid2.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar='PM' and idupdate='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, data already process.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ1.Open dsn
    SQL1 = "select * from gl_transaksi where notrx = '" & txtbukti & "' and identry='auto'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close
        
        MsgBox "Data tidak bisa dihapus, transaksi pembayaran ada di transaksi GL.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ1.Close
        
    OBJ.Open dsn
    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashsub where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_aropnfil where nobkt = '" & txtbukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kode, nama, idupdate from AM_collector"
    namatabel = "Collector"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecol = hasil
    lblnamacol = hasil1
    caricollector
    hasil = ""
End Sub

Private Sub cmdswitch_Click()
    Frame1.Visible = True
    txtcari = ""
    txtcari1 = ""
    txtcari.SetFocus
    Frame1.Left = 120
    Frame1.Top = 60
End Sub

Private Sub date1_Change()
    date1 = date2
End Sub

Private Sub date3_CloseUp()
    Grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")
    
    Grid1.SetFocus
    Grid1.Row = posrow
    date3.Visible = False
End Sub

Private Sub date3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date3.Visible = False
    If KeyCode = 13 Then
        Grid1.TextMatrix(posrow, 3) = Format(date3, "dd/MM/yyyy")
        
        Grid1.SetFocus
        Grid1.Row = posrow
        date3.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    
    cmbtype.AddItem "Tunai"
    'cmbtype.AddItem "Cek"
    cmbtype.AddItem "Giro"
    cmbtype.AddItem "Transfer"
        
    grid2.TextMatrix(0, 0) = "No Apply"
    grid2.TextMatrix(0, 1) = "Piutang"
    grid2.TextMatrix(0, 2) = "Nilai Bayar"
    grid2.TextMatrix(0, 3) = "Disc Bayar"
    grid2.TextMatrix(0, 4) = "Selisih"
    grid2.TextMatrix(0, 5) = "Selisih Kurs"
    grid2.TextMatrix(0, 6) = "Sisa Piutang"
    
    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300
    grid2.ColWidth(6) = 1300
    
    Grid1.TextMatrix(0, 1) = "Type Bayar"
    Grid1.TextMatrix(0, 2) = "No Cek/Giro"
    Grid1.TextMatrix(0, 3) = "J/T - Trans"
    Grid1.TextMatrix(0, 4) = "Bank"
    Grid1.TextMatrix(0, 5) = "Acc Sparta"
    Grid1.TextMatrix(0, 6) = "Nilai"
    Grid1.TextMatrix(0, 7) = "Cair"
    Grid1.TextMatrix(0, 8) = "Tolak"
    
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 0
    Grid1.ColWidth(8) = 0
    
    Grid1.RowHeightMin = 300
    grid2.RowHeightMin = 300
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalgrid()
    If Grid1.TextMatrix(Grid1.Row, 3) = "" Then
        tanggalgrid = "01/01/1900"
    Else
        tanggalgrid = Month(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 3)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 3))
    End If
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdclear_Click()
    txtbukti = ""
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtkodecol = ""
    lblnamacol = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtsisa = 0
    txtketerangan = ""
    hapusgrid
    hapusgrid1
    lblflag.Visible = False
    cmdsearch.Enabled = True
    txtbukti.Enabled = True
    txtbukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub hapusgrid()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.ColWidth(0) = 1150
    grid2.ColWidth(1) = 1300
    grid2.ColWidth(2) = 1300
    grid2.ColWidth(3) = 1300
    grid2.ColWidth(4) = 1300
    grid2.ColWidth(5) = 1300
    grid2.ColWidth(6) = 1300
    
    lblapply = "Total Apply : 0.00"
    lblbayar = "Bayar Apply : 0.00"
End Sub

Private Sub hapusgrid1()
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        Grid1.TextMatrix(Grid1.Row, 1) = ""
        Grid1.TextMatrix(Grid1.Row, 2) = ""
        Grid1.TextMatrix(Grid1.Row, 3) = ""
        Grid1.TextMatrix(Grid1.Row, 4) = ""
        Grid1.TextMatrix(Grid1.Row, 5) = ""
        Grid1.TextMatrix(Grid1.Row, 6) = ""
        Grid1.TextMatrix(Grid1.Row, 7) = ""
        Grid1.TextMatrix(Grid1.Row, 8) = ""
        Grid1.Col = 0
        Set Grid1.CellPicture = blank
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = 2
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 0
    Grid1.ColWidth(8) = 0
    
    lbltotal = "Total Bayar : 0.00"
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_cashhdr where kodebayar = 'PM' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_cashhdr where kodebayar = 'PM'"
    End If
    namatabel = "Pembayaran"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    hasil = ""
    Cariar
End Sub

Private Sub Cariar()
    If txtbukti = "" Then Exit Sub
    
    hapusgrid
    hapusgrid1
    txtsup = ""
    lblsup = ""
    txtkodecol = ""
    lblnamacol = ""
    txtkurs = ""
    txtketerangan = ""
    txtnilaikurs = 0
    date1 = Date
    lblflag.Visible = False
    
    OBJ.Open dsn
    SQL = "Select * From AM_CashHdr Where NoBkt = '" & txtbukti & "' And kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglbkt
        date2 = RST!tglbkt
        txtsup = RST!kodecust
        txtkodecol = RST!kodecol
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        txtketerangan = RST!keterangan
        If RST!posted = "1" Then lblflag.Visible = True
        If RST!posted = "0" Then lblflag.Visible = False
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
        Set RST = OBJ.Execute(SQL)
        If RST!base = 1 Then lblbase = "1" Else lblbase = "0"
        
        SQL = "Select * From AM_customer Where kodecust = '" & txtsup & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsup = RST!namacust
        Else
            lblsup = ""
        End If
        
        SQL = "Select * From am_collector Where kode = '" & txtkodecol & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblnamacol = RST!nama
        Else
            lblnamacol = ""
        End If
        
        txtbukti.Enabled = False
        cmdsearch.Enabled = False
        'keluarkan record dari cashlin
        grid2.Row = 1
        SQL1 = "SELECT * from AM_CashLin WHERE NoBkt = '" & txtbukti & "' and kodebayar = 'PM'"
        Set RST1 = OBJ.Execute(SQL1)
        Do While Not RST1.EOF
            grid2.TextMatrix(grid2.Row, 0) = RST1!noapply
            grid2.TextMatrix(grid2.Row, 2) = Format(RST1!jumlah, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 3) = Format(RST1!potongan, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 4) = Format(RST1!selisih, "###,###,###,##0.00")
            grid2.TextMatrix(grid2.Row, 5) = Format(RST1!nilaikurs, "###,###,###,##0.00")
            
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST1.MoveNext
        Loop
        'keluarkan record dari aropnfil
        If lblbase = "0" Then   'USD/YEN DLL
            SQL = "select a.NoApply, sum(((a.Amount + a.potongan + a.PPN + a.selisih)* a.nilaikurs)-isnull(b.nilaikurs,0)) as Total from AM_Aropnfil a left join am_cashlin b on a.nobkt=b.nobkt and a.noapply=b.noapply WHERE a.nobkt <> '" & txtbukti & "' and a.kodecust = '" & txtsup & "' and a.tglbkt <= '" & tanggal1 & "' group by a.Noapply"
        Else                    'IDR
            SQL = "select NoApply, sum(Amount + potongan + PPN + selisih) as Total from AM_Aropnfil WHERE nobkt <> '" & txtbukti & "' and kodecust = '" & txtsup & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' group by Noapply"
        End If
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            'If Round(RST!total, 0) = 0 Then
            If RST!total = 0 Then
                RST.MoveNext
                GoTo jump3
            End If
            'cek antara yg di grid ama aropnfil
            For i = 1 To grid2.Rows - 2
                If grid2.TextMatrix(i, 0) = RST!noapply Then
                    grid2.TextMatrix(i, 1) = Format(RST!total, "###,###,###,##0.00")
                    grid2.TextMatrix(i, 6) = Format(RST!total - Val(Format(grid2.TextMatrix(i, 2), "general number")) - Val(Format(grid2.TextMatrix(i, 3), "general number")) + Val(Format(grid2.TextMatrix(i, 4), "general number")), "###,###,###,##0.00")
            
                    RST.MoveNext
                    GoTo jump3
                End If
            Next i
            'cek yg tanggalnya lebih dari tanggal piutang
            If lblbase = "1" Then
                SQL1 = "select a.NoApply, sum(((a.Amount + a.potongan + a.PPN + a.selisih)* a.nilaikurs)-isnull(b.nilaikurs,0)) as Total from AM_Aropnfil a left join am_cashlin b on a.nobkt=b.nobkt and a.noapply=b.noapply WHERE a.noapply = '" & RST!noapply & "' and a.kodecust = '" & txtsup & "' group by a.Noapply"
            Else
                SQL1 = "select NoApply, sum(Amount + potongan + PPN + selisih) as Total from AM_Aropnfil WHERE noapply = '" & RST!noapply & "' and kodecust = '" & txtsup & "' and kodecur = '" & txtkurs & "' group by Noapply"
            End If
            Set RST1 = OBJ.Execute(SQL1)
            'If Round(RST1!total, 0) = 0 Then
            If RST1!total = 0 Then
                RST.MoveNext
                GoTo jump3
            End If
            
            'kalo nga ada di grid nambah dari aropnfil
            grid2.TextMatrix(grid2.Row, 0) = RST!noapply
            grid2.TextMatrix(grid2.Row, 1) = Format(RST!total, "###,###,###,##0.00")
            If grid2.TextMatrix(grid2.Row, 2) = "" Then grid2.TextMatrix(grid2.Row, 2) = "0.00"
            If grid2.TextMatrix(grid2.Row, 3) = "" Then grid2.TextMatrix(grid2.Row, 3) = "0.00"
            If grid2.TextMatrix(grid2.Row, 4) = "" Then grid2.TextMatrix(grid2.Row, 4) = "0.00"
            If grid2.TextMatrix(grid2.Row, 5) = "" Then grid2.TextMatrix(grid2.Row, 5) = "0.00"
            grid2.TextMatrix(grid2.Row, 6) = Format(RST!total, "###,###,###,##0.00")
            
            RST.MoveNext
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
jump3:
        Loop
        
        lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
        
        grid2.Rows = grid2.Rows - 1
        grid2.Col = 0
        grid2.Sort = flexSortStringAscending
        grid2.Rows = grid2.Rows + 1
        
        Grid1.Row = 1
        
        OBJ1.Open dsn
        SQL1 = "SELECT * from AM_Cashsub WHERE NoBkt = '" & txtbukti & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        Do While Not RST1.EOF
            If RST1!Typebayar = "TN" Then Grid1.TextMatrix(Grid1.Row, 1) = "Tunai"
            If RST1!Typebayar = "C" Then Grid1.TextMatrix(Grid1.Row, 1) = "Cek"
            If RST1!Typebayar = "G" Then Grid1.TextMatrix(Grid1.Row, 1) = "Giro"
            If RST1!Typebayar = "TF" Then Grid1.TextMatrix(Grid1.Row, 1) = "Transfer"
            Grid1.TextMatrix(Grid1.Row, 2) = RST1!nogiro
            
            If RST1!Typebayar <> "TN" Then Grid1.TextMatrix(Grid1.Row, 3) = Format(RST1!tgljt, "dd/MM/yyyy")
            
            Grid1.TextMatrix(Grid1.Row, 4) = RST1!bank
            Grid1.TextMatrix(Grid1.Row, 5) = RST1!acbank
            Grid1.TextMatrix(Grid1.Row, 6) = Format(RST1!jumlah, "###,###,###,##0.00")
            
            Grid1.TextMatrix(Grid1.Row, 7) = Format(RST1!tglcair, "dd/MM/yyyy")
            Grid1.TextMatrix(Grid1.Row, 8) = Format(RST1!tgltolak, "dd/MM/yyyy")
            
            SetRow Grid1.Row, True
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Row = Grid1.Row + 1
            RST1.MoveNext
        Loop
        OBJ1.Close
        
        lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")
        
        txtsisa = hitbayar1 - hitbayar
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtbukti = ""
    End If
    OBJ.Close
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    Grid1.Row = idx
    Grid1.Col = 0
    If hapus Then Set Grid1.CellPicture = uncheck.Picture
    Grid1.Col = 1
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtsisa <> 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid2.Rows = 2 Or Grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
 'GoTo lompati:
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar='PM' and idupdate='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close '-----------------------
        MsgBox "Can not update, data already process.", vbExclamation, "Warning"
        Exit Sub '------------------------
    End If
    OBJ.Close
    
    OBJ1.Open dsn
    SQL1 = "select * from gl_transaksi where notrx = '" & txtbukti & "' and identry='auto'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close ''----------------------
        
        MsgBox "Data tidak bisa diupdate, transaksi pembayaran ada di transaksi GL.", vbExclamation, "Warning"
        Exit Sub '-------------------------
    End If
    OBJ1.Close
'lompati:
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        
        If Grid1.TextMatrix(Grid1.Row, 1) <> "Tunai" And Grid1.TextMatrix(Grid1.Row, 3) = "" Then
            MsgBox "Data Entry Not Complete, acc sparta is empty.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        Grid1.Row = Grid1.Row + 1
    Loop
    
    str2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            str2 = 1
            Exit Do
        End If
        grid2.Row = grid2.Row + 1
    Loop
    
    If str2 = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
            
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            OBJ.Open dsn
            SQL = "select * from am_aropnfil where noapply = '" & grid2.TextMatrix(grid2.Row, 0) & "' and transtype <> 'PM'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                MsgBox "Data Entry Not Complete, please refresh customer.", vbExclamation, "Warning"
                OBJ.Close
                Exit Sub
            End If
            OBJ.Close
        End If
        
        grid2.Row = grid2.Row + 1
    Loop
        
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    in1
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub in1()
    OBJ.Open dsn
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!identry
        date2 = RST!dateentry
    End If
    
    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashsub where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_aropnfil where nobkt = '" & txtbukti & "' and transtype = 'PM'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_CashHdr"
    SQL = SQL + " (Kodecust"
    SQL = SQL + ", NoBkt"
    SQL = SQL + ", TglBkt"
    SQL = SQL + ", kodebayar"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", Amount"
    SQL = SQL + ", noac"
    SQL = SQL + ", kodecol"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", Posted"
    SQL = SQL + ", IdEntry"
    SQL = SQL + ", DateEntry"
    SQL = SQL + ", IdUpdate"
    SQL = SQL + ", DateUpdate)"
                
    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", 'PM'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ",Convert(Money," & hitbayar & ")"
    SQL = SQL + ", '0'"
    SQL = SQL + ", '" & txtkodecol & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert(Money," & txtnilaikurs & ")"
    If lblflag.Visible = True Then SQL = SQL + ", '1'"
    If lblflag.Visible = False Then SQL = SQL + ", '0'"
    SQL = SQL + ", '" & str2 & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal2 & "')"
    SQL = SQL + ", '0'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    Grid1.Row = 1
    OBJ.Open dsn
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        
        SQL = "INSERT INTO AM_Cashsub"
        SQL = SQL + " (NoBkt"
        SQL = SQL + ", tglbkt"
        SQL = SQL + ", typeBayar"
        SQL = SQL + ", Kodecust"
        SQL = SQL + ", Nogiro"
        SQL = SQL + ", tgljt"
        SQL = SQL + ", tglcair"
        SQL = SQL + ", tgltolak"
        SQL = SQL + ", bank"
        SQL = SQL + ", acbank"
        SQL = SQL + ", jumlah)"
        
        SQL = SQL + " VALUES"
        SQL = SQL + " ('" & txtbukti & "'"
        SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then SQL = SQL + ", 'TN'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Cek" Then SQL = SQL + ", 'C'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Giro" Then SQL = SQL + ", 'G'"
        If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" Then SQL = SQL + ", 'TF'"
        SQL = SQL + ", '" & txtsup & "'"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
        If Grid1.TextMatrix(Grid1.Row, 3) = "" Then SQL = SQL + ",convert(datetime,'01/01/1900')"
        If Grid1.TextMatrix(Grid1.Row, 3) <> "" Then SQL = SQL + ",convert(datetime,'" & tanggalgrid & "')"
        SQL = SQL + ",convert(DateTime,'" & tanggalgrid1 & "')"
        SQL = SQL + ",convert(DateTime,'" & tanggalgrid2 & "')"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"
        SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
        SQL = SQL + ",convert(money,'" & Format(Grid1.TextMatrix(Grid1.Row, 6), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)
        
        Grid1.Row = Grid1.Row + 1
    Loop
    OBJ.Close
    
    grid2.Row = 1
    OBJ.Open dsn
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        
        If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
            
            SQL = "INSERT INTO AM_CashLin"
            SQL = SQL + " (NoBkt"
            SQL = SQL + ", tglbkt"
            SQL = SQL + ", KodeBayar"
            SQL = SQL + ", Kodecust"
            SQL = SQL + ", NoApply"
            SQL = SQL + ", nilaikurs"
            SQL = SQL + ", jumlah"
            SQL = SQL + ", selisih"
            SQL = SQL + ", potongan)"
            
            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
            SQL = SQL + ", 'PM'"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 5), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 2), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 4), "general number") & "')"
            SQL = SQL + ",convert(money,'" & Format(grid2.TextMatrix(grid2.Row, 3), "general number") & "'))"
            Set RST = OBJ.Execute(SQL)
            
            addpay
        End If
        grid2.Row = grid2.Row + 1
    Loop
    OBJ.Close
End Sub

Private Sub addpay()
    SQL = "INSERT INTO AM_Aropnfil"
    SQL = SQL + " (KodeCust"
    SQL = SQL + ", NoBkt"
    SQL = SQL + ", TglBkt"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", TransType"
    SQL = SQL + ", JatuhTempo"
    SQL = SQL + ", Keterangan"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", Amount"
    SQL = SQL + ", Potongan"
    SQL = SQL + ", selisih"
    SQL = SQL + ", PPN)"

    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
    SQL = SQL + ", '" & grid2.TextMatrix(grid2.Row, 0) & "'"
    SQL = SQL + ", 'PM'"
    SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
    SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 2), "General number") * -1 & "')"
    SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 3), "General number") * -1 & "')"
    SQL = SQL + ",Convert (Money, '" & Format(grid2.TextMatrix(grid2.Row, 4), "General number") & "')"
    SQL = SQL + ",Convert (Money, 0))"
    Set RST = OBJ.Execute(SQL)
End Sub

Function hitbayar()
    hitbayar = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        hitbayar = Val(hitbayar) + Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number"))
        
        grid2.Row = grid2.Row + 1
    Loop
End Function

Function hitbayar1()
    hitbayar1 = 0
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        hitbayar1 = Val(hitbayar1) + Val(Format(Grid1.TextMatrix(Grid1.Row, 6), "general number"))
        
        Grid1.Row = Grid1.Row + 1
    Loop
End Function

Function hitbayar2()
    hitbayar2 = 0
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        hitbayar2 = Val(hitbayar2) + Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number"))
        
        grid2.Row = grid2.Row + 1
    Loop
End Function

Private Sub hapusrow()
    Grid1.TextMatrix(Grid1.Row, 1) = ""
    Grid1.TextMatrix(Grid1.Row, 2) = ""
    Grid1.TextMatrix(Grid1.Row, 3) = ""
    Grid1.TextMatrix(Grid1.Row, 4) = ""
    Grid1.TextMatrix(Grid1.Row, 5) = ""
    Grid1.TextMatrix(Grid1.Row, 6) = ""
    Grid1.TextMatrix(Grid1.Row, 7) = ""
    Grid1.TextMatrix(Grid1.Row, 8) = ""
    Do While True
        If Grid1.TextMatrix(Grid1.Row + 1, 1) = "" Then
            Grid1.TextMatrix(Grid1.Row, 1) = ""
            Grid1.TextMatrix(Grid1.Row, 2) = ""
            Grid1.TextMatrix(Grid1.Row, 3) = ""
            Grid1.TextMatrix(Grid1.Row, 4) = ""
            Grid1.TextMatrix(Grid1.Row, 5) = ""
            Grid1.TextMatrix(Grid1.Row, 6) = ""
            Grid1.TextMatrix(Grid1.Row, 7) = ""
            Grid1.TextMatrix(Grid1.Row, 8) = ""
            Exit Do
        End If
        Grid1.TextMatrix(Grid1.Row, 1) = Grid1.TextMatrix(Grid1.Row + 1, 1)
        Grid1.TextMatrix(Grid1.Row, 2) = Grid1.TextMatrix(Grid1.Row + 1, 2)
        Grid1.TextMatrix(Grid1.Row, 3) = Grid1.TextMatrix(Grid1.Row + 1, 3)
        Grid1.TextMatrix(Grid1.Row, 4) = Grid1.TextMatrix(Grid1.Row + 1, 4)
        Grid1.TextMatrix(Grid1.Row, 5) = Grid1.TextMatrix(Grid1.Row + 1, 5)
        Grid1.TextMatrix(Grid1.Row, 6) = Grid1.TextMatrix(Grid1.Row + 1, 6)
        Grid1.TextMatrix(Grid1.Row, 7) = Grid1.TextMatrix(Grid1.Row + 1, 7)
        Grid1.TextMatrix(Grid1.Row, 8) = Grid1.TextMatrix(Grid1.Row + 1, 8)
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = Grid1.Rows - 1
    Grid1.Col = 0
    Set Grid1.CellPicture = blank
    
    lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")
            
    txtsisa = hitbayar1 - hitbayar
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtkodecol = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 0
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.CellPicture = uncheck Then
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    hapusrow
                    Exit Sub
                End If
            End If
        Case 1
            If Grid1.TextMatrix(Grid1.Row, 1) <> "" Then Exit Sub
            
            If Frame2.Visible = True Then Exit Sub
            
            Frame2.Width = Grid1.ColWidth(Grid1.Col) - 20
            cmbtype.Width = Grid1.ColWidth(Grid1.Col) - 20
            cmbtype = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            Frame2.Left = Grid1.Left + Grid1.CellLeft - 10
            Frame2.Top = Grid1.Top + Grid1.CellTop - 20
            Frame2.Visible = True
            cmbtype.SetFocus
        Case 3
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            
            If date3.Visible = True Then Exit Sub
            
            date3.Width = Grid1.ColWidth(Grid1.Col) - 20
            date3.Height = 290
            If Grid1.TextMatrix(Grid1.Row, Grid1.Col) <> "" Then date3 = Grid1.TextMatrix(Grid1.Row, 3)
            date3.Left = Grid1.Left + Grid1.CellLeft - 10
            date3.Top = Grid1.Top + Grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
        Case 2, 4
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtket = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtket.Left = Grid1.Left + Grid1.CellLeft
            txtket.Top = Grid1.Top + Grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 5
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            carisql1 = "select acc,description from am_bank"
            namatabel = "Acc Sparta"

            frmsearch.Show vbModal
        Case 6
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_EnterCell()
    If txtbukti = "" Or txtsup = "" Or txtkurs = "" Or txtkodecol = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 3
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
                        
            If date3.Visible = True Then Exit Sub
            
            date3.Width = Grid1.ColWidth(Grid1.Col) - 20
            date3.Height = 290
            If Grid1.TextMatrix(Grid1.Row, 3) <> "" Then date3 = Grid1.TextMatrix(Grid1.Row, 3)
            date3.Left = Grid1.Left + Grid1.CellLeft - 10
            date3.Top = Grid1.Top + Grid1.CellTop - 20
            date3.Visible = True
            date3 = Date
            date3.SetFocus
        Case 2, 4
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Tunai" Then Exit Sub
            If Grid1.TextMatrix(Grid1.Row, 1) = "Transfer" And Grid1.Col = 2 Then Exit Sub
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtket = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtket.Left = Grid1.Left + Grid1.CellLeft
            txtket.Top = Grid1.Top + Grid1.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 6
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai1.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai1 = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai1.Left = Grid1.Left + Grid1.CellLeft
            txtnilai1.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai1.Visible = True
            txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case Grid1.Col
        Case 5
            Grid1.Row = posrow
            Grid1.Col = 5
            Grid1.CellAlignment = 1
            Grid1.TextMatrix(Grid1.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
    End Select
End Sub

Private Sub Grid1_Scroll()
    Frame2.Visible = False
    txtket.Visible = False
    txtnilai1.Visible = False
    date3.Visible = False
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    Cariar
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdcari.SetFocus
End Sub

Private Sub txtcari_LostFocus()
    If txtcari = "" Then Exit Sub
        
    OBJ.Open dsn
    SQL = "select top 1 nobkt from am_cashsub where nogiro = '" & txtcari & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtbukti = RST!nobkt
        
        Frame1.Visible = False
        txtcari = ""
    Else
        MsgBox "Data Cek/Giro not found.", vbExclamation, "Warning"
        txtcari = ""
        txtcari.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    Cariar
End Sub

Private Sub txtcari1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdcari1.SetFocus
End Sub

Private Sub txtcari1_LostFocus()
    If txtcari1 = "" Then Exit Sub
        
    OBJ.Open dsn
    SQL = "select count(nobkt)'hit' from am_cashlin where noapply = '" & txtcari1 & "' and kodebayar = 'PM' group by noapply"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then i = RST!hit Else i = 0
    OBJ.Close
    
    If i = 1 Then
        OBJ.Open dsn
        SQL = "select nobkt from am_cashlin where noapply = '" & txtcari1 & "' and kodebayar = 'PM'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtbukti = RST!nobkt
        OBJ.Close
        
        Cariar
        
        Frame1.Visible = False
        txtcari1 = ""
    ElseIf i = 0 Then
        MsgBox "Data Faktur not found.", vbExclamation, "Warning"
        txtcari1 = ""
        txtcari1.SetFocus
    ElseIf i > 1 Then
        carisql1 = "select nobkt,noapply from AM_cashlin where kodebayar = 'PM' and noapply = '" & txtcari1 & "'"
        namatabel = "Faktur"
        
        frmsearch.Show vbModal
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 27 Then
        txtket_LostFocus
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case Grid1.Col
            Case 2
                For i = 1 To Grid1.Rows - 2
                    If Grid1.TextMatrix(i, 1) = "" Then Exit For
                    If Grid1.TextMatrix(i, 2) = Trim(txtket) Then
                        txtket = ""
                        txtket.Visible = False
                        MsgBox "No Cek/Giro Already exist.", vbExclamation, "Information"
                        
                        Exit Sub
                    End If
                Next i
                
                OBJ2.Open dsn
                SQL2 = "select * from am_cashsub where nogiro = '" & Trim(txtket) & "'"
                Set RST2 = OBJ2.Execute(SQL2)
                If Not RST2.EOF Then
                    OBJ2.Close
                    txtket = ""
                    txtket.Visible = False
                    MsgBox "No Cek/Giro Already exist.", vbExclamation, "Information"
                    
                    Exit Sub
                End If
                OBJ2.Close
                
                Grid1.SetFocus
                Grid1.TextMatrix(Grid1.Row, 2) = Trim(txtket)
                txtket = ""
                txtket.Visible = False
            Case 4
                Grid1.Row = posrow
                
                Grid1.SetFocus
                Grid1.Col = 4
                Grid1.CellAlignment = 1
                Grid1.TextMatrix(Grid1.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Grid1.SetFocus
End Sub

Private Sub txtkodecol_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkodecol_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtkurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, grid2.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
                
        If grid2.Col = 3 Then
            If Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) < 0 Then
                grid2.SetFocus
                grid2.TextMatrix(grid2.Row, 3) = "0.00"
                txtnilai = 0
                Exit Sub
            End If
        End If
        
        lblapply = "Total Apply : " & Format(hitbayar2, "###,###,##0.00")
        lblbayar = "Bayar Apply : " & Format(hitbayar, "###,###,##0.00")
        
        txtsisa = hitbayar1 - hitbayar
        If lblbase = "0" Then hitselisihkurs
                
        grid2.TextMatrix(posrow, 6) = Format((Format(grid2.TextMatrix(posrow, 1), "general number") - Format(grid2.TextMatrix(posrow, 2), "general number") - Format(grid2.TextMatrix(posrow, 3), "general number") + Format(grid2.TextMatrix(posrow, 4), "general number")), "###,###,###,##0.00")
        
        txtnilai.Visible = False
        grid2.SetFocus
        grid2.Row = posrow
    End If
    If KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnilai1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Format(txtnilai1, "###,###,##0.00")
        txtnilai1 = 0
        
        lbltotal = "Total Bayar : " & Format(hitbayar1, "###,###,##0.00")
        
        txtsisa = hitbayar1 - hitbayar
        
        txtnilai1.Visible = False
        Grid1.SetFocus
        Grid1.Row = posrow
    End If
    If KeyAscii = 27 Then
        txtnilai1 = 0
        txtnilai1.Visible = False
    End If
End Sub

Private Sub txtnilai1_LostFocus()
    txtnilai1.Visible = False
    txtnilai1 = 0
End Sub

Private Sub txtnilaikurs_Change()
    If grid2.Rows > 2 Then
        grid2.Row = 1
        Do While True
            grid2.TextMatrix(grid2.Row, 5) = "0.00"
            If grid2.TextMatrix(grid2.Row, 2) <> "0.00" Then
                posrow = grid2.Row
                If lblbase = "0" Then hitselisihkurs
            End If
            
            grid2.Row = grid2.Row + 1
            If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        Loop
    End If
End Sub

Private Sub txtsisa_Change()
    lblsisa = " Sisa : " & Format(txtsisa, "###,###,##0.00")
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub txtsup_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkodecol.SetFocus
    KeyAscii = 0
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Then Exit Sub
    posrow = grid2.Row
    
    Select Case grid2.Col
    Case 2
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
            
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 3
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
            
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
            
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid2_EnterCell()
    If grid2.MouseRow = 0 Then Exit Sub
    If txtbukti = "" Or txtsup = "" Or txtkodecol = "" Then Exit Sub
    Select Case grid2.Col
    Case 2
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
            
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 3), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 3
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
        
        txtnilai = (Val(Format(grid2.TextMatrix(grid2.Row, 1), "general number")) - Val(Format(grid2.TextMatrix(grid2.Row, 2), "general number")) + Val(Format(grid2.TextMatrix(grid2.Row, 4), "general number")))
        If txtnilai < 0 Then txtnilai = txtnilai * -1
    Case 4
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Sub
        
        posrow = grid2.Row
        
        txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
        txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
        txtnilai.Left = grid2.Left + grid2.CellLeft
        txtnilai.Top = grid2.Top + grid2.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub hitselisihkurs()
    OBJ.Open dsn
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype<>'PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai2 = RST!total
    Else
        txtnilai2 = 0
    End If
    
    SQL = "select isnull(sum(Amount + potongan + PPN + selisih),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and nobkt <> '" & txtbukti & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype='PM'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnilai3 = RST!total * -1
    Else
        txtnilai3 = 0
    End If
    
    txtnilai4 = Val(Format(grid2.TextMatrix(posrow, 2), "general number")) + Val(Format(grid2.TextMatrix(posrow, 3), "general number")) - Val(Format(grid2.TextMatrix(posrow, 4), "general number"))
    txtnilai5 = txtnilai4 + txtnilai3
    grid2.TextMatrix(posrow, 5) = "0.00"
    If txtnilai2 = txtnilai5 Then
        SQL = "select isnull(sum((Amount + potongan + PPN + selisih)*nilaikurs),0) as Total from AM_Aropnfil WHERE noapply = '" & grid2.TextMatrix(posrow, 0) & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and transtype<>'PM'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtnilai2 = RST!total
        Else
            txtnilai2 = 0
        End If
        
        SQL = "select isnull(sum(((a.Amount + a.potongan + a.PPN + a.selisih)* a.nilaikurs)-isnull(b.nilaikurs,0)),0) as Total from AM_Aropnfil a left join am_cashlin b on a.nobkt=b.nobkt and a.noapply=b.noapply WHERE a.noapply = '" & grid2.TextMatrix(posrow, 0) & "' and a.nobkt <> '" & txtbukti & "' and a.kodecur = '" & txtkurs & "' and a.tglbkt <= '" & tanggal1 & "' and a.transtype='PM'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtnilai3 = RST!total * -1
        Else
            txtnilai3 = 0
        End If
        
        txtnilai4 = Val((Format(grid2.TextMatrix(posrow, 2), "general number")) + Val(Format(grid2.TextMatrix(posrow, 3), "general number")) - Val(Format(grid2.TextMatrix(posrow, 4), "general number"))) * txtnilaikurs
        txtnilai5 = txtnilai4 + txtnilai3
        
        If txtnilai2 <> txtnilai5 Then
            grid2.TextMatrix(posrow, 5) = Format(txtnilai2 - txtnilai3 - txtnilai4, "###,###,##0.00")
        End If
    End If
    OBJ.Close
End Sub

Private Sub caricollector()
    If txtkodecol = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * from AM_collector WHERE Kode = '" & txtkodecol & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Collector " & RST!nama & " is not active !", vbExclamation, "Warning"
        txtkodecol = ""
        lblnamacol = ""
    End If
    OBJ.Close
End Sub
