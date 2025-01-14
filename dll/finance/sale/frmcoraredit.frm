VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcoraredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Koreksi Piutang"
   ClientHeight    =   6900
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   10125
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
   Icon            =   "frmcoraredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmcoraredit.frx":2372
      Caption         =   "frmcoraredit.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":23FE
      Keys            =   "frmcoraredit.frx":241C
      Spin            =   "frmcoraredit.frx":245E
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.PictureBox picbutton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   3855
      TabIndex        =   35
      Top             =   6360
      Width           =   3855
      Begin Chameleon.chameleonButton cmdclose 
         Height          =   375
         Left            =   2880
         TabIndex        =   36
         Top             =   0
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
         MICON           =   "frmcoraredit.frx":2486
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
         Left            =   1920
         TabIndex        =   37
         Top             =   0
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
         MICON           =   "frmcoraredit.frx":27A0
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
         Left            =   960
         TabIndex        =   38
         Top             =   0
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
         MICON           =   "frmcoraredit.frx":2ABA
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
         Left            =   0
         TabIndex        =   39
         Top             =   0
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
         MICON           =   "frmcoraredit.frx":2DD4
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
   Begin VB.ComboBox cmbtype 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   114294785
      CurrentDate     =   38516
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin TDBNumber6Ctl.TDBNumber txtjumlah 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":30EE
      Caption         =   "frmcoraredit.frx":310E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":317A
      Keys            =   "frmcoraredit.frx":3198
      Spin            =   "frmcoraredit.frx":31E2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBText6Ctl.TDBText txtbukti 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmcoraredit.frx":320A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3276
      Key             =   "frmcoraredit.frx":3294
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
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   960
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
      Format          =   114294787
      CurrentDate     =   37421
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   8
      Top             =   3000
      Width           =   6255
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No. Bukti"
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
      MICON           =   "frmcoraredit.frx":32D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtkoreksi 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":35F2
      Caption         =   "frmcoraredit.frx":3612
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":367E
      Keys            =   "frmcoraredit.frx":369C
      Spin            =   "frmcoraredit.frx":36E6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00;(#,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00;(#,###,###,##0.00)"
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
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBNumber6Ctl.TDBNumber txtotal 
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   885
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmcoraredit.frx":370E
      Caption         =   "frmcoraredit.frx":372E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":379A
      Keys            =   "frmcoraredit.frx":37B8
      Spin            =   "frmcoraredit.frx":37FA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483631
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   16777215
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   360
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmcoraredit.frx":3822
      Caption         =   "frmcoraredit.frx":3842
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":38AE
      Keys            =   "frmcoraredit.frx":38CC
      Spin            =   "frmcoraredit.frx":390E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtdiscount 
      Height          =   240
      Left            =   6240
      TabIndex        =   21
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   423
      Calculator      =   "frmcoraredit.frx":3936
      Caption         =   "frmcoraredit.frx":3956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":39C2
      Keys            =   "frmcoraredit.frx":39E0
      Spin            =   "frmcoraredit.frx":3A22
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmcoraredit.frx":3A4A
      Caption         =   "frmcoraredit.frx":3A6A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3AD6
      Keys            =   "frmcoraredit.frx":3AF4
      Spin            =   "frmcoraredit.frx":3B36
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdRetur 
      Height          =   285
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Nomor Retur"
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
      MICON           =   "frmcoraredit.frx":3B5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtNomorBPB 
      Height          =   285
      Left            =   1560
      TabIndex        =   34
      Top             =   3720
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmcoraredit.frx":3E78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcoraredit.frx":3EE4
      Key             =   "frmcoraredit.frx":3F02
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
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   2400
      TabIndex        =   30
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Netto"
      Height          =   255
      Left            =   5670
      TabIndex        =   25
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5670
      TabIndex        =   23
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Koreksi"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3390
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Nila Apply"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label lblsup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3240
      TabIndex        =   14
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Type"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Apply To"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000011&
      Height          =   345
      Left            =   5520
      TabIndex        =   24
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "Koreksi"
      Height          =   255
      Left            =   5670
      TabIndex        =   26
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5520
      TabIndex        =   27
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmcoraredit"
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
Dim nort As String
Dim str2 As String
Dim retur As Boolean

Private Sub addkoreksi()
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
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ",Convert(dateTime, '" & tanggal1 & "')"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ",Convert (Money, '" & txtnilaikurs & "')"
    If lbltype = "CN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksi * -1 & "')"
    If lbltype = "DN" Then SQL = SQL + ",Convert (Money, '" & txtkoreksi & "')"
    SQL = SQL + ",Convert (Money, 0)"
    SQL = SQL + ",Convert (Money, 0)"
    SQL = SQL + ",Convert (Money, 0))"
    Set RST = OBJ.Execute(SQL)
End Sub
Private Sub setform()
    If retur = True Then
        Me.Height = 7245
        Me.Width = 10215
        picbutton.Top = 6360
        picbutton.Left = 6000
        grid1.Visible = True
        cmdRetur.Visible = True
        txtNomorBPB.Visible = True
    Else
        Me.Height = 4860
        Me.Width = 8025
        picbutton.Top = 3840
        picbutton.Left = 4080
        grid1.Visible = False
        cmdRetur.Visible = False
        txtNomorBPB.Visible = False
    End If
End Sub
Private Sub cmbtype_Click()
    Select Case cmbtype
        Case "Credit Note"
            lbltype = "CN"
            retur = False
            setform
        Case "Debit Note"
            lbltype = "DN"
            retur = False
            setform
        Case "Retur"
            lbltype = "CN"
            retur = True
            setform
    End Select
    
    txtbukti.SetFocus
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    If cmbtype = "" Then Exit Sub
    
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtbukti = ""
    txtapply = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
End Sub

Private Sub cmdclear_Click()
    txtbukti.Enabled = True
    cmdsearch.Enabled = True
    cmbtype.Enabled = True
    date1.Enabled = True
    
    txtbukti = ""
    cmbtype = ""
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtapply = ""
    txtketerangan = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtjumlah = 0
    txtkoreksi = 0
    hapusgrid1

    Me.Height = 4860
    Me.Width = 8025
    picbutton.Top = 3840
    picbutton.Left = 4080
    grid1.Visible = False
    cmdRetur.Visible = False
    txtNomorBPB.Visible = False
    
    cmbtype.SetFocus
End Sub

Private Sub cmdel_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtsup = "" Or txtbukti = "" Or cmbtype = "" Or txtapply = "" Or txtkoreksi = 0 Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    If lbltype = "CN" And txtapply = txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtbukti & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply = txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtbukti & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    If lbltype = "CN" And txtapply <> txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtapply & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply <> txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtapply & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete data already have apply.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    

    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_aropnfil where nobkt = '" & txtbukti & "' and transtype = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    If retur = True Then
        SQL = "Delete form am_returjual Where nobkt = '" & txtbukti & "' and noretur = '" & txtNomorBPB & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close

    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='242' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='243' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    Select Case grid1.Col
        Case 7
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai.Left = grid1.Left + grid1.CellLeft
            txtnilai.Top = grid1.Top + grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub txtapply_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtjumlah.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtbukti_LostFocus()
    caricor
End Sub

Private Sub Form_Load()
    
    initGrid1
    date1 = Date
    cmbtype.AddItem "Debit Note"
    cmbtype.AddItem "Credit Note"
    cmbtype.AddItem "Retur"

    Me.Height = 4860
    Me.Width = 8025
    picbutton.Top = 3840
    picbutton.Left = 4080
    grid1.Visible = False
    cmdRetur.Visible = False
    txtNomorBPB.Visible = False
End Sub

Private Sub txtjumlah_Change()
    txtneto = txtjumlah
    txtdiscount = txtkoreksi
    
    If lbltype = "CN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtkoreksi_Change()
    txtneto = txtjumlah
    txtdiscount = txtkoreksi
    
    If lbltype = "CN" Then
        txtotal = txtneto - txtdiscount
    Else
        txtotal = txtneto + txtdiscount
    End If
End Sub

Private Sub txtkurs_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.TextMatrix(grid1.Row, grid1.Col) = Format(txtnilai, "###,###,##0.00")
        grid1.TextMatrix(grid1.Row, 8) = grid1.TextMatrix(grid1.Row, 3) * grid1.TextMatrix(grid1.Row, 7)
        grid1.TextMatrix(grid1.Row, 8) = Format(grid1.TextMatrix(grid1.Row, 8), "###,###,##0.00")
        total
        txtnilai = 0
        txtnilai.Visible = False
        grid1.SetFocus
    End If
    If KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
    txtnilai = 0
End Sub

Private Sub total()
    Dim tj As Double
    tj = 0
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        tj = tj + CDbl(grid1.TextMatrix(grid1.Row, 8))
        grid1.Row = grid1.Row + 1
    Loop
    txtkoreksi = tj
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
    If txtapply <> txtbukti Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsup.SetFocus
    If lblbase = "1" Then KeyAscii = 0
    If txtapply <> txtbukti Then KeyAscii = 0
End Sub

Private Sub txtsup_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtapply.SetFocus
    KeyAscii = 0
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        
        Select Case cmbtype
        Case "Credit Note"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'CN-%' and kodebayar = '" & lbltype & "' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
        Case "Debit Note"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'DN-%' and kodebayar = '" & lbltype & "' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
        Case "Retur"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'RT-%' and kodebayar = '" & lbltype & "' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
        End Select
    Else
        Select Case cmbtype
        Case "Credit Note"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'CN-%' and kodebayar = '" & lbltype & "'"
        Case "Debit Note"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'DN-%' and kodebayar = '" & lbltype & "'"
        Case "Retur"
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt' from AM_Cashhdr where nobkt like 'RT-%' and kodebayar = '" & lbltype & "'"
        End Select
    End If
    namatabel = "Koreksi"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtbukti = hasil
    nort = hasil
    caricor
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub caricor()
    If txtbukti = "" Or cmbtype = "" Then Exit Sub
    If txtbukti.SelLength <> 0 Then Exit Sub
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_CashHdr WHERE NoBkt = '" & txtbukti & "' And kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtsup = RST!kodecust
        date1 = RST!tglbkt
        txtapply = RST!noapply
        txtketerangan = RST!keterangan
        txtjumlah = RST!amount
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        
        SQL = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblbase = RST!base
        
        SQL = "SELECT * FROM AM_Cashlin WHERE NoBkt = '" & txtbukti & "' And kodebayar = '" & lbltype & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtkoreksi = RST!jumlah
        
        SQL = "SELECT * FROM AM_Customer WHERE kodecust = '" & txtsup & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsup = RST!namacust
        Else
            lblsup = ""
        End If
    
        txtbukti.Enabled = False
        cmdsearch.Enabled = False
        cmbtype.Enabled = False
        date1.Enabled = False
    End If
    OBJ.Close
    If retur = True Then addreturdetail
End Sub
Private Sub addreturdetail()
    OBJ.Open dsn
    SQL = "Select a.*,b.NamaBarang,c.NamaSatuan From am_returjual a inner join am_itemdtl b"
    SQL = SQL + " on a.kodebarang = b.KodeBarang inner join am_unit c"
    SQL = SQL + " on a.kodesatuan = c.KodeSatuan where a.nobkt = '" & nort & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then OBJ.Close: Exit Sub
    txtNomorBPB = RST!noretur
    grid1.Row = 1
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = RST!noretur
        grid1.TextMatrix(grid1.Row, 2) = RST!namabarang
        grid1.TextMatrix(grid1.Row, 3) = RST!qty
        grid1.TextMatrix(grid1.Row, 4) = RST!kodesatuan
        grid1.TextMatrix(grid1.Row, 5) = RST!namasatuan
        grid1.TextMatrix(grid1.Row, 6) = RST!kodebarang
        grid1.TextMatrix(grid1.Row, 7) = Format(RST!nilai, "#,##0.00")
        grid1.TextMatrix(grid1.Row, 8) = Format(RST!nilai * RST!qty, "#,##0.00")
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtbukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtbukti.SetFocus
        Exit Sub
    End If
    
    If txtsup = "" Or txtbukti = "" Or cmbtype = "" Or txtapply = "" Or txtkoreksi = 0 Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggal1 & "' and tanggal2 >= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    If lbltype = "CN" And txtapply = txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtbukti & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply = txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtbukti & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    If lbltype = "CN" And txtapply <> txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtapply & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'DN')"
    If lbltype = "DN" And txtapply <> txtbukti Then SQL = "select * from am_aropnfil where noapply = '" & txtapply & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'PM' or transtype = 'CN')"
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can not delete data already have apply.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    in1
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub in1()
    OBJ.Open dsn
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!identry
        date2 = RST!dateentry
    End If
    
    SQL = "delete from am_cashhdr where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_cashlin where nobkt = '" & txtbukti & "' and kodebayar = '" & lbltype & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_aropnfil where nobkt = '" & txtbukti & "' and transtype = '" & lbltype & "'"
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
    SQL = SQL + ", Posted"
    SQL = SQL + ", kodecur"
    SQL = SQL + ", nilaikurs"
    SQL = SQL + ", IdEntry"
    SQL = SQL + ", DateEntry"
    SQL = SQL + ", IdUpdate"
    SQL = SQL + ", DateUpdate)"
        
    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtsup & "'"
    SQL = SQL + ", '" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ", '" & txtketerangan & "'"
    SQL = SQL + ", Convert(Money,'" & txtjumlah & "')"
    SQL = SQL + ", 'x'"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '0'"
    SQL = SQL + ", '" & txtkurs & "'"
    SQL = SQL + ", Convert(Money,'" & txtnilaikurs & "')"
    SQL = SQL + ", '" & str2 & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal2 & "')"
    SQL = SQL + ", '0'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_CashLin "
    SQL = SQL + " (NoBkt"
    SQL = SQL + ", tglbkt"
    SQL = SQL + ", KodeBayar"
    SQL = SQL + ", Kodecust"
    SQL = SQL + ", NoApply"
    SQL = SQL + ", Jumlah"
    SQL = SQL + ", selisih"
    SQL = SQL + ", Potongan"
    SQL = SQL + ", NilaiKurs)"

    SQL = SQL + " VALUES"
    SQL = SQL + " ('" & txtbukti & "'"
    SQL = SQL + ",convert(datetime,'" & tanggal1 & "')"
    SQL = SQL + ", '" & lbltype & "'"
    SQL = SQL + ", '" & txtsup & "'"
    SQL = SQL + ", '" & txtapply & "'"
    SQL = SQL + ",convert(money,'" & txtkoreksi & "')"
    SQL = SQL + ",convert(money,'0')"
    SQL = SQL + ",convert(money,'0')"
    SQL = SQL + ",convert(money,'" & txtnilaikurs & "'))"
    Set RST = OBJ.Execute(SQL)
            
    addkoreksi
    
    If retur = True Then
        SQL = "Delete From am_returjual Where nobkt = '" & txtbukti & "' and noretur = '" & txtNomorBPB & "'"
        Set RST = OBJ.Execute(SQL)
        
        grid1.Row = 1
        Do While True
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            SQL = "Insert into am_returjual "
            SQL = SQL + " (nobkt"
            SQL = SQL + ", kdcust"
            SQL = SQL + ", noapply"
            SQL = SQL + ", noretur"
            SQL = SQL + ", kodebarang"
            SQL = SQL + ", nilai"
            SQL = SQL + ", qty"
            SQL = SQL + ", kodesatuan"
            SQL = SQL + ", lineitem)"
            
            SQL = SQL + " VALUES"
            SQL = SQL + " ('" & txtbukti & "'"
            SQL = SQL + ", '" & txtsup & "'"
            SQL = SQL + ", '" & txtapply & "'"
            SQL = SQL + ", '" & txtNomorBPB & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 6) & "'"
            SQL = SQL + ",convert(money,'" & Format(grid1.TextMatrix(grid1.Row, 7), "general number") & "')"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 3) & "'"
            SQL = SQL + ", '" & grid1.TextMatrix(grid1.Row, 4) & "'"
            SQL = SQL + ", '" & grid1.Row & "')"
            Set RST = OBJ.Execute(SQL)
            grid1.Row = grid1.Row + 1
        Loop
        retur = False
    End If
    OBJ.Close
End Sub

Private Sub initGrid1()
    With grid1
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE" 'KD RETUR
        .TextMatrix(0, 2) = "NAMA"
        .TextMatrix(0, 3) = "QTY"
        .TextMatrix(0, 4) = "K/Satuan"
        .TextMatrix(0, 5) = "N/Satuan"
        .TextMatrix(0, 6) = "KODE BARANG"
        .TextMatrix(0, 7) = "PRICE"
        .TextMatrix(0, 8) = "TOTAL"
    End With
    setGrid1
End Sub

Private Sub setGrid1()
    With grid1
        .Width = 9975
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 0
        .ColWidth(7) = 1300
        .ColWidth(8) = 1300
    End With
End Sub

Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.TextMatrix(grid1.Row, 7) = ""
        grid1.TextMatrix(grid1.Row, 8) = ""
        grid1.Col = 1
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    
End Sub
