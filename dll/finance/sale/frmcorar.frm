VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcorar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Koreksi Piutang"
   ClientHeight    =   6915
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   7905
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
   Icon            =   "frmcorar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   2895
      TabIndex        =   35
      Top             =   6360
      Width           =   2895
      Begin Chameleon.chameleonButton cmdclose 
         Height          =   375
         Left            =   1920
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
         MICON           =   "frmcorar.frx":2372
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
         Left            =   960
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
         MICON           =   "frmcorar.frx":268C
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
         TabIndex        =   38
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmcorar.frx":29A6
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
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   6600
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmcorar.frx":2CC0
      Caption         =   "frmcorar.frx":2CE0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":2D4C
      Keys            =   "frmcorar.frx":2D6A
      Spin            =   "frmcorar.frx":2DAC
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   31
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cmbtype 
      Height          =   315
      ItemData        =   "frmcorar.frx":2DD4
      Left            =   1560
      List            =   "frmcorar.frx":2DD6
      TabIndex        =   0
      Top             =   240
      Width           =   1575
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
      Calculator      =   "frmcorar.frx":2DD8
      Caption         =   "frmcorar.frx":2DF8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":2E64
      Keys            =   "frmcorar.frx":2E82
      Spin            =   "frmcorar.frx":2ECC
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
      ValueVT         =   5
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
      Caption         =   "frmcorar.frx":2EF4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":2F60
      Key             =   "frmcorar.frx":2F7E
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
      Format          =   134742019
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
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Customer"
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
      MICON           =   "frmcorar.frx":2FC2
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
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Apply To"
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
      MICON           =   "frmcorar.frx":32DC
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
      Calculator      =   "frmcorar.frx":35F6
      Caption         =   "frmcorar.frx":3616
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":3682
      Keys            =   "frmcorar.frx":36A0
      Spin            =   "frmcorar.frx":36EA
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
      TabIndex        =   21
      Top             =   885
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmcorar.frx":3712
      Caption         =   "frmcorar.frx":3732
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":379E
      Keys            =   "frmcorar.frx":37BC
      Spin            =   "frmcorar.frx":37FE
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
      TabIndex        =   23
      Top             =   360
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmcorar.frx":3826
      Caption         =   "frmcorar.frx":3846
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":38B2
      Keys            =   "frmcorar.frx":38D0
      Spin            =   "frmcorar.frx":3912
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
      TabIndex        =   22
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   423
      Calculator      =   "frmcorar.frx":393A
      Caption         =   "frmcorar.frx":395A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":39C6
      Keys            =   "frmcorar.frx":39E4
      Spin            =   "frmcorar.frx":3A26
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
      Calculator      =   "frmcorar.frx":3A4E
      Caption         =   "frmcorar.frx":3A6E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":3ADA
      Keys            =   "frmcorar.frx":3AF8
      Spin            =   "frmcorar.frx":3B3A
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
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Currency"
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
      MICON           =   "frmcorar.frx":3B62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdRetur 
      Height          =   285
      Left            =   120
      TabIndex        =   32
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
      MICON           =   "frmcorar.frx":3E7C
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
      TabIndex        =   33
      Top             =   3720
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "frmcorar.frx":4196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcorar.frx":4202
      Key             =   "frmcorar.frx":4220
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
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   10
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
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Netto"
      Height          =   255
      Left            =   5670
      TabIndex        =   26
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000011&
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5670
      TabIndex        =   24
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Koreksi"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3390
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Nilai Apply"
      Height          =   255
      Left            =   240
      TabIndex        =   19
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
      Caption         =   "No. Bukti"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000011&
      Height          =   345
      Left            =   5520
      TabIndex        =   25
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "Koreksi"
      Height          =   255
      Left            =   5670
      TabIndex        =   27
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5520
      TabIndex        =   28
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmcorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim str99 As String
Dim int3 As Integer
Dim posrow As String
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

Private Sub cmbtype_Click()
    Select Case cmbtype
        Case "Credit Note"
            lbltype = "CN"
            
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'CN' and nobkt like 'CN-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtbukti = "CN-000" & str99
            If Len(str99) = 2 Then txtbukti = "CN-00" & str99
            If Len(str99) = 3 Then txtbukti = "CN-0" & str99
            If Len(str99) = 4 Then txtbukti = "CN-" & str99
            retur = False
            initGrid1
        Case "Debit Note"
            lbltype = "DN"
            
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'DN' and nobkt like 'DN-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtbukti = "DN-000" & str99
            If Len(str99) = 2 Then txtbukti = "DN-00" & str99
            If Len(str99) = 3 Then txtbukti = "DN-0" & str99
            If Len(str99) = 4 Then txtbukti = "DN-" & str99
            retur = False
            initGrid1
        Case "Retur"
            lbltype = "CN"
            
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'CN' and nobkt like 'RT-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtbukti = "RT-000" & str99
            If Len(str99) = 2 Then txtbukti = "RT-00" & str99
            If Len(str99) = 3 Then txtbukti = "RT-0" & str99
            If Len(str99) = 4 Then txtbukti = "RT-" & str99
            retur = True
            initGrid1
    End Select
    txtapply = txtbukti
    date1.SetFocus
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbtype_LostFocus()
    If cmbtype = "" Then Exit Sub
    If retur = True Then Exit Sub
    date1 = Date
    txtsup = ""
    lblsup = ""
    'txtbukti = ""
    'txtapply = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
End Sub

Private Sub cmdclear_Click()
    txtbukti = ""
    cmbtype = ""
    date1 = Date
    txtsup = ""
    lblsup = ""
    txtapply = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
    txtNomorBPB = ""
    hapusgrid1
    cmbtype.SetFocus
End Sub

Private Sub cmdRetur_Click()
    If txtbukti = "" Then Exit Sub
    If txtkurs = "" Then Exit Sub
    carisql1 = "select convert(char(12),tglbpb,103) as tanggal ,nobpb, keterangan from am_bpbhdr where nobpb like 'TR0-%'"
    namatabel = "Retur"
    frmsearch.Show vbModal
End Sub

Private Sub cmdRetur_GotFocus()
    If hasil = "" Then Exit Sub
    txtNomorBPB = hasil1
    hasil = ""
    hasil1 = ""
    namatabel = ""
    carisql1 = ""
    cariDataRetur
End Sub

Private Sub cariDataRetur()
    On Error GoTo err_handler:
    OBJ.Open dsn
    'query bpbhdr
    SQL = "select a.*, b.NamaCust from am_bpbhdr a "
    SQL = SQL + "inner join am_customer b on a.noref = b.KodeCust "
    SQL = SQL + "where a.Nobpb ='" & txtNomorBPB & "'"
    Set RST = OBJ.Execute(SQL)
    txtsup = RST!noref
    lblsup = RST!namacust
    
    'query bpblin
    SQL = "select distinct  a.*,b.namabarang, c.namasatuan from am_bpblin a "
    SQL = SQL + "inner join am_itemdtl b on a.kodebarang=b.kodebarang "
    SQL = SQL + "inner join am_unit c on a.kodesatuan = c.kodesatuan "
    SQL = SQL + " where nobpb = '" & txtNomorBPB & "'"
    Set RST = OBJ.Execute(SQL)
    hapusgrid1
    Grid1.Row = 1
    Do While Not RST.EOF
        With Grid1
            .TextMatrix(.Row, 1) = RST!nobpb
            .TextMatrix(.Row, 2) = RST!namabarang
            .TextMatrix(.Row, 3) = RST!qty
            .TextMatrix(.Row, 4) = RST!kodesatuan
            .TextMatrix(.Row, 5) = RST!namasatuan
            If retur = True Then
            .TextMatrix(.Row, 6) = RST!kodebarang
            .TextMatrix(.Row, 7) = "0.00"
            .TextMatrix(.Row, 8) = "0.00"
            End If
            .Rows = .Rows + 1
            .Row = .Row + 1
        End With
        RST.MoveNext
    Loop
    OBJ.Close
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub


Private Sub cmdsearch3_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs.text = hasil
    carikurs
    hasil = ""
End Sub

Private Sub carikurs()
    If txtkurs = "" Then Exit Sub
    OBJ2.Open dsn
    SQL2 = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If Not RST2.EOF Then
        If RST2!base = 1 Then
            lblbase = "1"
        Else
            lblbase = "0"
        End If
        Select Case Month(date1)
        Case 1
            txtnilaikurs = RST2!kurs1
        Case 2
            txtnilaikurs = RST2!kurs2
        Case 3
            txtnilaikurs = RST2!kurs3
        Case 4
            txtnilaikurs = RST2!kurs4
        Case 5
            txtnilaikurs = RST2!kurs5
        Case 6
            txtnilaikurs = RST2!kurs6
        Case 7
            txtnilaikurs = RST2!kurs7
        Case 8
            txtnilaikurs = RST2!kurs8
        Case 9
            txtnilaikurs = RST2!kurs9
        Case 10
            txtnilaikurs = RST2!kurs10
        Case 11
            txtnilaikurs = RST2!kurs11
        Case 12
            txtnilaikurs = RST2!kurs12
        End Select
    Else
        MsgBox "Currency " & txtkurs & " Not Found.", vbInformation, "Information"
        txtkurs = ""
        txtkurs.SetFocus
    End If
    OBJ2.Close
End Sub

Private Sub Form_Activate()
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "The period is empty !!" & vbCrLf & _
        "Please define Period on proces, Starting period date and Ending period date.", vbCritical, "Critical"
        
        OBJ.Close
        Unload Me
        Exit Sub
    End If
    OBJ.Close
    
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='241' and b.kodeuser = '1" & kuser & "'"
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

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    Select Case Grid1.Col
        Case 7
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Sub
            txtnilai.Width = Grid1.ColWidth(Grid1.Col) - 40
            txtnilai = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
            txtnilai.Left = Grid1.Left + Grid1.CellLeft
            txtnilai.Top = Grid1.Top + Grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub txtapply_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtapply_LostFocus()
    cariapply
End Sub

Private Sub date1_Change()
    txtsup = ""
    lblsup = ""
    If txtbukti <> txtapply Then txtapply = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
End Sub

Private Sub txtbukti_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtbukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date1.SetFocus
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    cmbtype.AddItem "Debit Note"
    cmbtype.AddItem "Credit Note"
    cmbtype.AddItem "Retur"
    initGrid1
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

Private Sub txtkurs_Change()
    txtsup = ""
    lblsup = ""
    txtapply = txtbukti
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
End Sub

Private Sub txtkurs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkurs_LostFocus
End Sub

Private Sub txtkurs_LostFocus()
    carikurs
End Sub


Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Format(txtnilai, "###,###,##0.00")
        Grid1.TextMatrix(Grid1.Row, 8) = Grid1.TextMatrix(Grid1.Row, 3) * Grid1.TextMatrix(Grid1.Row, 7)
        Grid1.TextMatrix(Grid1.Row, 8) = Format(Grid1.TextMatrix(Grid1.Row, 8), "###,###,##0.00")
        total
        txtnilai = 0
        txtnilai.Visible = False
        Grid1.SetFocus
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
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        tj = tj + CDbl(Grid1.TextMatrix(Grid1.Row, 8))
        Grid1.Row = Grid1.Row + 1
    Loop
    txtkoreksi = tj
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
    If txtbukti <> txtapply Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsup.SetFocus
    If lblbase = "1" Then KeyAscii = 0
    If txtbukti <> txtapply Then KeyAscii = 0
End Sub

Private Sub txtsup_LostFocus()
    If txtsup = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_customer where Kodecust = '" & txtsup & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsup = RST!namacust
    Else
        MsgBox "Customer " & txtsup & " Not Found.", vbExclamation, "Warning"
        txtsup = ""
        lblsup = ""
        txtsup.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtsup_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtapply.SetFocus
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtjumlah.SetFocus
    KeyAscii = 0
End Sub

Private Sub cmdsearch_Click()
txtapply = ""
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtsup = hasil
    lblsup = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select noapply,convert(char(11),tglbkt )'tglbkt', case when transtype = 'I' then 'Faktur' when transtype = 'CN' or transtype = 'DN' then 'Koreksi' end as 'type' from AM_Aropnfil where kodecust = '" & txtsup & "' and nobkt = noapply and (transtype = 'I' or transtype = 'CN' or transtype = 'DN') and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select noapply,convert(char(11),tglbkt )'tglbkt', case when transtype = 'I' then 'Faktur' when transtype = 'CN' or transtype = 'DN' then 'Koreksi' end as 'type' from AM_Aropnfil where kodecust = '" & txtsup & "' and nobkt = noapply and (transtype = 'I' or transtype = 'CN' or transtype = 'DN') and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "'"
    End If
    namatabel = "Apply to..."
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtapply = hasil
    cariapply
    hasil = ""
    hasil1 = ""
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggal1()
      tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
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
    SQL = "SELECT * FROM AM_CashHdr WHERE NoBkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        
        cmdclear_Click
    End If
    OBJ.Close
End Sub

Private Sub cariapply()
    If txtapply = "" Or txtsup = "" Or txtkurs = "" Or txtbukti = "" Then Exit Sub
    If txtapply = txtbukti Then Exit Sub
    txtketerangan = ""
    txtjumlah = 0
    txtkoreksi = 0
    
    OBJ.Open dsn
    SQL = "SELECT sum(Amount+potongan+ppn+selisih) as total from AM_Aropnfil WHERE Kodecust = '" & txtsup & "' AND Noapply = '" & txtapply & "' and kodecur = '" & txtkurs & "' and tglbkt <= '" & tanggal1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtjumlah = RST!total
        
        If lblbase = "0" Then
            SQL = "SELECT nilaikurs from AM_Aropnfil WHERE Kodecust = '" & txtsup & "' AND Noapply = '" & txtapply & "' and kodecur = '" & txtkurs & "' and nobkt = '" & txtapply & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then txtnilaikurs = RST!nilaikurs
        End If
    Else
        MsgBox "Data not found.", vbInformation, "Information"
        txtapply = ""
        txtapply.SetFocus
    End If
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
        MsgBox "Can not add, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    int3 = 0
    OBJ.Open dsn
    SQL = "select * from am_cashhdr where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        Select Case cmbtype
        Case "Credit Note"
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'CN' and nobkt like 'CN-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
            If txtbukti = txtapply Then
                If Len(str99) = 1 Then txtbukti = "CN-000" & str99
                If Len(str99) = 2 Then txtbukti = "CN-00" & str99
                If Len(str99) = 3 Then txtbukti = "CN-0" & str99
                If Len(str99) = 4 Then txtbukti = "CN-" & str99
                txtapply = txtbukti
            Else
                If Len(str99) = 1 Then txtbukti = "CN-000" & str99
                If Len(str99) = 2 Then txtbukti = "CN-00" & str99
                If Len(str99) = 3 Then txtbukti = "CN-0" & str99
                If Len(str99) = 4 Then txtbukti = "CN-" & str99
            End If
        Case "Debit Note"
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'DN' and nobkt like 'DN-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
            If txtbukti = txtapply Then
                If Len(str99) = 1 Then txtbukti = "DN-000" & str99
                If Len(str99) = 2 Then txtbukti = "DN-00" & str99
                If Len(str99) = 3 Then txtbukti = "DN-0" & str99
                If Len(str99) = 4 Then txtbukti = "DN-" & str99
                txtapply = txtbukti
            Else
                If Len(str99) = 1 Then txtbukti = "DN-000" & str99
                If Len(str99) = 2 Then txtbukti = "DN-00" & str99
                If Len(str99) = 3 Then txtbukti = "DN-0" & str99
                If Len(str99) = 4 Then txtbukti = "DN-" & str99
            End If
        Case "Retur"
            OBJ.Open dsn
            SQL = "select top 1 nobkt from am_cashhdr where kodebayar = 'CN' and nobkt like 'RT-%' order by nobkt desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobkt, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
            If txtbukti = txtapply Then
                If Len(str99) = 1 Then txtbukti = "RT-000" & str99
                If Len(str99) = 2 Then txtbukti = "RT-00" & str99
                If Len(str99) = 3 Then txtbukti = "RT-0" & str99
                If Len(str99) = 4 Then txtbukti = "RT-" & str99
                txtapply = txtbukti
            Else
                If Len(str99) = 1 Then txtbukti = "RT-000" & str99
                If Len(str99) = 2 Then txtbukti = "RT-00" & str99
                If Len(str99) = 3 Then txtbukti = "RT-0" & str99
                If Len(str99) = 4 Then txtbukti = "RT-" & str99
            End If
        End Select
    
        int3 = 1
    Else
        OBJ.Close
    End If

    
    OBJ.Open dsn
    
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    
    SQL = "select * from am_aropnfil where nobkt = '" & txtbukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        
        Exit Sub
    End If
    
    If txtbukti <> txtapply Then
        SQL = "select * from am_aropnfil where noapply = '" & txtapply & "' and tglbkt >= '" & tanggal1 & "' and (transtype = 'CN' or transtype = 'DN')" '(transtype = 'PM' or transtype = 'CN' or transtype = 'DN')"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ.Close
        
            MsgBox "Save aborted, cause already transaction with above or equal date." & vbCrLf & _
            "Please try to change correction date.", vbInformation, "Information"
            date1.SetFocus
        
            Exit Sub
        End If
    End If
    OBJ.Close
    
    If txtapply <> txtbukti Then
        OBJ.Open dsn
        SQL = "select * from am_aropnfil where noapply = '" & txtapply & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            OBJ.Close
            
            MsgBox "No Apply not found, please re-enter No Apply.", vbExclamation, "Warning"
            txtapply.SetFocus
                            
            Exit Sub
        End If
        OBJ.Close
    End If
    
    in1
    
    If int3 = 1 Then
        MsgBox "Data already exist, data was saved with next number " & txtbukti & vbCrLf & _
        "Click OK To Continue ...", vbExclamation, "Warning"
    Else
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    cmdclear_Click
End Sub

Private Sub in1()
    OBJ.Open dsn
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
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ",convert(datetime,'" & tanggalsekarang & "')"
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
    SQL = SQL + ",Nilaikurs)"

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
        Grid1.Row = 1
        Do While True
            If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
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
            SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
            SQL = SQL + ",convert(money,'" & Format(Grid1.TextMatrix(Grid1.Row, 7), "general number") & "')"
            SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 3) & "'"
            SQL = SQL + ", '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"
            SQL = SQL + ", '" & Grid1.Row & "')"
            Set RST = OBJ.Execute(SQL)
            Grid1.Row = Grid1.Row + 1
        Loop
        retur = False
    End If
    OBJ.Close
End Sub

Private Sub initGrid1()
    With Grid1
        If retur = True Then
            cmdRetur.Visible = True
            txtNomorBPB.Visible = True
            Picture1.Top = 6390
            Picture1.Left = 7200
            Me.Height = 7335
            Me.Width = 10125
            .Visible = True
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
        Else
            .Visible = False
            cmdRetur.Visible = False
            txtNomorBPB.Visible = False
            Picture1.Left = 4920
            Picture1.Top = 3960
            Me.Height = 4860
            Me.Width = 7995
        End If
    End With
    setGrid1
End Sub

Private Sub setGrid1()
    With Grid1
        If retur = True Then
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
        Else
            .Width = 7815
            .ColWidth(0) = 300
            .ColWidth(1) = 1200
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
        End If
    End With
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
        Grid1.Col = 1
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = 2
    
End Sub
