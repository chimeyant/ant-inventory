VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpurchaseorder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Purchase Order"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
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
   Icon            =   "frmpurchaseorder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2895
      Left            =   0
      TabIndex        =   32
      Top             =   1920
      Width           =   9135
      _Version        =   851970
      _ExtentX        =   16113
      _ExtentY        =   5106
      _StockProps     =   68
      AllowReorder    =   -1  'True
      Appearance      =   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "PO"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "grid"
      Item(0).Control(1)=   "txtnilai"
      Item(1).Caption =   "Nota Permintaan"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "grid2"
      Item(1).Control(1)=   "cmbstatus"
      Item(1).Control(2)=   "txtketpo"
      Begin VB.TextBox txtketpo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   37
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbstatus 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   600
         Width           =   1215
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   -69880
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Calculator      =   "frmpurchaseorder.frx":2372
         Caption         =   "frmpurchaseorder.frx":2392
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmpurchaseorder.frx":23FE
         Keys            =   "frmpurchaseorder.frx":241C
         Spin            =   "frmpurchaseorder.frx":245E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,###,##0.00;(##,###,###,##0.00);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,###,##0.00;(##,###,###,##0.00)"
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2415
         Left            =   -69880
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483632
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
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   2415
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483632
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
         _Band(0).Cols   =   8
      End
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   255
      Left            =   6600
      TabIndex        =   30
      Top             =   120
      Width           =   1935
      _Version        =   851970
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Use Request number"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "No.Permintaan"
      Height          =   615
      Left            =   6600
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
      Begin Chameleon.chameleonButton cmdnop 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "No."
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
         MICON           =   "frmpurchaseorder.frx":2486
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblnop 
         BackColor       =   &H80000014&
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox cmbkode 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtket4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtket2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   8
      Top             =   5190
      Width           =   3495
   End
   Begin VB.TextBox txtket1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   7
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   1
      Top             =   480
      Width           =   1575
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
      Left            =   3960
      Picture         =   "frmpurchaseorder.frx":27A0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   120
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
      Left            =   4200
      Picture         =   "frmpurchaseorder.frx":2AEE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3720
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   480
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
      Format          =   144506883
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmpurchaseorder.frx":2DD0
      Caption         =   "frmpurchaseorder.frx":2DF0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpurchaseorder.frx":2E5C
      Keys            =   "frmpurchaseorder.frx":2E7A
      Spin            =   "frmpurchaseorder.frx":2EBC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmpurchaseorder.frx":2EE4
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
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmpurchaseorder.frx":31FE
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
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmpurchaseorder.frx":3518
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
      Left            =   4560
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      Calculator      =   "frmpurchaseorder.frx":3832
      Caption         =   "frmpurchaseorder.frx":3852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpurchaseorder.frx":38BE
      Keys            =   "frmpurchaseorder.frx":38DC
      Spin            =   "frmpurchaseorder.frx":391E
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmpurchaseorder.frx":3946
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4560
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
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
      Format          =   144506883
      CurrentDate     =   37426
   End
   Begin VB.Label Label5 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sub Divisi"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4950
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dikirim "
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jumlah :"
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "No P.O."
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal P.O."
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   510
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6360
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "frmpurchaseorder"
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

Dim str1, str2, str3, str4, str21, str99 As String
Dim int3 As Integer
Dim posrow As String

Private Sub CheckBox1_Click()
    If CheckBox1.Value = xtpChecked Then
        lblnop = ""
        Frame1.Visible = True
    ElseIf CheckBox1.Value = xtpUnchecked Then
        Frame1.Visible = False
        hapusgrid2
    End If
End Sub

Private Sub cmbkode_Click()
    clearandfind
End Sub

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbstatus_Click()
    If cmbstatus.text = "Close" Then
        grid2.TextMatrix(grid2.Row, 4) = cmbstatus
        cmbstatus.Visible = False
        grid2.TextMatrix(grid2.Row, 5) = txtnobukti
        grid2.TextMatrix(grid2.Row, 6) = txtnobukti
        grid2.TextMatrix(grid2.Row, 7) = "1"
    Else
        grid2.TextMatrix(grid2.Row, 4) = cmbstatus
        cmbstatus.Visible = False
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = "0"
    End If
End Sub

Private Sub cmbstatus_LostFocus()
    cmbstatus.Visible = False
End Sub

Private Sub cmdadd_Click()
    If txtnobukti = "" Or txtkodecust = "" Or txtkurs = "" Or txtnilaikurs = 0 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If CheckBox1.Value = xtpChecked And lblnop = "" Then
        MsgBox "Request number is empty", vbExclamation, "Warning"
        Exit Sub
    End If

    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        hitamount

        If grid.TextMatrix(grid.Row, 4) = "" Or grid.TextMatrix(grid.Row, 7) = "0.00" Then
            MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
            Exit Sub
        End If

        grid.Row = grid.Row + 1
    Loop
    hitneto

    If txtneto = 0 Then
        MsgBox "There Is No Data To Save.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    int3 = 0
    OBJ.Open dsn
    SQL = "select nopo from am_pohdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        History
        txtnobukti = str21
        int3 = 1

        GoTo jump99
        Exit Sub
    End If
    OBJ.Close

jump99:

    OBJ.Open dsn
    SQL = "select nopo from am_pohdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        History
        txtnobukti = str21
        int3 = 1

        GoTo jump98
        Exit Sub
    End If
    OBJ.Close

jump98:

    OBJ.Open dsn
    SQL = "select * from am_pohdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        Exit Sub
    End If

    OBJ.Close

    OBJ.Open dsn
    SQL = "insert into am_pohdr ("
    SQL = SQL + "nopo, "
    SQL = SQL + "tglpo, "
    SQL = SQL + "kodesupp, "
    SQL = SQL + "tglkirim, "
    SQL = SQL + "ref, "
    SQL = SQL + "kodecur, "
    SQL = SQL + "nilaikurs, "
    SQL = SQL + "flag, "
    SQL = SQL + "ket1, "
    SQL = SQL + "ket2, "
    SQL = SQL + "ket3, "
    SQL = SQL + "ket4)"
    
    SQL = SQL + " values('" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
    SQL = SQL + "'" & txtkodecust & "',"
    SQL = SQL + "convert(datetime,'" & tanggalkirim & "'),"
    SQL = SQL + "'B',"
    SQL = SQL + "'" & txtkurs & "',"
    SQL = SQL + "convert(money,'" & txtnilaikurs & "'),"
    SQL = SQL + "'0',"
    SQL = SQL + "'" & txtket1 & "',"
    SQL = SQL + "'" & txtket2 & "',"
    SQL = SQL + "'',"
    SQL = SQL + "'" & txtket4 & "')"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                
        SQL = "insert into am_polin ("
        SQL = SQL + "nopo, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "qtybeli, "
        SQL = SQL + "price, "
        SQL = SQL + "lineitem, "
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "convert(money,'0'),"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "'),"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)

        grid.Row = grid.Row + 1
    Loop

    'CEK item permintaan apakah sudah selesai semua
    Dim np, jml As Integer
    If CheckBox1.Value = xtpChecked Then
        'simpan per baris ke am_perminin
        grid2.Row = 1
        Do While True
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            SQL = "Update am_perminin set nopo='" & grid2.TextMatrix(grid2.Row, 5) & "',"
            SQL = SQL + "nopo2='" & grid2.TextMatrix(grid2.Row, 6) & "',"
            SQL = SQL + "tglpo=convert(datetime,'" & tanggalbuat & "'),"
            If grid2.TextMatrix(grid2.Row, 4) = "Close" Then
                SQL = SQL + "status='1' Where lineitem='" & grid2.TextMatrix(grid2.Row, 0) & "'"
            ElseIf grid2.TextMatrix(grid2.Row, 4) = "Pending" Then
                SQL = SQL + "status='0' Where lineitem='" & grid2.TextMatrix(grid2.Row, 0) & "'"
            End If
            SQL = SQL + " and nobkt='" & lblnop & "'"
            Set RST = OBJ.Execute(SQL)
            grid2.Row = grid2.Row + 1
        Loop
        
        SQL = "Select COUNT(nobkt)'tot' from am_perminin Where nobkt = '" & lblnop & "'"
        Set RST = OBJ.Execute(SQL)
        np = RST!tot
    
        SQL = "Select SUM(cast(status as int))'jumlah' from am_perminin Where nobkt = '" & lblnop & "'"
        Set RST = OBJ.Execute(SQL)
        jml = RST!jumlah
        
        'jika Nota selesai semua close NP
        If np = jml Then
            SQL = "Update am_perminapp set nopo='" & txtnobukti & "',"
            SQL = SQL + "kodesupp='" & txtkodecust & "',"
            SQL = SQL + "status='1',tglpo=convert(datetime,'" & tanggalpo & "'),"
            SQL = SQL + "tgl=convert(datetime,'" & tanggalpo & "') Where nobkt='" & lblnop & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "Update am_perminhdr set flag='1' Where nobkt='" & lblnop & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        
    End If
    OBJ.Close
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='75' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If Not RST.EOF Then
    '        OBJ.Close
            
    '        frmpurchaseordershow.Show vbModal
    '    Else
    '        OBJ.Close
    '    End If
    'Else
        frmpurchaseordershow.Show vbModal
    'End If

    If int3 = 1 Then
        MsgBox "Data already exist, data was saved with next number " & txtnobukti & vbCrLf & _
        "Click OK To Continue ...", vbExclamation, "Warning"
    Else
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    clearandfind
End Sub

Private Sub cmdclear_Click()
    clearandfind
    hapusgrid2
    txtnobukti = ""
    cmbkode = ""
    lblnop = ""
    txtketpo = ""
    cmbstatus.Visible = False
    txtketpo.Visible = False
    CheckBox1.Value = xtpUnchecked
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdnop_Click()
    carisql1 = "Select nobkt,pemesan From am_perminhdr"
    namatabel = "Permintaan"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnop_GotFocus()
    If hasil = "" Then Exit Sub
    lblnop = hasil
    hasil = ""
    opennota
End Sub
Private Sub opennota()
    hapusgrid2
    OBJ.Open dsn

    'SQL = "select * from am_perminin where nobkt = '" & lblnop & "' and (status is null or status = '0') Order By lineitem asc"
    SQL = "Select a.*,b.namasatuan From am_perminin a left join am_apunit b on a.kdsatuan=b.kodesatuan"
    SQL = SQL + " Where a.nobkt='" & lblnop & "' and (a.status is null or a.status='0') order by a.lineitem asc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        grid2.Row = 1
        Do Until RST.EOF
            grid2.Col = 1
            grid2.CellAlignment = 1
            grid2.TextMatrix(grid2.Row, 0) = RST!lineitem
            grid2.TextMatrix(grid2.Row, 1) = RST!nmbrg
            grid2.TextMatrix(grid2.Row, 2) = Format(RST!qty, "###,###,##0.00")
            If IsNull(RST!namasatuan) Or RST!namasatuan = "" Then
                grid2.TextMatrix(grid2.Row, 3) = ""
            Else
                grid2.TextMatrix(grid2.Row, 3) = RST!namasatuan
            End If
            If IsNull(RST!Status) Or RST!Status = "0" Then  '0 = PENDING : 1 = CLOSE
                grid2.TextMatrix(grid2.Row, 4) = "Pending"
            Else
                grid2.TextMatrix(grid2.Row, 4) = "Close"
            End If
            If IsNull(RST!nopo2) Or RST!nopo2 = "" Then
                grid2.TextMatrix(grid2.Row, 5) = ""
                grid2.TextMatrix(grid2.Row, 6) = ""
            Else
                grid2.TextMatrix(grid2.Row, 5) = RST!nopo
                grid2.TextMatrix(grid2.Row, 6) = RST!nopo2
            End If
            grid2.TextMatrix(grid2.Row, 7) = "0"
            'SetRow2 grid2.Row, True
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub
Private Sub cmdsearch_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs = hasil
    carikurs
    hasil = ""
    
    hapusgrid
    txtkodecust = ""
    lblnamacust = ""
End Sub

Private Sub date1_Change()
    History
    txtnobukti = str21
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
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='111' and b.kodeuser = '2" & kuser & "'"
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
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Qty"
    grid.TextMatrix(0, 5) = "Satuan"
    grid.TextMatrix(0, 6) = "Harga"
    grid.TextMatrix(0, 7) = "Total"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 800
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1200
    grid.ColWidth(7) = 1500
    grid.RowHeightMin = 300
    
    grid2.TextMatrix(0, 1) = "Nama Barang"
    grid2.TextMatrix(0, 2) = "Qty"
    grid2.TextMatrix(0, 3) = "Satuan"
    grid2.TextMatrix(0, 4) = "Status"
    grid2.TextMatrix(0, 5) = "No.PO"
    grid2.TextMatrix(0, 6) = "No.PO"
    grid2.ColWidth(0) = 250
    grid2.ColWidth(1) = 3200
    grid2.ColWidth(2) = 1000
    grid2.ColWidth(3) = 900
    grid2.ColWidth(4) = 1000
    grid2.ColWidth(5) = 0
    grid2.ColWidth(6) = 2000
    grid2.ColWidth(7) = 0
    
    date1.Value = Date
    date2.Value = Date
    txtnilai.Visible = False
    txtketpo.Visible = False
    cmbstatus.Visible = False
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str2 = RST!kode1
        str4 = RST!kode2
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1.MinDate = RST!tanggal1
        date1.MaxDate = RST!tanggal2
    End If
    OBJ.Close
    
    cmbstatus.AddItem "Pending"
    cmbstatus.AddItem "Close"
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If cmbkode = "" Or txtnobukti = "" Or txtkurs = "" Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            'If grid.Rows > 7 Then Exit Sub

            If cmbkode = "BB/L" Or cmbkode = "KM/L" Then
                If grid.Row = 1 Then
                    carisql1 = "select a.kodebarang,b.namabarang,a.kodesatuan,a.price,convert(char(11),a.lastupdate)lastupdate,a.kodesupp,c.namasupp from am_price a left join am_apitemmst b on a.kodebarang=b.kodebarang left join am_supplier c on a.kodesupp=c.kodesupp where a.kodecurr = '" & txtkurs & "' and (b.kodeproduk = 'BB/L' or b.kodeproduk = 'KM/L')"
                Else
                    carisql1 = "select a.kodebarang,b.namabarang,a.kodesatuan,a.price,convert(char(11),a.lastupdate)lastupdate,a.kodesupp,c.namasupp from am_price a left join am_apitemmst b on a.kodebarang=b.kodebarang left join am_supplier c on a.kodesupp=c.kodesupp where a.kodesupp = '" & txtkodecust & "' and a.kodecurr = '" & txtkurs & "' and (b.kodeproduk = 'BB/L' or b.kodeproduk = 'KM/L')"
                End If
            Else
                If grid.Row = 1 Then
                    carisql1 = "select a.kodebarang,b.namabarang,a.kodesatuan,a.price,convert(char(11),a.lastupdate)lastupdate,a.kodesupp,c.namasupp from am_price a left join am_apitemmst b on a.kodebarang=b.kodebarang left join am_supplier c on a.kodesupp=c.kodesupp where a.kodecurr = '" & txtkurs & "' and b.kodeproduk = '" & cmbkode & "'"
                Else
                    carisql1 = "select a.kodebarang,b.namabarang,a.kodesatuan,a.price,convert(char(11),a.lastupdate)lastupdate,a.kodesupp,c.namasupp from am_price a left join am_apitemmst b on a.kodebarang=b.kodebarang left join am_supplier c on a.kodesupp=c.kodesupp where a.kodesupp = '" & txtkodecust & "' and a.kodecurr = '" & txtkurs & "' and b.kodeproduk = '" & cmbkode & "'"
                End If
            End If
            namatabel = "Barang per Divisi"
            
            frmsearch.Show vbModal
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecust = "" Then Exit Sub

            txtnilai.Format = "##,###,###,##0.00;(##,###,###,##0.00)"
            txtnilai.DisplayFormat = "##,###,###,##0.00;(##,###,###,##0.00);0"
            txtnilai.Value = 0
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6
            If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecust = "" Then Exit Sub

            txtnilai.Format = "##,###,###,##0.0000;(##,###,###,##0.0000)"
            txtnilai.DisplayFormat = "##,###,###,##0.0000;(##,###,###,##0.0000);0"
            txtnilai.Value = 0
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtkurs = "" Then Exit Sub
    
    Select Case grid.Col
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecust = "" Then Exit Sub

        posrow = grid.Row

        txtnilai.Format = "##,###,###,##0.00;(##,###,###,##0.00)"
        txtnilai.DisplayFormat = "##,###,###,##0.00;(##,###,###,##0.00);0"
        txtnilai.Value = 0
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 6
        If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecust = "" Then Exit Sub

        posrow = grid.Row

        txtnilai.Format = "##,###,###,##0.0000;(##,###,###,##0.0000)"
        txtnilai.DisplayFormat = "##,###,###,##0.0000;(##,###,###,##0.0000);0"
        txtnilai.Value = 0
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
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 1) = hasil Then
                MsgBox "Item already exist.", vbInformation, "Information"
                hasil = ""
                hasil1 = ""
                hasil2 = ""
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop

        grid.Row = posrow
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = hasil
        txtkodecust = hasil2
        hasil = ""
        hasil1 = ""
        hasil2 = ""

        OBJ.Open dsn
        SQL = "select b.namabarang,b.kodesatuan,c.namasatuan,isnull(a.price,0)'price' from am_price a right join am_apitemmst b on a.kodebarang=b.kodebarang left join am_apunit c on b.kodesatuan=c.kodesatuan where a.kodecurr = '" & txtkurs & "' and b.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesupp = '" & txtkodecust & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = "0.00"
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            grid.TextMatrix(grid.Row, 6) = Format(RST!Price, "###,###,###,##0.0000")
            grid.TextMatrix(grid.Row, 7) = "0.00"

            SQL = "select namasupp from am_supplier where kodesupp = '" & txtkodecust & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then lblnamacust = RST!namasupp Else lblnamacust = ""

            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1

            SetRow grid.Row, True
            grid.SetFocus
            grid.Col = 2
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
            txtkodecust = ""
        End If
        OBJ.Close
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    Select Case grid2.Col
        Case 4:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            cmbstatus.Width = grid2.ColWidth(grid2.Col) - 40
            cmbstatus = grid2.TextMatrix(grid2.Row, grid2.Col)
            cmbstatus.Left = grid2.Left + grid2.CellLeft
            cmbstatus.Top = grid2.Top + grid2.CellTop + 20
            cmbstatus.Visible = True
            cmbstatus.SetFocus
        Case 6:
            txtketpo.Width = grid2.ColWidth(grid2.Col) - 40
            txtketpo = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtketpo.Left = grid2.Left + grid2.CellLeft
            txtketpo.Top = grid2.Top + grid2.CellTop + 20
            txtketpo.Visible = True
            txtketpo.SetFocus
    End Select
End Sub

Private Sub grid2_EnterCell()
    If grid2.MouseRow = 0 Then Exit Sub
    
    Select Case grid2.Col
        Case 4:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            cmbstatus.Width = grid2.ColWidth(grid2.Col) - 40
            cmbstatus = grid2.TextMatrix(grid2.Row, grid2.Col)
            cmbstatus.Left = grid2.Left + grid2.CellLeft
            cmbstatus.Top = grid2.Top + grid2.CellTop + 20
            cmbstatus.Visible = True
            cmbstatus.SetFocus
        Case 6:
            txtketpo.Width = grid2.ColWidth(grid2.Col) - 40
            txtketpo = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtketpo.Left = grid2.Left + grid2.CellLeft
            txtketpo.Top = grid2.Top + grid2.CellTop + 20
            txtketpo.Visible = True
            txtketpo.SetFocus
    End Select
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case TabControl1.SelectedItem
        Case 0: txtnilai.Visible = False
        Case 1: cmbstatus.Visible = False
                txtketpo.Visible = False
    End Select
End Sub

Private Sub txtket1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtket2.SetFocus
End Sub

Private Sub txtket2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdadd.SetFocus
End Sub

Private Sub txtket4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtketpo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    posrow = grid2.Row
    Select Case grid2.Col
        Case 6:
            If txtketpo = "" Then
                txtketpo.Visible = False
                grid2.SetFocus
                grid2.Row = posrow
                Exit Sub
            End If
            If Len(Trim(txtketpo)) > 100 Then
                MsgBox "Max.100 karakter", vbExclamation, AppName
                Exit Sub
            End If
            grid2.TextMatrix(grid2.Row, 6) = txtketpo.text
            If grid2.Row = (grid2.Rows - 1) Then grid2.Rows = grid2.Rows + 1
            grid2.SetFocus
    End Select
        txtketpo = ""
        txtketpo.Visible = False
        grid2.SetFocus
        grid2.Row = posrow
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
        Select Case grid.Col
            Case 3
                grid.TextMatrix(grid.Row, 3) = Format(txtnilai, "###,###,###,##0.00")
            Case 6
                grid.TextMatrix(grid.Row, 6) = Format(txtnilai, "###,###,###,##0.0000")
        End Select
        
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        hitamount
        hitneto
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai_LostFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnilaikurs_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblbase = "1" Then KeyCode = 0
End Sub

Private Sub txtnilaikurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtket4.SetFocus
    If lblbase = "1" Then KeyAscii = 0
End Sub

Private Sub History()
    OBJ1.Open dsn
    SQL1 = "select top 1 nopo from am_pohdr where nopo like '" & Format(date1, str2) & "%" & cmbkode & "' order by nopo desc"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        str99 = Mid(RST1!nopo, Len(str2) + 1, 3)
    Else
        str99 = 0
    End If
    
    str99 = str99 + 1
    
    If Len(str99) = 1 Then str21 = Format(date1, str2) & "00" & str99 & str4 & cmbkode
    If Len(str99) = 2 Then str21 = Format(date1, str2) & "0" & str99 & str4 & cmbkode
    If Len(str99) = 3 Then str21 = Format(date1, str2) & str99 & str4 & cmbkode
        
    OBJ1.Close
End Sub

Private Sub hitamount()
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
    str3 = grid.TextMatrix(grid.Row, 6) * grid.TextMatrix(grid.Row, 3)
    grid.TextMatrix(grid.Row, 7) = Format(str3, "###,###,###,##0.00")
End Sub

Private Sub hitneto()
    grid.Row = 1
    str1 = 0
    Do While True
        If grid.Rows = 2 Then Exit Do

        str1 = Val(str1) + (Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) * Val(Format(grid.TextMatrix(grid.Row, 6), "general number")))

        If grid.TextMatrix(grid.Row + 1, 1) = "" Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    txtneto.Value = str1
    If txtneto = 0 Then Exit Sub
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub hapusgrid2()
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
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
End Sub
Private Sub hapusgrid()
    txtneto = 0

    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 800
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1200
    grid.ColWidth(7) = 1500
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub clearandfind()
    hapusgrid
    hapusgrid2
    txtnobukti = ""
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    date2.Value = Date
    txtkodecust = ""
    lblnamacust = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtket1 = ""
    txtket2 = ""
    txtket4 = ""
    
    date1.SetFocus
    History
    txtnobukti = str21
End Sub

Private Sub carikurs()
    If txtkurs = "" Then Exit Sub

    OBJ2.Open dsn
    SQL2 = "select * from gl_kurs where kdkurs = '" & txtkurs & "'"
    Set RST2 = OBJ2.Execute(SQL2)
    If Not RST2.EOF Then
        lblbase = RST2!base

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
    
    hapusgrid
End Sub
Function tanggalbuat()
    tanggalbuat = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
Function tanggalpo()
      tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalkirim()
      tanggalkirim = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
