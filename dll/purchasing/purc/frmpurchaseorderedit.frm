VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpurchaseorderedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Purchase Order"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
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
   Icon            =   "frmpurchaseorderedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "No.Permintaan"
      Height          =   615
      Left            =   6600
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
      Begin Chameleon.chameleonButton cmdnop 
         Height          =   285
         Left            =   120
         TabIndex        =   33
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
         MICON           =   "frmpurchaseorderedit.frx":2372
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
         TabIndex        =   34
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
      TabIndex        =   10
      Top             =   4470
      Width           =   3495
   End
   Begin VB.TextBox txtket1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   9
      Top             =   4200
      Width           =   3495
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmpurchaseorderedit.frx":268C
      Caption         =   "frmpurchaseorderedit.frx":26AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpurchaseorderedit.frx":2718
      Keys            =   "frmpurchaseorderedit.frx":2736
      Spin            =   "frmpurchaseorderedit.frx":2778
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
      Left            =   4560
      Picture         =   "frmpurchaseorderedit.frx":27A0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
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
      Left            =   4800
      Picture         =   "frmpurchaseorderedit.frx":2AEE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
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
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
   End
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   7080
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmpurchaseorderedit.frx":2DD0
      Caption         =   "frmpurchaseorderedit.frx":2DF0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpurchaseorderedit.frx":2E5C
      Keys            =   "frmpurchaseorderedit.frx":2E7A
      Spin            =   "frmpurchaseorderedit.frx":2EBC
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
      TabIndex        =   14
      Top             =   4200
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
      MICON           =   "frmpurchaseorderedit.frx":2EE4
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
      TabIndex        =   13
      Top             =   4200
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
      MICON           =   "frmpurchaseorderedit.frx":31FE
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
      Left            =   4920
      TabIndex        =   11
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmpurchaseorderedit.frx":3518
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3600
      TabIndex        =   15
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No P.O."
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
      MICON           =   "frmpurchaseorderedit.frx":3832
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdel 
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmpurchaseorderedit.frx":3B4C
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
      Calculator      =   "frmpurchaseorderedit.frx":3E66
      Caption         =   "frmpurchaseorderedit.frx":3E86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpurchaseorderedit.frx":3EF2
      Keys            =   "frmpurchaseorderedit.frx":3F10
      Spin            =   "frmpurchaseorderedit.frx":3F52
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
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   240
      TabIndex        =   27
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
      MICON           =   "frmpurchaseorderedit.frx":3F7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   255
      Left            =   6600
      TabIndex        =   31
      Top             =   120
      Width           =   1935
      _Version        =   851970
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Use Request number"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Nilai Kurs"
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lblbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Sub Divisi"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dikirim"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jumlah :"
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   870
      Width           =   4695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal P.O."
      Height          =   255
      Left            =   3360
      TabIndex        =   20
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
Attribute VB_Name = "frmpurchaseorderedit"
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

Dim str1, str2, str3 As String
Dim posrow As String
Dim int79 As Integer

Private Sub CheckBox1_Click()
    If CheckBox1.Value = xtpChecked Then
        lblnop = ""
        Frame1.Visible = True
    ElseIf CheckBox1.Value = xtpUnchecked Then
        Frame1.Visible = False
    End If
End Sub

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    
    If MsgBox("Update nilai saja ?", vbQuestion + vbYesNo, "Question") = vbYes Then
        grid.Row = 1
        Do While True
            If grid.Rows = grid.Row + 1 Then Exit Do
            hitamount
            
            If grid.TextMatrix(grid.Row, 4) = "" Or grid.TextMatrix(grid.Row, 7) = "0.00" Then
                MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
                Exit Sub
            End If
            
            OBJ.Open dsn
            SQL = "select * from am_polin where nopo='" & txtnobukti & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                OBJ.Close
                MsgBox "Data already change, user can not change price.", vbExclamation, "Warning"
                Exit Sub
            End If
            OBJ.Close
            
            grid.Row = grid.Row + 1
        Loop
        
        OBJ.Open dsn
        SQL = "select * from am_polin where nopo='" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            int79 = 0
            grid.Row = 1
            Do While True
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                
                If grid.TextMatrix(grid.Row, 1) = RST!kodebarang Then
                    int79 = 1
                    Exit Do
                End If
                
                grid.Row = grid.Row + 1
            Loop
            
            If int79 = 0 Then
                OBJ.Close
                MsgBox "Data already change, user can not change price.", vbExclamation, "Warning"
                Exit Sub
            End If
            
            RST.MoveNext
        Loop
        OBJ.Close
        
        hitneto
        
        If txtneto = 0 Then
            MsgBox "There Is No Data To Save.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If MsgBox("Are you sure want to update price ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            cmdclear_Click
            Exit Sub
        End If
        
        OBJ.Open dsn
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                    
            SQL = "update am_polin set "
            SQL = SQL + "price = convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "')"
            SQL = SQL + " where nopo = '" & txtnobukti & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "update am_beliapp set "
            SQL = SQL + "kodecur = '" & txtkurs & "',nilaikurs = '" & txtnilaikurs & "',"
            SQL = SQL + "qty = convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
            SQL = SQL + "price = convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "')"
            SQL = SQL + " where nopo = '" & txtnobukti & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            grid.Row = grid.Row + 1
        Loop
        OBJ.Close
    Else
        OBJ.Open dsn
        SQL = "select * from am_period where tanggal1 <= '" & tanggalpo & "' and tanggal2 >= '" & tanggalpo & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            OBJ.Close
            MsgBox "Can not update, Data already close.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        SQL = "select * from am_pohdr where nopo = '" & txtnobukti & "' and flag = '1'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ.Close
            MsgBox "Can not update, PO Cancel/Close.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        'SQL = "select * from am_pohdr where nopo = '" & txtnobukti & "' and ref = 'P'"
        'Set RST = OBJ.Execute(SQL)
        'If Not RST.EOF Then
        '    OBJ.Close
        '    MsgBox "Can not update, PO already printed out.", vbExclamation, "Warning"
        '    Exit Sub
        'End If
        OBJ.Close
            
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
        
        If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            cmdclear_Click
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "select * from am_belihdr where nopo = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ.Close
            MsgBox "Can not update, record in use.", vbExclamation, "Warning"
            Exit Sub
        End If
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "select * from am_pohdr where ket3 = '' and nopo = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Is this PO REVISI ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                str2 = "Revisi"
            Else
                str2 = ""
            End If
        Else
            str2 = ""
        End If
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "delete from am_polin where nopo = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_pohdr set "
        SQL = SQL + "kodesupp = '" & txtkodecust & "', "
        SQL = SQL + "kodecur = '" & txtkurs & "', "
        SQL = SQL + "nilaikurs = convert(money,'" & txtnilaikurs & "'), "
        SQL = SQL + "ket1 = '" & txtket1 & "', "
        SQL = SQL + "ket2 = '" & txtket2 & "', "
        If str2 = "Revisi" Then SQL = SQL + "ket3 = 'PO REVISI', "
        SQL = SQL + "ket4 = '" & txtket4 & "'"
        SQL = SQL + " where nopo = '" & txtnobukti & "'"
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
        OBJ.Close
    End If
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    clearandfind
End Sub

Private Sub cmdclear_Click()
    clearandfind
    cmbkode = ""
    cmbkode.SetFocus
    lblnop = ""
    CheckBox1.Value = xtpUnchecked
    Frame1.Visible = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel_Click()
    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complte", vbExclamation, "Warning"
        grid.SetFocus
        grid.Row = 1
        grid.Col = 1
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalpo & "' and tanggal2 >= '" & tanggalpo & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, Data already close.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_pohdr where nopo = '" & txtnobukti & "' and flag = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not delete, PO Cancel/Close.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Close
   
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_belihdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Can not update, record in use.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "delete am_pohdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete am_polin where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "Update am_perminapp set nopo='',kdsupp='',status='0' Where nopo='" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    clearandfind
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
    'opennota
End Sub

Private Sub cmdsearch_Click()
    If cmbkode = "" Then Exit Sub
    
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nopo, convert(char(11),tglpo)'tglpo' from am_pohdr where nopo like '%" & cmbkode & "' and tglpo >= '" & batas1 & "' and tglpo <= '" & batas2 & "'"
    Else
        carisql1 = "select nopo, convert(char(11),tglpo)'tglpo' from am_pohdr where nopo like '%" & cmbkode & "'"
    End If
    namatabel = "Purchase Order"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtnobukti = hasil
    hasil = ""
    hasil1 = ""
    
    carinvoice
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs = hasil
    carikurs
    hasil = ""

    hapusgrid
    txtkodecust = ""
    lblnamacust = ""
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='112' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='113' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdel.Enabled = False
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
    
    date1.Value = Date
    date2.Value = Date
    
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtkurs = "" Then Exit Sub
    
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
            If grid.Rows > 7 Then Exit Sub
            
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
            
            txtnilai.Format = "##,###,##0.00;(##,###,##0.00)"
            txtnilai.DisplayFormat = "##,###,##0.00;(##,###,##0.00);0"
            txtnilai.Value = 0
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6
            If grid.TextMatrix(grid.Row, 1) = "" Or txtkodecust = "" Then Exit Sub
            
            txtnilai.Format = "##,###,##0.0000;(##,###,##0.0000)"
            txtnilai.DisplayFormat = "##,###,##0.0000;(##,###,##0.0000);0"
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
        
        txtnilai.Format = "##,###,##0.00;(##,###,##0.00)"
        txtnilai.DisplayFormat = "##,###,##0.00;(##,###,##0.00);0"
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
        
        txtnilai.Format = "##,###,##0.0000;(##,###,##0.0000)"
        txtnilai.DisplayFormat = "##,###,##0.0000;(##,###,##0.0000);0"
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
        SQL = "select b.namabarang,b.kodesatuan,c.namasatuan,isnull(a.price,0)'price' from am_price a right join am_apitemmst b on a.kodebarang=b.kodebarang left join am_apunit c on b.kodesatuan=c.kodesatuan where a.kodecurr = '" & txtkurs & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesupp = '" & txtkodecust & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = "0.00"
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            grid.TextMatrix(grid.Row, 6) = Format(RST!Price, "###,###,##0.0000")
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

Private Sub txtket1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtket2.SetFocus
End Sub

Private Sub txtket2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdadd.SetFocus
End Sub

Private Sub txtket4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
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
                grid.TextMatrix(grid.Row, 3) = Format(txtnilai, "###,###,##0.00")
            Case 6
                grid.TextMatrix(grid.Row, 6) = Format(txtnilai, "###,###,##0.0000")
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

Function tanggalpo()
      tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalkirim()
      tanggalkirim = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub clearandfind()
    hapusgrid
    
    cmdsearch.Enabled = True
    date1.Enabled = True
    cmbkode.Enabled = True
    
    txtnobukti = ""
    date1.Value = Date
    date2.Value = Date
    txtkodecust = ""
    lblnamacust = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtket1 = ""
    txtket2 = ""
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

Private Sub hitamount()
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
    str3 = grid.TextMatrix(grid.Row, 6) * grid.TextMatrix(grid.Row, 3)
    grid.TextMatrix(grid.Row, 7) = Format(str3, "###,###,##0.00")
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

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub

    hapusgrid
    
    date1 = Date
    date2 = Date
    txtkodecust = ""
    lblnamacust = ""
    txtkurs = ""
    txtnilaikurs = 0
    txtket1 = ""
    txtket2 = ""
    txtket4 = ""
    
    OBJ.Open dsn
    SQL = "select * from am_pohdr where nopo = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglpo
        date2 = RST!tglkirim
        txtkodecust = RST!kodesupp
        txtket1 = RST!ket1
        txtket2 = RST!ket2
        txtket4 = RST!ket4
        txtkurs = RST!kodecur
        txtnilaikurs = RST!nilaikurs
        
        SQL = "select * from am_supplier where kodesupp = '" & txtkodecust & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblnamacust = RST!namasupp
        
        grid.Row = 1
        SQL = "select * from am_polin where nopo = '" & txtnobukti & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 6) = Format(RST!Price, "###,###,##0.0000")
            
            OBJ1.Open dsn
            SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 2) = RST1!namabarang
            
            SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            OBJ1.Close
        
            SetRow grid.Row, True
            hitamount
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
        hitneto
        
        SQL = "Select nobkt From am_perminapp Where nopo='" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblnop = RST!nobkt
            CheckBox1.Value = xtpChecked
            Frame1.Visible = True
        End If
        
        cmdsearch.Enabled = False
        date1.Enabled = False
        cmbkode.Enabled = False
        grid.SetFocus
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        cmdsearch.SetFocus
    End If
    OBJ.Close
End Sub


Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function
