VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmaddsop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah / Ubah SOP"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10920
      Top             =   960
   End
   Begin VB.TextBox txtnobpb 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10080
      TabIndex        =   58
      Top             =   690
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   3105
      Picture         =   "frmAddSOP.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   7680
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
      Left            =   3360
      Picture         =   "frmAddSOP.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   7680
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
      TabIndex        =   11
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produk"
      Height          =   1410
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   11340
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KEMASAN"
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
         Left            =   8760
         TabIndex        =   60
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BAHAN BAKU + PEROLEHAN"
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
         Left            =   8760
         TabIndex        =   59
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SELESAI"
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
         Left            =   8760
         TabIndex        =   57
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtnoreaktor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         TabIndex        =   53
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtproduk 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2235
         TabIndex        =   10
         Top             =   225
         Width           =   3060
      End
      Begin VB.TextBox txtnolot 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1155
         TabIndex        =   4
         Top             =   570
         Width           =   4140
      End
      Begin VB.TextBox txtkodeproduk 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   225
         Width           =   1050
      End
      Begin XtremeSuiteControls.PushButton cmdproduksi 
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   990
         _Version        =   851970
         _ExtentX        =   1746
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "PRODUK :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker datebahan 
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Top             =   945
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   1
      End
      Begin XtremeSuiteControls.PushButton cmdnolot 
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   600
         Width           =   990
         _Version        =   851970
         _ExtentX        =   1746
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "NO LOT :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker datedone 
         Height          =   315
         Left            =   7095
         TabIndex        =   56
         Top             =   600
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   1
      End
      Begin VB.Label lbleditmode 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "PERHATIAN : MODE UBAH/TAMBAH AKTIF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   65
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "TANGGAL SELESAI  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   55
         Top             =   630
         Width           =   1770
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "NO KOCEKAN :"
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
         Left            =   5745
         TabIndex        =   54
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "TANGGAL  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   2
         Top             =   990
         Width           =   960
      End
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   600
      Left            =   10185
      TabIndex        =   6
      Top             =   7545
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   1058
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAddSOP.frx":0630
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   600
      Left            =   7560
      TabIndex        =   7
      Top             =   7545
      Width           =   1275
      _Version        =   851970
      _ExtentX        =   2249
      _ExtentY        =   1058
      _StockProps     =   79
      Caption         =   "SIMPAN"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnDelete 
      Height          =   600
      Left            =   8880
      TabIndex        =   8
      Top             =   7545
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   1058
      _StockProps     =   79
      Caption         =   "HAPUS"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnNew 
      Height          =   600
      Left            =   6225
      TabIndex        =   9
      Top             =   7545
      Width           =   1305
      _Version        =   851970
      _ExtentX        =   2302
      _ExtentY        =   1058
      _StockProps     =   79
      Caption         =   "BARU"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5580
      Left            =   105
      TabIndex        =   15
      Top             =   1905
      Width           =   11370
      _Version        =   851970
      _ExtentX        =   20055
      _ExtentY        =   9842
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemCount       =   6
      SelectedItem    =   5
      Item(0).Caption =   "Pemakaian Bahan Baku"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Pemakaian Bahan Baku Tambahan"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Waktu Produksi"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "La. QC"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
      Item(4).Caption =   "Perolehan Produksi"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "grid4"
      Item(4).Control(1)=   "txtnilai4"
      Item(4).Control(2)=   "Label14"
      Item(5).Caption =   "Pemakaian Kaleng dan Karton"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "TabControlPage5"
      Begin TDBNumber6Ctl.TDBNumber txtnilai4 
         Height          =   255
         Left            =   -59890
         TabIndex        =   51
         Top             =   4890
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calculator      =   "frmAddSOP.frx":0F0A
         Caption         =   "frmAddSOP.frx":0F2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAddSOP.frx":0F96
         Keys            =   "frmAddSOP.frx":0FB4
         Spin            =   "frmAddSOP.frx":0FF6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0.000;(###,###,###,##0.000);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.000;(###,###,###,##0.000)"
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   5220
         Left            =   30
         TabIndex        =   16
         Top             =   330
         Width           =   11310
         _Version        =   851970
         _ExtentX        =   19950
         _ExtentY        =   9208
         _StockProps     =   1
         Page            =   7
         Begin TDBNumber6Ctl.TDBNumber txtnilai3 
            Height          =   255
            Left            =   9885
            TabIndex        =   17
            Top             =   4740
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            Calculator      =   "frmAddSOP.frx":101E
            Caption         =   "frmAddSOP.frx":103E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAddSOP.frx":10AA
            Keys            =   "frmAddSOP.frx":10C8
            Spin            =   "frmAddSOP.frx":110A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.000;(###,###,###,##0.000);0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.000;(###,###,###,##0.000)"
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
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
            Height          =   4620
            Left            =   0
            TabIndex        =   18
            Top             =   465
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   8149
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "*  Untuk pemakaian kemasan gunakan tanggal penambahan bahan baku dan kemasan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3120
            TabIndex        =   70
            Top             =   120
            Width           =   8010
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PEMAKAIAN KALENG DAN KARTON :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   19
            Top             =   165
            Width           =   3015
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5220
         Left            =   -69970
         TabIndex        =   20
         Top             =   330
         Visible         =   0   'False
         Width           =   11310
         _Version        =   851970
         _ExtentX        =   19950
         _ExtentY        =   9208
         _StockProps     =   1
         Page            =   2
         Begin VB.TextBox txtwaktupelarutan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   195
            Width           =   900
         End
         Begin VB.TextBox txtwaktutambahan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtwaktukemasan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   915
            Width           =   900
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Waktu Pelarutan :"
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
            Left            =   345
            TabIndex        =   26
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Waktu Tambahan :"
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
            Left            =   240
            TabIndex        =   25
            Top             =   585
            Width           =   1500
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Waktu Pengemasan :"
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
            Left            =   120
            TabIndex        =   24
            Top             =   945
            Width           =   1635
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5220
         Left            =   -69970
         TabIndex        =   27
         Top             =   330
         Visible         =   0   'False
         Width           =   11310
         _Version        =   851970
         _ExtentX        =   19950
         _ExtentY        =   9208
         _StockProps     =   1
         Page            =   1
         Begin VB.TextBox txtnolot2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   390
            TabIndex        =   29
            Top             =   4245
            Visible         =   0   'False
            Width           =   2010
         End
         Begin TDBNumber6Ctl.TDBNumber txtnilai1 
            Height          =   255
            Left            =   9960
            TabIndex        =   28
            Top             =   4635
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            Calculator      =   "frmAddSOP.frx":1132
            Caption         =   "frmAddSOP.frx":1152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAddSOP.frx":11BE
            Keys            =   "frmAddSOP.frx":11DC
            Spin            =   "frmAddSOP.frx":121E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.000;(###,###,###,##0.000);0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.000;(###,###,###,##0.000)"
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
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber txttotalproduksi 
            Height          =   315
            Left            =   1830
            TabIndex        =   30
            Top             =   4740
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   556
            Calculator      =   "frmAddSOP.frx":1246
            Caption         =   "frmAddSOP.frx":1266
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAddSOP.frx":12D2
            Keys            =   "frmAddSOP.frx":12F0
            Spin            =   "frmAddSOP.frx":133A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999999
            MinValue        =   -9999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
            Height          =   4110
            Left            =   45
            TabIndex        =   31
            Top             =   450
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   7250
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin TDBNumber6Ctl.TDBNumber txttotalhasilproduksi 
            Height          =   315
            Left            =   5640
            TabIndex        =   32
            Top             =   4725
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   556
            Calculator      =   "frmAddSOP.frx":1362
            Caption         =   "frmAddSOP.frx":1382
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAddSOP.frx":13EE
            Keys            =   "frmAddSOP.frx":140C
            Spin            =   "frmAddSOP.frx":1456
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "*  Untuk pemakaian kemasan gunakan tanggal penambahan bahan baku dan kemasan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3120
            TabIndex        =   73
            Top             =   120
            Width           =   8010
         End
         Begin VB.Label tg2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   255
            Left            =   8385
            TabIndex        =   62
            Top             =   4740
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL PRODUKSI :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   4785
            Width           =   1560
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TAMBAHAN PEMAKAIAN BAHAN BAKU :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   34
            Top             =   195
            Width           =   3180
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL HASIL PRODUKSI :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   33
            Top             =   4785
            Width           =   2370
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5220
         Left            =   -69970
         TabIndex        =   36
         Top             =   330
         Visible         =   0   'False
         Width           =   11310
         _Version        =   851970
         _ExtentX        =   19950
         _ExtentY        =   9208
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.GroupBox gbbundle 
            Height          =   495
            Left            =   3840
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin TDBNumber6Ctl.TDBNumber txtbundle 
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   160
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   450
               Calculator      =   "frmAddSOP.frx":147E
               Caption         =   "frmAddSOP.frx":149E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmAddSOP.frx":150A
               Keys            =   "frmAddSOP.frx":1528
               Spin            =   "frmAddSOP.frx":156A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0;(###,###,###,##0)"
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
               ValueVT         =   1638405
               Value           =   1
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Bundle"
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
               Left            =   960
               TabIndex        =   69
               Top             =   160
               Width           =   615
            End
         End
         Begin VB.CheckBox chkWIP 
            Caption         =   "BASE WIP"
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
            Left            =   2280
            TabIndex        =   66
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txtnolot1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   375
            TabIndex        =   37
            Top             =   4740
            Visible         =   0   'False
            Width           =   2010
         End
         Begin TDBNumber6Ctl.TDBNumber txtnilai 
            Height          =   255
            Left            =   9630
            TabIndex        =   38
            Top             =   4755
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            Calculator      =   "frmAddSOP.frx":1592
            Caption         =   "frmAddSOP.frx":15B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAddSOP.frx":161E
            Keys            =   "frmAddSOP.frx":163C
            Spin            =   "frmAddSOP.frx":167E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.000;(###,###,###,##0.000);0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.000;(###,###,###,##0.000)"
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
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
            Height          =   4620
            Left            =   30
            TabIndex        =   39
            Top             =   495
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   8149
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label tg1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   255
            Left            =   9360
            TabIndex        =   61
            Top             =   195
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PEMAKAIAN BAHAN BAKU :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   210
            Width           =   3015
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   5220
         Left            =   -69970
         TabIndex        =   41
         Top             =   330
         Visible         =   0   'False
         Width           =   11310
         _Version        =   851970
         _ExtentX        =   19950
         _ExtentY        =   9208
         _StockProps     =   1
         Page            =   3
         Begin VB.TextBox txtsolid 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   45
            Top             =   930
            Width           =   2130
         End
         Begin VB.TextBox txtviskositas 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   44
            Top             =   570
            Width           =   2130
         End
         Begin VB.TextBox txttesvisual 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   43
            Top             =   210
            Width           =   2130
         End
         Begin VB.ComboBox cmbqc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1485
            TabIndex        =   42
            Top             =   1305
            Width           =   2145
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Solid (%) :"
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
            Left            =   45
            TabIndex        =   49
            Top             =   975
            Width           =   1380
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Viskositas mPA.s :"
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
            Left            =   45
            TabIndex        =   48
            Top             =   615
            Width           =   1380
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tes Visual :"
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
            Left            =   45
            TabIndex        =   47
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lulus / Tidak :"
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
            Left            =   45
            TabIndex        =   46
            Top             =   1365
            Width           =   1380
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid4 
         Height          =   4800
         Left            =   -69925
         TabIndex        =   50
         Top             =   705
         Visible         =   0   'False
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   8467
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   16777215
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label14 
         Caption         =   "PEROLEHAN BARANG JADI :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -69880
         TabIndex        =   52
         Top             =   450
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin XtremeSuiteControls.PushButton cmdstok 
      Height          =   600
      Left            =   4920
      TabIndex        =   63
      Top             =   7545
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   1058
      _StockProps     =   79
      Caption         =   "CEK STOK"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   2160
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.PushButton cmdtambahrec 
      Height          =   585
      Left            =   105
      TabIndex        =   64
      Top             =   7530
      Width           =   2010
      _Version        =   851970
      _ExtentX        =   3545
      _ExtentY        =   1032
      _StockProps     =   79
      Caption         =   "UBAH DATA"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAddSOP.frx":16A6
   End
   Begin XtremeSuiteControls.DateTimePicker Datetambah 
      Height          =   315
      Left            =   9960
      TabIndex        =   71
      Top             =   1560
      Width           =   1515
      _Version        =   851970
      _ExtentX        =   2672
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "TANGGAL PENAMBAHAN BAHAN BAKU DAN KEMASAN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   72
      Top             =   1605
      Width           =   4290
   End
End
Attribute VB_Name = "frmaddsop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset
Private SQL1 As String
Private OBJ2 As New ADODB.Connection
Private RST2 As ADODB.Recordset
Private SQL2 As String
Private edit_mode As Boolean
Private edit_mode2 As Boolean

Private poscol As Integer
Private posrow As Integer

Private noref1 As String
Private noref2 As String
Private noproses As Integer
Private pesan As Integer
Private akses As Boolean

Private Sub btnClose_Click()
    'Cek HPP Base yang digunakan, kalau cancel hapus di am_stoklot
    cekbase
End Sub

Private Sub btnDelete_Click()
    If MsgBox("Anda yakin ingin menghapus data ini ?", vbYesNo + vbQuestion, AppName) = vbYes Then
        HapusSOP
        Exit Sub
    End If
End Sub

Private Sub HapusSOP()
    On Error GoTo Err_handler:
    
    If MsgBox("Apakah anda yakin akan menghapus SOP Tersebut...? ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Anda tidak memiliki akses...! ", vbCritical, AppName
        Exit Sub
    End If
    
proses:
    OBJ.Open dsn
    If txtnolot(0) = "" Then Exit Sub
    
    'PROSES HAPUS DATA LIST_PRODUKSI_CHILD ********************************************************************************
    SQL = "delete from list_produksi_child where nolot= '" & txtnolot(0) & "' and kode_produk='" & txtkodeproduk & "'"
    OBJ.Execute SQL
    
    'PROSES HAPUS DATA LIST_PRODUKSI_MASTER *******************************************************************************
    SQL = "Delete from list_produksi_master Where nolot= '" & txtnolot(0) & "' and kode_produk='" & txtkodeproduk & "'"
    OBJ.Execute SQL
    
    'PROSES HAPUS DATA LIST_PRODUKSI_HASIL ********************************************************************************
    SQL = "Delete from list_produksi_hasil Where nolot= '" & txtnolot(0) & "' and kode_produk='" & txtkodeproduk & "'"
    OBJ.Execute SQL
    
    
    'PROSES HAPUS DATA IST_PRODUKSI_KEMASAN *******************************************************************************
    SQL = "Delete from list_produksi_kemasan Where nolot= '" & txtnolot(0) & "' and kode_produk='" & txtkodeproduk & "'"
    OBJ.Execute SQL
    
    'PROSES HAPUS DATA AM_USERHDR *****************************************************************************************
    SQL = "Delete From am_usehdr Where nobpb like  '" & txtnolot(0) & "%'"
    OBJ.Execute SQL
    
    'PROSES HAPUS DATA AM_USELIN ******************************************************************************************
    SQL = "Delete From am_uselin Where nobpb like '" & txtnolot(0) & "%'"
    OBJ.Execute SQL
    
    'PROSES HAPUS DATA AM_BPBHDR ******************************************************************************************
    'SQL = "Delete From am_bpbhdr Where keterangan like '" &  & "%'"
    'OBJ.Execute SQL
        
    'PROSES HAPUS DATA AM_BPBLIN ******************************************************************************************
    'SQL = "Delete From am_bpblin Where nobpb = '" & txtnobpb & "%'"
    'OBJ.Execute SQL
    
    OBJ.Close
    MsgBox "Data berhasil dihapus", vbInformation, AppName
    btnNew_Click
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Proses hapus tidak berhasil..! " + Err.Description, vbCritical, AppName
End Sub

Private Sub btnNew_Click()
    txtkodeproduk = ""
    txtproduk = ""
    txtnolot(0) = ""
    txtnobpb = ""
    txtnoreaktor = ""
    datebahan = Date
    hapusgrid1
    hapusgrid2
    hapusgrid3
    hapusgrid4
    tg1 = "0.00"
    txttotalproduksi = "0.00"
    txttotalhasilproduksi = "0.00"
    txtwaktupelarutan = ""
    txtwaktukemasan = ""
    txtwaktutambahan = ""
    txttesvisual = "0"
    txtviskositas = "0"
    txtsolid = "0"
    cmbqc.text = "Lulus"
    edit_mode = False
    edit_mode2 = False
    noref1 = ""
    noref2 = ""
    TabControl1.Item(0).Selected = True
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    chkWIP.Value = 0
    btnsave.Enabled = True
    noproses = 0
    Timer1.Enabled = False
End Sub

Private Sub btnsave_Click()
    'On Error GoTo Err_handler:
    'deklarasi variable proses simpan data sop
    Dim total_hpp As Double
    Dim hpp As Double
    Dim hpp_bahan_utama As Double
    Dim hpp_bahan_tambahan As Double
    Dim hpp_bahan_kemasan As Double
    
    Dim stokbahan As Double
    Dim stokbahan_tambahan As Double
    Dim stokbahan_kemasan As Double
    
    Dim totalproduksi As Double
    Dim totalhasilproduksi As Double
    
    
    If txtkodeproduk = "" Then
        MsgBox "Data tidak lengkap", vbInformation, AppName
        Exit Sub
    End If
    
    If txtnolot(0) = "" Then
        MsgBox "Nomor Lot produksi harus diisi...!", vbCritical, AppName
        Exit Sub
    End If
    
'AWAL PROSES EDIT SOP *************************************************************************************
    If edit_mode = True Then
        If edit_mode2 = False Then
            'cari urutan proses
            Dim proses As Integer
            
            OBJ.Open dsn
            SQL = "select max(proses_ke) as proses from list_historisop where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            proses = RST!proses + 1
            
            'proses simpan data bahan tambahan
            SQL = "select * from list_produksi_child where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid2.Row = 1
            Do While True
                If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!kode_produk = txtkodeproduk
                RST!nolot = txtnolot(0)
                RST!kode_bahan = grid2.TextMatrix(grid2.Row, 1)
                RST!Lot_bahan = grid2.TextMatrix(grid2.Row, 3)
                
                RST!qty_bahan = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
                RST!KODE_SATUAN = grid2.TextMatrix(grid2.Row, 5)
                RST!flag_tambahan = "1"
                RST!hpp = Format(grid2.TextMatrix(grid2.Row, 7), "general number")
                RST!tanggal = Format(Datetambah, "yyyy/MM/dd")
                RST!REF = txtnolot(0) & "/" & Trim(Str(proses))
                RST!Line = "0"
                RST!proses_ke = proses
                RST.Update
                grid2.Row = grid2.Row + 1
            Loop
            
            'proses simpan pemakaian kemasan
            SQL = "select * from list_produksi_kemasan where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid3.Row = 1
            Do While True
                If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!kode_produk = txtkodeproduk
                RST!nolot = txtnolot(0)
                RST!kode_bahan = grid3.TextMatrix(grid3.Row, 1)
                RST!Lot_bahan = grid3.TextMatrix(grid3.Row, 3)
                RST!qty_bahan = Format(grid3.TextMatrix(grid3.Row, 4), "general number")
                RST!KODE_SATUAN = grid3.TextMatrix(grid3.Row, 5)
                RST!flag_tambahan = "0"
                RST!hpp = Format(grid3.TextMatrix(grid3.Row, 7), "general number")
                RST!tanggal = Format(Datetambah, "yyyy/MM/dd")
                RST!noref = txtnolot(0) & "/" & Trim(Str(proses))
                
                RST!proses_ke = proses
                RST.Update
                grid3.Row = grid3.Row + 1
            Loop

            'Proses Simpan Ke tabel am_usehdr
            SQL = "Select * From am_usehdr Where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                .AddNew
                !nobpb = txtnolot(0) & "/" & Trim(Str(proses))
                !tglbpb = Format(Datetambah, "yyyy/MM/dd")
                !noorder = txtnolot(0) & "/" & Trim(Str(proses))
                .Update
            End With
            
            'proses simpan ke uselin data bahan baku tambahan
            SQL = "Select * From am_uselin Where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid2.Row = 1
            Do While True
                If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!nobpb = txtnolot(0) & "/" & Trim(Str(proses))
                RST!kodebarang = grid2.TextMatrix(grid2.Row, 1)
                RST!qty = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
                RST!kodesatuan = grid2.TextMatrix(grid2.Row, 5)
                RST!lineitem = grid2.Row
                RST.Update
                grid2.Row = grid2.Row + 1
            Loop
            
            'proses simpan ke uselin data kemasan
            SQL = "Select * From am_uselin Where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid3.Row = 1
            Do While True
                If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!nobpb = txtnolot(0) & "/" & Trim(Str(proses))
                RST!kodebarang = grid3.TextMatrix(grid3.Row, 1)
                RST!qty = Format(grid3.TextMatrix(grid3.Row, 4), "general number")
                RST!kodesatuan = grid3.TextMatrix(grid3.Row, 5)
                RST!lineitem = grid3.Row + grid2.Rows - 1
                RST.Update
                grid3.Row = grid3.Row + 1
            Loop
            
            'proses simpan ke produksi hasil
            'cari nobpb
            txtnobpb = getnobpb(Format(Date, "yymm"))
        
            SQL = "select * from list_produksi_hasil where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid4.Row = 1
            Do While True
                If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!kode_produk = txtkodeproduk
                RST!nolot = txtnolot(0)
                RST!kode_bahan = grid4.TextMatrix(grid4.Row, 1)
                RST!Lot_bahan = grid4.TextMatrix(grid4.Row, 3)
                RST!qty_bahan = Format(grid4.TextMatrix(grid4.Row, 4), "general number")
                RST!KODE_SATUAN = grid4.TextMatrix(grid4.Row, 5)
                RST!flag_tambahan = "0"
                RST!tanggal = Format(datedone, "yyyy/MM/dd")
                RST!noref = txtnobpb
                RST!proses_ke = proses
                RST.Update
                grid4.Row = grid4.Row + 1
            Loop
            
            'CARI TOTAL PRODUKSI TERAKHIR
            SQL = "select sum(qty_bahan) as totalprod from list_produksi_child where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                totalproduksi = 0
            Else
                totalproduksi = RST!totalprod
            End If
            
            'CARI TOTAL HASIL PRODUKSI
            SQL = "select * from list_produksi_hasil where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                totalhasilproduksi = 0
            Else
                Do While Not RST.EOF
                    totalhasilproduksi = totalhasilproduksi + GetToKilogram(RST!kode_bahan, RST!KODE_SATUAN, Format(datedone, "yyyy/MM/dd"))
                    RST.MoveNext
                Loop
            End If
            
            
            'CARI HPP BAHAN BAKU
            SQL = "select sum(hpp) as jmlhpp from list_produksi_child where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                hpp_bahan_utama = 0
            Else
                hpp_bahan_utama = RST!jmlhpp
            End If
            
            'CARI HPP KEMASAN
            SQL = "select sum(hpp) as jmlhpp from list_produksi_kemasan where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If IsNull(RST!jmlhpp) Then
                hpp_bahan_kemasan = 0
            Else
                hpp_bahan_kemasan = RST!jmlhpp
            End If
            
            'TOTAL HPP
            hpp = hpp_bahan_utama + hpp_bahan_kemasan
            
            'proses update header
            'RUBAH DATA TABLE LIST PRODUKSI MASTER
            SQL = "select * from list_produksi_master where nolot ='" & txtnolot(0) & "'"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                !total_produksi = totalproduksi
                !total_hasil_produksi = totalhasilproduksi
                !waktu_larut = txtwaktupelarutan
                !waktu_tambahan = txtwaktutambahan
                !waktu_kemasan = txtwaktukemasan
                If txttesvisual = "" Then
                    !qc_test_visual = "0"
                Else
                    !qc_test_visual = txttesvisual
                End If
                If txtviskositas = "" Then
                    !qc_viskositas = "0"
                Else
                    !qc_viskositas = txtviskositas
                End If
                If txtsolid = "" Then
                    !qc_solid = "0"
                Else
                    !qc_solid = txtsolid
                End If
                If cmbqc.text = "Lulus" Then
                    !flag_status = "0"
                Else
                    !flag_status = "1"
                End If
                !ref1 = ""
                !ref2 = ""

                !tglakhir = datedone
                If Check2.Value = 1 And Check3.Value = 0 And Check1.Value = 0 Then
                    !flagprint = "2"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 0 Then
                    !flagprint = "3"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 1 Then
                    !flagprint = "4"
                End If
                !userupdate = nmuser
                !hpp = hpp
                .Update
            End With
            
            'PROSES SIMPAN KE HISTORI SOP
            SQL = "select * from list_historisop where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                .AddNew
                !nolot = txtnolot(0)
                !tanggal = Format(datebahan, "yyyy/MM/dd")
                !proses_ke = Trim(Str(proses))
                !UserName = nmuser
                !ref1 = txtnolot(0) & "/" & Trim(Str(proses))
                !ref2 = txtnobpb
                If Check2.Value = 1 And Check3.Value = 0 And Check1.Value = 0 Then
                    !flagprint = "2"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 0 Then
                    !flagprint = "3"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 1 Then
                    !flagprint = "4"
                End If
                .Update
            End With
            txttotalproduksi.Value = GetTotalProduksi(txtnolot(0))
            OBJ.Close
            
            Call hppproduksi
            MsgBox "Data is saved...!", vbInformation, AppName
            btnNew_Click
            Exit Sub
        End If

        If MsgBox("Apakah anda yakin akan merubah data tersebut", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
                
        If Check2.Value = 1 Then
            If txttotalhasilproduksi.Value = 0 Then
                MsgBox "Total Hasil Produksi harus diisi...!", vbCritical, AppName
                Exit Sub
            End If
        End If
        
        If Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 0 Then
            pesan = MsgBox("Apakah proses SOP telah selesai ?", vbQuestion + vbYesNo, AppName)
            If pesan = vbYes Then
                Check1.Value = 1
            End If
        End If
        
        OBJ.Open dsn
        
        'hapus data bahan baku tambahan
        SQL = "delete from list_produksi_child where ref='" & noref1 & "'"
        OBJ.Execute SQL
        
        'simpan data bahan baku tambahan
        SQL = "select * from list_produksi_child where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid2.Row = 1
        Do While True
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!kode_produk = txtkodeproduk
            RST!nolot = txtnolot(0)
            RST!kode_bahan = grid2.TextMatrix(grid2.Row, 1)
            RST!Lot_bahan = grid2.TextMatrix(grid2.Row, 3)
            RST!qty_bahan = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
            RST!KODE_SATUAN = grid2.TextMatrix(grid2.Row, 5)
            RST!flag_tambahan = "1"
            RST!hpp = Format(grid2.TextMatrix(grid2.Row, 7), "general number")
            RST!tanggal = Format(datedone, "yyyy/MM/dd")
            RST!REF = txtnolot(0) & "/" & Trim(Str(noproses))
            RST!proses_ke = noproses
            RST.Update
            grid2.Row = grid2.Row + 1
        Loop
        
        'hapus data kemasan
        SQL = "delete from list_produksi_kemasan where noref='" & noref1 & "'"
        OBJ.Execute SQL
        
        'simpan data kemasan
        SQL = "select * from list_produksi_kemasan where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid3.Row = 1
        Do While True
            If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!kode_produk = txtkodeproduk
            RST!nolot = txtnolot(0)
            RST!kode_bahan = grid3.TextMatrix(grid3.Row, 1)
            RST!Lot_bahan = grid3.TextMatrix(grid3.Row, 3)
            RST!qty_bahan = Format(grid3.TextMatrix(grid3.Row, 4), "general number")
            RST!KODE_SATUAN = grid3.TextMatrix(grid3.Row, 5)
            RST!flag_tambahan = "0"
            RST!hpp = Format(grid3.TextMatrix(grid3.Row, 7), "general number")
            RST!tanggal = Format(datebahan, "yyyy/MM/dd")
            RST!noref = noref1
            RST!proses_ke = proses
            RST.Update
            grid3.Row = grid3.Row + 1
        Loop
        
        'delete di data usehdr
        SQL = "delete from am_usehdr where nobpb ='" & noref1 & "'"
        OBJ.Execute SQL
        
        'Proses Simpan Ke tabel am_usehdr
        SQL = "Select * From am_usehdr Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        With RST
            .AddNew
            !nobpb = noref1
            !tglbpb = Format(datedone, "yyyy/MM/dd")
            !noorder = noref1
            .Update
        End With
        
        'proses hapus data uselin
        SQL = "delete from am_uselin where nobpb='" & noref1 & "'"
        OBJ.Execute SQL
        
        'proses simpan ke uselin datab bahan baku tambahan
        SQL = "Select * From am_uselin Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid2.Row = 1
        Do While True
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!nobpb = noref1
            RST!kodebarang = grid2.TextMatrix(grid2.Row, 1)
            RST!qty = Format(grid2.TextMatrix(grid2.Row, 4), "general number")
            RST!kodesatuan = grid2.TextMatrix(grid2.Row, 5)
            RST!lineitem = grid2.Row
            RST.Update
            grid2.Row = grid2.Row + 1
        Loop
        
        'proses simpan ke uselin data kemasan
        SQL = "Select * From am_uselin Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid3.Row = 1
        Do While True
            If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!nobpb = noref1
            RST!kodebarang = grid3.TextMatrix(grid3.Row, 1)
            RST!qty = Format(grid3.TextMatrix(grid3.Row, 4), "general number")
            RST!kodesatuan = grid3.TextMatrix(grid3.Row, 5)
            RST!lineitem = grid3.Row = grid2.Rows - 1
            RST.Update
            grid3.Row = grid3.Row + 1
        Loop
        
        'hapus data produksi hasil
        SQL = "delete from list_produksi_hasil where noref ='" & noref2 & "'"
        
        'simpan data produksi hasil
        SQL = "select * from list_produksi_hasil where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid4.Row = 1
        Do While True
            If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
            RST.AddNew
            RST!kode_produk = txtkodeproduk
            RST!nolot = txtnolot(0)
            RST!kode_bahan = grid4.TextMatrix(grid4.Row, 1)
            RST!Lot_bahan = grid4.TextMatrix(grid4.Row, 3)
            RST!qty_bahan = Format(grid4.TextMatrix(grid4.Row, 4), "general number")
            RST!KODE_SATUAN = grid4.TextMatrix(grid4.Row, 5)
            RST!flag_tambahan = "0"
            RST!tanggal = Format(datedone, "yyyy/MM/dd")
            RST!noref = noref2
            RST!proses_ke = proses
            RST.Update
            grid4.Row = grid4.Row + 1
        Loop
        
        'proses update hasil produksi
        'CARI TOTAL PRODUKSI TERAKHIR
            SQL = "select sum(qty_bahan) as totalprod from list_produksi_child where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                totalproduksi = 0
            Else
                totalproduksi = RST!totalprod
            End If
            
            'CARI TOTAL HASIL PRODUKSI
            SQL = "select * from list_produksi_hasil where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                totalhasilproduksi = 0
            Else
                Do While Not RST.EOF
                    totalhasilproduksi = totalhasilproduksi + GetToKilogram(RST!kode_bahan, RST!KODE_SATUAN, Format(datedone, "yyyy/MM/dd"))
                    RST.MoveNext
                Loop
            End If
            
            
            'CARI HPP BAHAN BAKU
            SQL = "select sum(hpp) as jmlhpp from list_produksi_child where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                hpp_bahan_utama = 0
            Else
                hpp_bahan_utama = RST!jmlhpp
            End If
            
            'CARI HPP KEMASAN
            SQL = "select sum(hpp) as jmlhpp from list_produksi_kemasan where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
            
            If RST.EOF Then
                hpp_bahan_kemasan = 0
            Else
                hpp_bahan_kemasan = RST!jmlhpp
            End If
            
            'TOTAL HPP
            hpp = hpp_bahan_utama + hpp_bahan_kemasan
            
            'proses update header
            'RUBAH DATA TABLE LIST PRODUKSI MASTER
            SQL = "select * from list_produksi_master where nolot ='" & txtnolot(0) & "'"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                !total_produksi = totalproduksi
                !total_hasil_produksi = totalhasilproduksi
                !waktu_larut = txtwaktupelarutan
                !waktu_tambahan = txtwaktutambahan
                !waktu_kemasan = txtwaktukemasan
                If txttesvisual = "" Then
                    !qc_test_visual = "0"
                Else
                    !qc_test_visual = txttesvisual
                End If
                If txtviskositas = "" Then
                    !qc_viskositas = "0"
                Else
                    !qc_viskositas = txtviskositas
                End If
                If txtsolid = "" Then
                    !qc_solid = "0"
                Else
                    !qc_solid = txtsolid
                End If
                If cmbqc.text = "Lulus" Then
                    !flag_status = "0"
                Else
                    !flag_status = "1"
                End If
                !ref1 = ""
                !ref2 = ""

                !tglakhir = datedone
                If Check2.Value = 1 And Check3.Value = 0 And Check1.Value = 0 Then
                    !flagprint = "2"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 0 Then
                    !flagprint = "3"
                ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 1 Then
                    !flagprint = "4"
                End If
                !userupdate = nmuser
                !hpp = hpp
                .Update
            End With
            
            'proses simpan bahan baku utama ke list produksi child
            SQL = "select * from list_produksi_child where nolot='" & txtnolot(0) & "' and proses_ke='1'"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid1.Row = 1
            Do While True
                If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
                RST!kode_produk = txtkodeproduk
                RST!nolot = txtnolot(0)
                RST!kode_bahan = grid1.TextMatrix(grid1.Row, 1)
                RST!Lot_bahan = grid1.TextMatrix(grid1.Row, 3)
                RST!qty_bahan = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
                RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
                RST!flag_tambahan = "0"
                RST!hpp = Format(grid1.TextMatrix(grid1.Row, 7), "##,###,###,##0.00")
                RST!tanggal = Format(datebahan, "yyyyMMdd")
                RST!REF = txtnolot(0) & "/1"
                RST!Line = Format(grid1.TextMatrix(grid1.Row, 0), "general number")
                RST!proses_ke = "1"
                RST.Update
                grid1.Row = grid1.Row + 1
            Loop

                
        'update flag histori sop
        SQL = "select * from list_historisop where ref1='" & noref1 & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        With RST
            If Check2.Value = 1 And Check3.Value = 0 And Check1.Value = 0 Then
                !flagprint = "2"
            ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 0 Then
                !flagprint = "3"
            ElseIf Check2.Value = 1 And Check3.Value = 1 And Check1.Value = 1 Then
                !flagprint = "4"
            End If
            .Update
        End With
        
        'update hpp bahan baku perkg di tabel am_stok
        SQL = "Select nolot,SUM(hpp)/SUM(qty_bahan)'perkg' From list_produksi_child Where nolot = '" & txtnolot(0) & "' group by nolot"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then

            OBJ1.Open dsn
            SQL1 = "Update am_stok set hpp='" & RST!perkg & "' Where nolot='" & txtnolot(0) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        
        If chkWIP.Value = 0 Then
            SQL = "Delete am_sopbase where nolot='" & txtnolot(0) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        
        OBJ.Close
        MsgBox "Data is saved...!", vbInformation, AppName
        btnNew_Click
        Exit Sub
    End If
'AKHIR PROSES EDIT DATA SOP
    
    
'PROSES SIMPAN DATA (ADD NEW) ################################################################################
    
If UserOnLineLevel = "creator" Then GoTo proses:
If UserOnLineLevel = "Administrator" Then GoTo proses:
If UserOnLineLevel <> "Supervisor" Then
    MsgBox "Proses ditolak...1", vbCritical, AppName
Exit Sub
End If
    
proses:
    OBJ.Open dsn
    'validasi nomor lot
    SQL = "select * from list_produksi_master where nolot ='" & txtnolot(0) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Nolot telah terpakai, Silahkan lakukan perubahan.....!", vbInformation, AppName
        Exit Sub
    End If
    
    'CARI HPP
    hpp_bahan_utama = 0
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        hpp_bahan_utama = hpp_bahan_utama + Format(grid1.TextMatrix(grid1.Row, 7), "general number")
        grid1.Row = grid1.Row + 1
    Loop
        
    'proses simpan ke table list produksi master
    SQL = "select * from list_produksi_master where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kode_produk = txtkodeproduk
        !nolot = txtnolot(0)
        !tanggal = Format(datebahan, "yyyy/MM/dd")
        !total_produksi = tg1
        !total_hasil_produksi = txttotalhasilproduksi
        !waktu_larut = txtwaktupelarutan
        !waktu_tambahan = txtwaktutambahan
        !waktu_kemasan = txtwaktukemasan
        If txttesvisual = "" Then
            !qc_test_visual = "0"
        Else
            !qc_test_visual = txttesvisual
        End If
        If txtviskositas = "" Then
            !qc_viskositas = "0"
        Else
            !qc_viskositas = txtviskositas
        End If
        If txtsolid = "" Then
            !qc_solid = "0"
        Else
            !qc_solid = txtsolid
        End If
        If cmbqc.text = "Lulus" Then
            !flag_status = "0"
        Else
            !flag_status = "1"
        End If
        !Usercreate = nmuser
        !ref1 = ""
        !ref2 = ""
        !tglakhir = datedone
        !flagprint = "1"
        !noreaktor = txtnoreaktor
        !userupdate = nmuser
        !hpp = hpp_bahan_utama
        .Update
    End With
    
    'proses simpan bahan baku utama ke list produksi child
    SQL = "select * from list_produksi_child where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = txtkodeproduk
        RST!nolot = txtnolot(0)
        RST!kode_bahan = grid1.TextMatrix(grid1.Row, 1)
        RST!Lot_bahan = grid1.TextMatrix(grid1.Row, 3)
        RST!qty_bahan = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
        RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
        RST!flag_tambahan = "0"
        RST!hpp = Format(grid1.TextMatrix(grid1.Row, 7), "##,###,###,##0.00")
        RST!tanggal = Format(datebahan, "yyyyMMdd")
        RST!REF = txtnolot(0) & "/1"
        RST!Line = Format(grid1.TextMatrix(grid1.Row, 0), "general number")
        RST!proses_ke = "1"
        RST.Update
        grid1.Row = grid1.Row + 1
    Loop
    
    'Proses Simpan Ke tabel am_usehdr
    SQL = "Select * From am_usehdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nobpb = txtnolot(0) & "/1"
        !tglbpb = datebahan
        !noorder = txtnolot(0) & "/1"
        .Update
    End With
    
    'PROSES SIMPAN KE 1 KE USE LIN
    SQL = "Select * From am_uselin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!nobpb = txtnolot(0) & "/1"
        RST!kodebarang = grid1.TextMatrix(grid1.Row, 1)
        RST!qty = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
        RST!kodesatuan = grid1.TextMatrix(grid1.Row, 5)
        RST!lineitem = grid1.Row
        RST.Update
        grid1.Row = grid1.Row + 1
    Loop
    
    'proses simpan perolehan
    SQL = "select * from list_produksi_hasil where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid4.Row = 1
    Do While True
        If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = txtkodeproduk
        RST!nolot = txtnolot(0)
        RST!kode_bahan = grid4.TextMatrix(grid4.Row, 1)
        RST!Lot_bahan = grid4.TextMatrix(grid4.Row, 3)
        RST!qty_bahan = Format(grid4.TextMatrix(grid4.Row, 4), "general number")
        RST!KODE_SATUAN = grid4.TextMatrix(grid4.Row, 5)
        RST!flag_tambahan = "0"
        RST!noref = ""
        RST!tanggal = Format(datebahan, "yyyyMMdd")
        RST!proses_ke = "1"
        RST.Update
        grid4.Row = grid4.Row + 1
    Loop
        
   'PROSES SIMPAN KE HISTORI SOP
   SQL = "select * from list_historisop where 0=1"
   Set RST = New ADODB.Recordset
   RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot(0)
        !tanggal = Format(datebahan, "yyyyMMdd")
        !proses_ke = "1"
        !UserName = nmuser
        !ref1 = ""
        !ref2 = ""
        !flagprint = "1"
        .Update
    End With
    
    'Simpan Otoritas Edit Data SOP
    SQL = "Select * From list_masterkeyLot where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !noso = txtnolot(0)
        !tgl = Format(datebahan, "dd-MM-yyyy")
        !UserName = nmuser
        !otoritas = "0"
        !cetaksop = "0"
        !otorisasi = ""
        !keterangan = ""
        .Update
    End With
    
    'PROSES SIMPAN KE AM_SOPBASE untuk menandakan sop base wip (untuk bahan baku)
    If chkWIP.Value = 1 Then
        SQL = "Insert into am_sopbase(nolot,tgl) values('" & txtnolot(0) & "',convert(datetime,'" & tanggalsekarang & "'))"
        Set RST = OBJ.Execute(SQL)
    End If
    
    OBJ.Close
    MsgBox "Data is saved...!", vbInformation, AppName
    cetaksop
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub hppproduksi()
    OBJ.Open dsn
    SQL = "Select a.noref,a.tanggal,b.kodebarang,b.kg,isnull(c.pack,0)'pack',e.thppbahan,e.perkilo,isnull(g.thpppack,0)'thpppack',g.thasil,(a.qty_bahan*b.kg)'hasil',"
    SQL = SQL + " e.thppbahan +isnull(g.thpppack,0)'brutto',(g.thasil*e.perkilo)+isnull(g.thpppack,0)'tjadi',"
    SQL = SQL + " (e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0))'loss',"
    SQL = SQL + " (((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil)'lossperkg',"
    SQL = SQL + " (isnull(c.pack,0)/(a.qty_bahan*b.kg))'packperkg',"
    SQL = SQL + " (e.perkilo + (isnull(c.pack,0)/(a.qty_bahan*b.kg))+(((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil))'hppperkg'"
    SQL = SQL + " From list_produksi_hasil a"
    SQL = SQL + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtnolot(0) & "',3,1)='A' then kg1"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "',3,1)='B' then kg2"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='C' then kg3"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='D' then kg4"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='E' then kg5"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='F' then kg6"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='G' then kg7"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='H' then kg8"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='J' then kg9"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='K' then kg10"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='L' then kg11"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
    SQL = SQL + " b on a.kode_bahan = b.kodebarang and a.kode_satuan = b.kodesatuan"
    SQL = SQL + " left join (select noref,isnull(SUM(hpp),0)'pack' from list_produksi_kemasan where nolot = '" & txtnolot(0) & "' group by noref) c on a.noref = c.noref"
    SQL = SQL + " left join list_produksi_child d on a.nolot = d.nolot"
    SQL = SQL + " inner join (Select x.nolot,y.noref,SUM(x.hpp)'thppbahan',SUM(x.hpp)/SUM(x.qty_bahan)'perkilo'"
    SQL = SQL + " from list_produksi_child x left join list_produksi_hasil y on x.nolot = y.nolot where x.nolot ='" & txtnolot(0) & "'"
    SQL = SQL + " group by x.nolot,y.noref) e on a.noref = e.noref"
    SQL = SQL + " left join (Select m.nolot,o.thpppack,SUM(m.qty_bahan * n.kg)'thasil' From list_produksi_hasil m"
    SQL = SQL + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtnolot(0) & "',3,1)='A' then kg1"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "',3,1)='B' then kg2"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='C' then kg3"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='D' then kg4"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='E' then kg5"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='F' then kg6"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='G' then kg7"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='H' then kg8"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='J' then kg9"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='K' then kg10"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='L' then kg11"
    SQL = SQL + " when SUBSTRING('" & txtnolot(0) & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
    SQL = SQL + " n on m.kode_bahan = n.kodebarang and m.kode_satuan = n.kodesatuan"
    SQL = SQL + " left join (Select nolot,isnull(SUM(hpp),0)'thpppack' from list_produksi_kemasan Where nolot = '" & txtnolot(0) & "' group by nolot)"
    SQL = SQL + " o on m.nolot=o.nolot"
    SQL = SQL + " Where m.nolot = '" & txtnolot(0) & "' and m.proses_ke = '2' and n.tahun = '20' + LEFT('" & txtnolot(0) & "',2) group by m.nolot,o.thpppack) g on a.nolot = g.nolot"
    SQL = SQL + " Where a.nolot = '" & txtnolot(0) & "' and b.tahun = '20' + LEFT('" & txtnolot(0) & "',2)"
    SQL = SQL + " group by a.noref,a.tanggal,b.kodebarang,b.kg,c.pack,e.thppbahan,e.perkilo,g.thpppack,g.thasil,a.qty_bahan order by a.noref asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "Update am_stok set hpp='" & RST!hppperkg & "',hpp_totpack='" & RST!pack & "' Where nolot='" & txtnolot(0) & "'"
        SQL1 = SQL1 + " and flag='0' and palet='" & RST!noref & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        'Cek list_hpp_produksi
        SQL1 = "Select * From list_hpp_produksi Where nolot='" & txtnolot(0) & "' and palet='" & RST!noref & "'"
        Set RST1 = New ADODB.Recordset
        RST1.Open SQL1, OBJ1, adOpenDynamic, adLockOptimistic
        If RST1.EOF Then
            RST1.AddNew
            RST1!nolot = txtnolot(0).text
            RST1!palet = RST!noref
            RST1!tanggal = Format(datedone, "yyyyMMdd")
            RST1!kodebarang = RST!kodebarang
            RST1!kg = RST!kg
            RST1!hpppackperpalet = RST!pack
            RST1!bahanperkg = RST!perkilo
            RST1!tkglot = RST!thasil
            RST1!kgperpalet = RST!hasil
            RST1!thpplot = RST!brutto
            RST1!thppbahan = RST!thppbahan
            RST1!thpppack = RST!thpppack
            RST1!thppjadi = RST!tjadi
            RST1!thpploss = RST!loss
            RST1!lossperkg = RST!lossperkg
            RST1!packperkg = RST!packperkg
            RST1!hppperkg = RST!hppperkg
            RST1!flag = "0"
            RST1.Update
        Else
            RST1!kodebarang = RST!kodebarang
            RST1!kg = RST!kg
            RST1!hpppackperpalet = RST!pack
            RST1!bahanperkg = RST!perkilo
            RST1!tkglot = RST!thasil
            RST1!kgperpalet = RST!hasil
            RST1!thpplot = RST!brutto
            RST1!thppbahan = RST!thppbahan
            RST1!thpppack = RST!thpppack
            RST1!thppjadi = RST!tjadi
            RST1!thpploss = RST!loss
            RST1!lossperkg = RST!lossperkg
            RST1!packperkg = RST!packperkg
            RST1!hppperkg = RST!hppperkg
            RST1!flag = "0"
            RST1.Update
        End If
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
End Sub
Function tanggalsekarang()
    tanggalsekarang = Year(Date) & "-" & Month(Date) & "-" & Day(Date)
End Function
Private Sub save()
On Error GoTo Err_handler:
Dim hpp_bahan_utama As Long
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Proses ditolak...1", vbCritical, AppName
        Exit Sub
    End If
    
proses:
    OBJ.Open dsn
    'validasi nomor lot
    SQL = "select * from list_produksi_master where nolot ='" & txtnolot(0) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Nolot telah terpakai, Silahkan lakukan perubahan.....!", vbInformation, AppName
        Exit Sub
    End If
    
    'CARI HPP
    hpp_bahan_utama = 0
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        hpp_bahan_utama = hpp_bahan_utama + Format(grid1.TextMatrix(grid1.Row, 7), "general number")
        grid1.Row = grid1.Row + 1
    Loop
        
    'proses simpan ke table list produksi master
    SQL = "select * from list_produksi_master where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kode_produk = txtkodeproduk
        !nolot = txtnolot(0)
        !tanggal = Format(datebahan, "yyyy/MM/dd")
        !total_produksi = tg1
        !total_hasil_produksi = txttotalhasilproduksi
        !waktu_larut = txtwaktupelarutan
        !waktu_tambahan = txtwaktutambahan
        !waktu_kemasan = txtwaktukemasan
        If txttesvisual = "" Then
            !qc_test_visual = "0"
        Else
            !qc_test_visual = txttesvisual
        End If
        If txtviskositas = "" Then
            !qc_viskositas = "0"
        Else
            !qc_viskositas = txtviskositas
        End If
        If txtsolid = "" Then
            !qc_solid = "0"
        Else
            !qc_solid = txtsolid
        End If
        If cmbqc.text = "Lulus" Then
            !flag_status = "0"
        Else
            !flag_status = "1"
        End If
        !Usercreate = nmuser
        !ref1 = ""
        !ref2 = ""
        !tglakhir = datedone
        !flagprint = "1"
        !noreaktor = txtnoreaktor
        !userupdate = nmuser
        !hpp = hpp_bahan_utama
        .Update
    End With
    
    'proses simpan bahan baku utama ke list produksi child
    SQL = "select * from list_produksi_child where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = txtkodeproduk
        RST!nolot = txtnolot(0)
        RST!kode_bahan = grid1.TextMatrix(grid1.Row, 1)
        RST!Lot_bahan = grid1.TextMatrix(grid1.Row, 3)
        RST!qty_bahan = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
        RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
        RST!flag_tambahan = "0"
        RST!hpp = Format(grid1.TextMatrix(grid1.Row, 7), "##,###,###,##0.00")
        RST!tanggal = Format(datebahan, "yyyyMMdd")
        RST!REF = txtnolot(0) & "/1"
        RST!Line = Format(grid1.TextMatrix(grid1.Row, 0), "general number")
        RST!proses_ke = "1"
        RST.Update
        grid1.Row = grid1.Row + 1
    Loop
    
    'Proses Simpan Ke tabel am_usehdr
    SQL = "Select * From am_usehdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nobpb = txtnolot(0) & "/1"
        !tglbpb = datebahan
        !noorder = txtnolot(0) & "/1"
        .Update
    End With
    
    'PROSES SIMPAN KE 1 KE USE LIN
    SQL = "Select * From am_uselin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!nobpb = txtnolot(0) & "/1"
        RST!kodebarang = grid1.TextMatrix(grid1.Row, 1)
        RST!qty = Format(grid1.TextMatrix(grid1.Row, 4), "general number")
        RST!kodesatuan = grid1.TextMatrix(grid1.Row, 5)
        RST!lineitem = grid1.Row
        RST.Update
        grid1.Row = grid1.Row + 1
    Loop
    
    'proses simpan perolehan
    SQL = "select * from list_produksi_hasil where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid4.Row = 1
    Do While True
        If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = txtkodeproduk
        RST!nolot = txtnolot(0)
        RST!kode_bahan = grid4.TextMatrix(grid4.Row, 1)
        RST!Lot_bahan = grid4.TextMatrix(grid4.Row, 3)
        RST!qty_bahan = Format(grid4.TextMatrix(grid4.Row, 4), "general number")
        RST!KODE_SATUAN = grid4.TextMatrix(grid4.Row, 5)
        RST!flag_tambahan = "0"
        RST!noref = ""
        RST!tanggal = Format(datebahan, "yyyyMMdd")
        RST!proses_ke = "1"
        RST.Update
        grid4.Row = grid4.Row + 1
    Loop
        
   'PROSES SIMPAN KE HISTORI SOP
   SQL = "select * from list_historisop where 0=1"
   Set RST = New ADODB.Recordset
   RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot(0)
        !tanggal = Format(datebahan, "yyyyMMdd")
        !proses_ke = "1"
        !UserName = nmuser
        !ref1 = ""
        !ref2 = ""
        !flagprint = "1"
        .Update
    End With
    
    OBJ.Close
    MsgBox "Data is saved...!", vbInformation, AppName
    
    'PROSES CETAK SOP
    cetaksop
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Data tidak berhasil disimpan, " + Err.Description, vbCritical, AppName
End Sub

Private Sub update_mode_satu()
    On Error GoTo Err_handler:
    
    'open koneksi
    OBJ.Open dsn
    grid1.Row = 1
    Do While True
        
    Loop
    OBJ.Close
    MsgBox "Proses ubah data berhasil...!", vbInformation, AppName
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Data tidak berhasil diubah, " + Err.Description, vbCritical, AppName
End Sub

Private Sub cmdnolot_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    carisql1 = carisql1 + " where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnolot_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeproduk = hasil
    txtproduk = hasil1
    txtnolot(0) = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
    OpenHeader
    posrow = 0
    poscol = 0
    Timer1.Enabled = True
End Sub

Private Sub OpenHeader()
    On Error GoTo Err_handler:
    OBJ.Open dsn
    SQL = "select a.*,b.nama_produk from list_produksi_master a "
    SQL = SQL + " inner join list_produk_master b on a.kode_produk = b.kode_produk  "
    SQL = SQL + " where nolot='" & txtnolot(0) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    txtkodeproduk = RST!kode_produk
    txtproduk = RST!nama_produk
    txtnolot(0) = RST!nolot
    txtnoreaktor = Format(RST!noreaktor, "##,###,##0")
    txtnobpb = getnobpb(Format(Date, "yymm"))
    tg1 = RST!total_produksi
    txttotalhasilproduksi = RST!total_hasil_produksi
    txtwaktupelarutan = RST!waktu_larut
    txtwaktutambahan = RST!waktu_tambahan
    txtwaktukemasan = RST!waktu_kemasan
    txttesvisual = RST!qc_test_visual
    txtviskositas = RST!qc_viskositas
    txtsolid = RST!qc_solid
    If RST!flag_status = "0" Then
        cmbqc.text = "Lulus"
    Else
        cmbqc.text = "Tidak"
    End If
    If RST!flagprint = "1" Then
        Check2.Value = 0
        Check3.Value = 0
        Check1.Value = 0
    ElseIf RST!flagprint = "2" Then
        Check2.Value = 1
        Check3.Value = 0
        Check1.Value = 0
    ElseIf RST!flagprint = "3" Then
        Check2.Value = 1
        Check3.Value = 1
        Check1.Value = 0
    End If
    
   
   
   'versi baru sementara
    SQL = "select a.*,isnull(b.nama_bahan,d.namabarang) as namabahan,c.namasatuan from list_produksi_child a "
    SQL = SQL + "left join list_produk_child b on a.kode_produk = b.kode_produk and  "
    SQL = SQL + "a.kode_bahan=b.kode_bahan "
    SQL = SQL + "left join am_apunit c on a.kode_satuan=c.kodesatuan "
    SQL = SQL + "left join am_apitemmst d on a.kode_bahan = d.kodebarang and a.kode_satuan = d.kodeSatuan "
    SQL = SQL + "where a.nolot='" & txtnolot(0) & "' and a.flag_tambahan ='0' order by a.line "
   
   
   Set RST = OBJ.Execute(SQL)
   hapusgrid1
   grid1.Row = 1
   Dim X As Integer
   X = 1
   Do While Not RST.EOF
        If (X = RST!Line) Then
            grid1.TextMatrix(grid1.Row, 1) = RST!kode_bahan
            grid1.TextMatrix(grid1.Row, 2) = RST!namabahan
            grid1.TextMatrix(grid1.Row, 3) = RST!Lot_bahan
            grid1.TextMatrix(grid1.Row, 4) = Format(RST!qty_bahan, "##,###,###,##0.0000")
            grid1.TextMatrix(grid1.Row, 5) = RST!KODE_SATUAN
            grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
            grid1.TextMatrix(grid1.Row, 7) = Format(RST!hpp, "##,###,###,##0.00")
            grid1.Col = 0
            Set grid1.CellPicture = uncheck
            setAlternatingGrid1 grid1.Row
            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
            X = X + 1
        End If
        RST.MoveNext
    Loop
    
    SQL = "select a.kode_bahan,a.qty_bahan,b.kodesatuan,b.namabarang,c.namasatuan from list_produksi_hasil a "
    SQL = SQL + "inner join am_itemdtl b on a.kode_bahan = b.kodebarang "
    SQL = SQL + "inner join am_unit c on b.kodesatuan = c.kodesatuan "
    SQL = SQL + "where a.NOLOT= '" & txtnolot(0) & "' and a.kode_satuan = c.KodeSatuan and a.proses_ke='1'"
    Set RST = OBJ.Execute(SQL)
    hapusgrid4
    grid4.Row = 1
    Do While Not RST.EOF
        grid4.TextMatrix(grid4.Row, 1) = RST!kode_bahan
        grid4.TextMatrix(grid4.Row, 2) = RST!namabarang
        grid4.TextMatrix(grid4.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00")
        grid4.TextMatrix(grid4.Row, 5) = RST!kodesatuan
        grid4.TextMatrix(grid4.Row, 6) = RST!namasatuan
        grid4.Rows = grid4.Rows + 1
        grid4.Row = grid4.Row + 1
        RST.MoveNext
    Loop
    
    'cari total produksi
    txttotalproduksi.Value = GetTotalProduksi(txtnolot(0))
    edit_mode = True
        
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub cmdproduksi_Click()
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Anda tidak memiliki akses..! ", vbCritical, AppName
        Exit Sub
    End If
proses:
    namatabel = "produk"
    carisql1 = "select kode_produk,nama_produk from list_produk_master"
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduksi_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeproduk = hasil
    txtproduk = hasil1
    hasil = ""
    hasil1 = ""
    carisql1 = ""
    OpenDataBahanBakuUtama
    gtotal
    If Left(txtkodeproduk, 1) = "K" Then
        gbbundle.Visible = True
    End If
End Sub

Private Sub cmdstok_Click()
    Dim stokbahan As Double
    Dim namabahan As String
    Dim namasatuan As String
    Dim d As Date
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        setAlternatingGrid1Yelow grid1.Row
        If grid1.TextMatrix(grid1.Row, 4) = 0 Then GoTo lompat_sini:
        d = DateAdd("d", 1, datebahan)
        GetStokBarang Format(d, "yyyyMMdd"), grid1.TextMatrix(grid1.Row, 1), namabahan, namasatuan, stokbahan
    'MsgBox stokbahan
        If stokbahan <= 0 Or stokbahan < Format(grid1.TextMatrix(grid1.Row, 4), "general number") Then
            MsgBox "Nama Bahan : " & namabahan & Chr(13) & _
            "Stok Terakhir : " & stokbahan & " " & namasatuan, vbCritical, "Peringatan STOK"
            btnsave.Enabled = False
            setAlternatingGrid1Red grid1.Row
            Exit Sub
        End If
lompat_sini:
        setAlternatingGrid1 grid1.Row
        grid1.Row = grid1.Row + 1
    Loop
    MsgBox "Stok mencukupi...!", vbInformation, AppName
    btnsave.Enabled = True
End Sub

Private Sub cmdtambahrec_Click()
    namatabel = "Pages"
    carisql1 = "select list_historisop.tanggal,list_historisop.nolot,list_historisop.proses_ke,list_historisop.flagprint from list_historisop "
    carisql1 = carisql1 + " inner join list_produksi_master on list_historisop.nolot=list_produksi_master.nolot "
    carisql1 = carisql1 + " where list_produksi_master.nolot='" & txtnolot(0) & "' and list_produksi_master.flagprint <> '4' and list_historisop.proses_ke <> '1'"
    frmsearch.Show
End Sub

Private Sub cmdtambahrec_GotFocus()
    If hasil = "" Then Exit Sub
    opendatapages
End Sub

Private Sub opendatapages()
    On Error GoTo Err_handler:
    
    OBJ.Open dsn
    'open data histori sop
    SQL = "select * from list_historisop where nolot ='" & txtnolot(0) & "' and proses_ke ='" & hasil1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        noproses = RST!proses_ke
        noref1 = RST!ref1
        noref2 = RST!ref2
    End If
    
    namatabel = ""
    carisql1 = ""
    hasil = ""
    hasil1 = ""
    
    'membuka data bahan baku tambahan
    SQL = "select distinct a.kode_bahan,a.lot_bahan,a.qty_bahan,a.kode_satuan,a.hpp,b.nama_bahan,c.namasatuan  from list_produksi_child a "
    SQL = SQL + "left join list_produk_child b on a.kode_bahan=b.kode_bahan "
    SQL = SQL + "inner join am_apunit c on a.kode_satuan=c.kodesatuan "
    SQL = SQL + "where a.ref ='" & noref1 & "'"
    Set RST = OBJ.Execute(SQL)
    hapusgrid2
    grid2.Row = 1
    Do While Not RST.EOF
        grid2.TextMatrix(grid2.Row, 1) = RST!kode_bahan
        grid2.TextMatrix(grid2.Row, 2) = RST!nama_bahan
        grid2.TextMatrix(grid2.Row, 3) = RST!Lot_bahan
        grid2.TextMatrix(grid2.Row, 4) = Format(RST!qty_bahan, "##,###,###,##0.0000")
        grid2.TextMatrix(grid2.Row, 5) = RST!KODE_SATUAN
        grid2.TextMatrix(grid2.Row, 6) = RST!namasatuan
        grid2.TextMatrix(grid2.Row, 7) = Format(RST!hpp, "##,###,###,##0.00")
        grid2.Col = 0
        Set grid2.CellPicture = uncheck
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        RST.MoveNext
    Loop
    
    'membuka data hasil produksi
    SQL = "select a.*,b.kodesatuan,b.namabarang,c.namasatuan from list_produksi_hasil a "
    SQL = SQL + "inner join am_itemdtl b on a.kode_bahan = b.kodebarang "
    SQL = SQL + "inner join am_unit c on b.kodesatuan = c.kodesatuan "
    SQL = SQL + "where a.noref= '" & noref2 & "' and a.kode_satuan = c.KodeSatuan"
    Set RST = OBJ.Execute(SQL)
    hapusgrid4
    grid4.Row = 1
    Do While Not RST.EOF
        grid4.TextMatrix(grid4.Row, 1) = RST!kode_bahan
        grid4.TextMatrix(grid4.Row, 2) = RST!namabarang
        grid4.TextMatrix(grid4.Row, 4) = Format(RST!qty_bahan, "##,###,##0.00")
        grid4.TextMatrix(grid4.Row, 5) = RST!kodesatuan
        grid4.TextMatrix(grid4.Row, 6) = RST!namasatuan
        grid4.Rows = grid4.Rows + 1
        grid4.Row = grid4.Row + 1
        RST.MoveNext
    Loop
    
    'membuka data kemasan barang
    SQL = "select a.*,b.namabarang,c.namasatuan  from list_produksi_kemasan a "
    SQL = SQL + "inner join am_apitemmst  b on a.kode_bahan =b.kodebarang  "
    SQL = SQL + "inner join am_apunit c on a.kode_satuan=c.kodesatuan "
    SQL = SQL + "where a.noref='" & noref1 & "' "
    Set RST = OBJ.Execute(SQL)
    hapusgrid3
    grid3.Row = 1
    Do While Not RST.EOF
        grid3.TextMatrix(grid3.Row, 1) = RST!kode_bahan
        grid3.TextMatrix(grid3.Row, 2) = RST!namabarang
        grid3.TextMatrix(grid3.Row, 3) = RST!Lot_bahan
        grid3.TextMatrix(grid3.Row, 4) = Format(RST!qty_bahan, "##,###,###,##0.0000")
        grid3.TextMatrix(grid3.Row, 5) = RST!KODE_SATUAN
        grid3.TextMatrix(grid3.Row, 6) = RST!namasatuan
        grid3.TextMatrix(grid3.Row, 7) = Format(RST!hpp, "##,###,###,##0.00")
        grid3.Col = 0
        Set grid3.CellPicture = uncheck
        grid3.Rows = grid3.Rows + 1
        grid3.Row = grid3.Row + 1
        RST.MoveNext
    Loop
    edit_mode2 = True
    OBJ.Close
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub
    

Private Sub Form_Load()
    'Periksa hak akses hpp
    OBJ.Open dsn
    SQL = "Select * From LIST_USERS Where username = '" & nmuser & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!gl = "1" Then
            akses = True
        Else
            akses = False
        End If
    Else
        If nmuser = "Creator" Then akses = True
    End If
    OBJ.Close
    
    initGrid1
    setGrid1
    initGrid2
    setGrid2
    initGrid3
    setGrid3
    initGrid4
    setGrid4
    TabControl1.SelectedItem = 0
    cmbqc.AddItem "Lulus"
    cmbqc.AddItem "Tidak"
    cmbqc.text = "Lulus"
    datebahan = Date
    datedone = Date
    Datetambah = Date
    txttotalproduksi = "0.000"
    txttotalhasilproduksi = "0.000"
    txtnobpb = ""
    txttesvisual = "0"
    txtviskositas = "0"
    txtsolid = "0"
    txtwaktupelarutan = "0"
    txtwaktutambahan = "0"
    txtwaktukemasan = "0"
    txtnilai4.Visible = False
    
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
End Sub

Private Sub initGrid1()
    With grid1
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "BAHAN"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "SATUAN"
        .TextMatrix(0, 7) = "HPP"
        .TextMatrix(0, 8) = "URUT"
    End With
End Sub

Private Sub setGrid1()
    With grid1
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1200
        .ColWidth(5) = 0
        .ColWidth(6) = 750
        If akses = True Then
            .ColWidth(7) = 1500
        Else
            .ColWidth(7) = 0
        End If
        .ColWidth(8) = 750
    End With
End Sub

Private Sub initGrid2()
    With grid2
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Bahan"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "HPP"
    End With
End Sub

Private Sub setGrid2()
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1200
        .ColWidth(5) = 500
        .ColWidth(6) = 750
        If akses = True Then
            .ColWidth(7) = 1500
        Else
            .ColWidth(7) = 0
        End If
    End With
End Sub

Private Sub initGrid3()
    With grid3
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Kemasan"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "HPP"
    End With
End Sub

Private Sub setGrid3()
    With grid3
        .ColWidth(0) = 300
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColWidth(5) = 500
        .ColWidth(6) = 750
        If akses = True Then
            .ColWidth(7) = 1500
        Else
            .ColWidth(7) = 0
        End If
    End With
End Sub

Private Sub initGrid4()
    With grid4
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Produk"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "K.Kg"
        .TextMatrix(0, 8) = "Jlm. Kg"
    End With
End Sub

Private Sub setGrid4()
    With grid4
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColWidth(5) = 500
        .ColWidth(6) = 1200
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
    End With
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtkodeproduk = "" Then Exit Sub
    
    
    poscol = grid1.Col
    posrow = grid1.Row
    
    Select Case grid1.Col
        Case 0:
                If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
        Case 1:
            
        Case 3:
                If txtnolot(0) = "" Then
                    MsgBox "Nomor Lot harus diisi", vbCritical, AppName
                    Exit Sub
                End If
                
                statuslot = False 'bukan update lot wip
                lotbahan = grid1.TextMatrix(grid1.Row, 1)
                lotbahan1 = grid1.TextMatrix(grid1.Row, 2)
                lotbahan2 = grid1.TextMatrix(grid1.Row, 4)
                lotbahan3 = grid1.TextMatrix(grid1.Row, 3)
                If grid1.TextMatrix(grid1.Row, 3) <> "" Then grid1.TextMatrix(grid1.Row, 3) = ""
                frmaddlot.Show vbModal
        Case 4:
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
                txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
                txtnilai.Left = grid1.Left + grid1.CellLeft
                txtnilai.Top = grid1.Top + grid1.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
    End Select
    
End Sub

Private Sub grid1_EnterCell()
    Select Case grid1.Col
        'Case 3:
                'If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                'txtnolot1.Width = grid1.ColWidth(grid1.Col) - 40
                'txtnolot1 = grid1.TextMatrix(grid1.Row, grid1.Col)
                'txtnolot1.Left = grid1.Left + grid1.CellLeft
                'txtnolot1.Top = grid1.Top + grid1.CellTop
                'txtnolot1.Visible = True
                'txtnolot1.SetFocus
        Case 4:
                If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
                txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
                txtnilai.Left = grid1.Left + grid1.CellLeft
                txtnilai.Top = grid1.Top + grid1.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
    End Select
End Sub

Private Sub grid1_GotFocus()
    On Error Resume Next
    If hasil = "" Then Exit Sub
    Select Case grid1.Col
        Case 1:
            grid1.TextMatrix(grid1.Row, 1) = hasil
            grid1.TextMatrix(grid1.Row, 2) = hasil2
            grid1.TextMatrix(grid1.Row, 3) = ""
            grid1.TextMatrix(grid1.Row, 4) = "0.0000"
            grid1.TextMatrix(grid1.Row, 5) = hasil3
                    
                    
            'cari satuan
            SQL = "select  initial from am_apunit where kodesatuan ='" & hasil3 & "'"
            OBJ.Open dsn
            Set RST = OBJ.Execute(SQL)
                    
            grid1.TextMatrix(grid1.Row, 6) = RST!initial
            OBJ.Close
                    
            grid1.Col = 0
            Set grid1.CellPicture = uncheck
                    
            namatabel = ""
            carisql1 = ""
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
        Case 3:
            grid1.TextMatrix(grid1.Row, 3) = hasil
            grid1.TextMatrix(grid1.Row, 7) = hasil1
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub grid2_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If txtkodeproduk = "" Then Exit Sub
    If Check2.Value = 0 Then
        MsgBox "Silahkan centang terlebih dahulu Bahan Baku dan Perolehan ", vbInformation, AppName
        Exit Sub
    End If
    poscol = grid2.Col
    posrow = grid2.Row
    
    Select Case grid2.Col
        Case 0:
                If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow2
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
                End If
        Case 1:
            namatabel = "Bahan Tambahan"
            carisql1 = "select distinct kode_bahan,nama_bahan,inisial,kode_satuan from list_produk_child"
            frmsearch.Show vbModal
        Case 3:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                txtnolot1.Width = grid2.ColWidth(grid2.Col) - 40
                txtnolot2 = grid2.TextMatrix(grid2.Row, grid2.Col)
                txtnolot2.Left = grid2.Left + grid2.CellLeft
                txtnolot2.Top = grid2.Top + grid2.CellTop
                txtnolot2.Visible = True
                txtnolot2.SetFocus
        Case 4:
                If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                txtnilai1.Width = grid2.ColWidth(grid2.Col) - 40
                txtnilai1 = grid2.TextMatrix(grid2.Row, grid2.Col)
                txtnilai1.Left = grid2.Left + grid2.CellLeft
                txtnilai1.Top = grid2.Top + grid2.CellTop + 20
                txtnilai1.Visible = True
                txtnilai1.SetFocus
    End Select
End Sub

Private Sub grid3_Click()
    If grid3.MouseRow = 0 Then Exit Sub
    If txtkodeproduk = "" Then Exit Sub
    If Check3.Value = 0 Then
        MsgBox "Silahkan centang terlebih dahulu kemasan..!", vbInformation, AppName
        Exit Sub
    End If
    poscol = grid3.Col
    posrow = grid3.Row
    
    Select Case grid3.Col
        Case 0:
                If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Sub
                If grid3.CellPicture = uncheck Then
                Set grid3.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid3.CellPicture = uncheck
                    hapusrow3
                    Exit Sub
                End If
                Set grid3.CellPicture = uncheck
                End If
        Case 1:
                    carisql1 = "select kodebarang, namabarang from am_apitemmst"
                    namatabel = "Bahan Baku"
                    frmsearch.Show vbModal
        Case 4:
                If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Sub
                txtnilai3.Width = grid3.ColWidth(grid3.Col) - 40
                txtnilai3 = grid3.TextMatrix(grid3.Row, grid3.Col)
                txtnilai3.Left = grid3.Left + grid3.CellLeft
                txtnilai3.Top = grid3.Top + grid3.CellTop + 20
                txtnilai3.Visible = True
                txtnilai3.SetFocus
    End Select
End Sub

Private Sub grid3_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid3.Col
        Case 1:
                    grid3.TextMatrix(grid3.Row, 1) = hasil
                    grid3.TextMatrix(grid3.Row, 2) = hasil1
                    'cari satuan
                    SQL = "select kodesatuan from am_apitemmst  "
                    SQL = SQL + "where kodebarang='" & hasil & "'"
                    OBJ.Open dsn
                    Set RST = New ADODB.Recordset
                    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
                    grid3.TextMatrix(grid3.Row, 4) = "0.0000"
                    grid3.TextMatrix(grid3.Row, 5) = RST!kodesatuan
                    
                    'cari nama satuan
                    SQL = "select * from am_apunit where kodesatuan ='" & grid3.TextMatrix(grid3.Row, 5) & "'"
                    Set RST = OBJ.Execute(SQL)
                    grid3.TextMatrix(grid3.Row, 6) = RST!namasatuan

                    OBJ.Close
                    
                    grid3.Col = 0
                    Set grid3.CellPicture = uncheck
                    grid3.Rows = grid3.Rows + 1
                    grid3.Row = grid3.Row + 1
                    hasil = ""
                    hasil1 = ""
                    hasil2 = ""
                    hasil3 = ""
                    namatabel = ""
                    carisql1 = ""
    End Select
End Sub

Private Sub grid4_Click()
    If grid4.MouseRow = 0 Then Exit Sub
    If txtkodeproduk = "" Then Exit Sub
    If Check2.Value = 0 Then Exit Sub
    poscol = grid4.Col
    posrow = grid4.Row
    
    Select Case grid4.Col
        Case 0:
                If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Sub
                If grid4.CellPicture = uncheck Then
                Set grid4.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid4.CellPicture = uncheck
                    hapusrow4
                    Exit Sub
                End If
                Set grid4.CellPicture = uncheck
                End If
        Case 1:
                    carisql1 = "select am_itemdtl.kodebarang, am_itemdtl.namabarang,list_produk_hasil.kode_satuan,am_unit.namasatuan from am_itemdtl  "
                    carisql1 = carisql1 + " inner join list_produk_hasil on am_itemdtl.kodebarang=list_produk_hasil.kode_barang_jadi "
                    carisql1 = carisql1 + " and am_itemdtl.kodesatuan = list_produk_hasil.kode_satuan and list_produk_hasil.kode_produk='" & txtkodeproduk & "' "
                    carisql1 = carisql1 + " inner join am_unit on list_produk_hasil.kode_satuan= am_unit.kodesatuan "
                    namatabel = "Barang Jadi"
                    frmsearch.Show vbModal
        Case 4:
                If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Sub
                txtnilai4.Width = grid4.ColWidth(grid4.Col) - 40
                txtnilai4 = grid4.TextMatrix(grid4.Row, grid4.Col)
                txtnilai4.Left = grid4.Left + grid4.CellLeft
                txtnilai4.Top = grid4.Top + grid4.CellTop + 20
                txtnilai4.Visible = True
                txtnilai4.SetFocus
    End Select
End Sub

Private Sub grid4_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid4.Col
        Case 1:
                    grid4.TextMatrix(grid4.Row, 1) = hasil
                    grid4.TextMatrix(grid4.Row, 2) = hasil1
                    
                    OBJ.Open dsn
                    
                    grid4.TextMatrix(grid4.Row, 4) = "0.00"
                    grid4.TextMatrix(grid4.Row, 5) = hasil2
    
                    'cek konversi ke kilogram
                    SQL = "select * from am_itemdtl where kodebarang='" & grid4.TextMatrix(grid4.Row, 1) & "' and kodesatuan ='" & grid4.TextMatrix(grid4.Row, 5) & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then
                        If RST!konversi = 0 Then
                            OBJ.Close
                            MsgBox "Silahkan konfigurasi kg set terlebih dahulu...!", vbCritical, AppName
                            hapusgrid4
                            Exit Sub
                        End If
                        grid4.TextMatrix(grid4.Row, 7) = RST!konversi
                        grid4.TextMatrix(grid4.Row, 8) = "0.00"
                    End If
                    
                    'cari nama satuan
                    SQL = "select * from am_unit where kodesatuan ='" & grid4.TextMatrix(grid4.Row, 5) & "'"
                    Set RST = OBJ.Execute(SQL)
                    grid4.TextMatrix(grid4.Row, 6) = RST!namasatuan

                    OBJ.Close
                    
                    grid4.Col = 0
                    Set grid4.CellPicture = uncheck
                    grid4.Rows = grid4.Rows + 1
                    grid4.Row = grid4.Row + 1
                    hasil = ""
                    hasil1 = ""
                    hasil2 = ""
                    hasil3 = ""
                    namatabel = ""
                    carisql1 = ""
    End Select
End Sub

Private Sub Timer1_Timer()
    If lbleditmode.Visible = False Then
        lbleditmode.Visible = True
    Else
        lbleditmode.Visible = False
    End If
End Sub

Private Sub txtbundle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid1.Row = 1
        Do While True
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            grid1.TextMatrix(grid1.Row, 4) = Format(grid1.TextMatrix(grid1.Row, 4) * txtbundle, "##,####,###,##0.000")
            grid1.TextMatrix(grid1.Row, 7) = Format(grid1.TextMatrix(grid1.Row, 7) * txtbundle, "##,####,###,##0.00")
            
            grid1.Row = grid1.Row + 1
        Loop
        gbbundle.Visible = False
        MsgBox "Qty telah dikalikan " & txtbundle & " Bundle" & vbCrLf & "Mohon periksa ketersediaan stok terlebih dahulu" & vbCrLf & "dengan klik tombol CEK STOK", vbExclamation, AppName
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim konvToKg As Double
    If KeyAscii = 13 Then
        If (Left(grid1.TextMatrix(grid1.Row, 1), 3) = "L11" Or Left(grid1.TextMatrix(grid1.Row, 1), 3) = "K05") And grid1.TextMatrix(grid1.Row, 5) <> "002" Then
            'konversi satuan wip ke kg
            OBJ.Open dsn
            SQL = "Select a.Nilai,a.KodeSatuanKonv,b.NamaSatuan From am_apunit_konversi a"
            SQL = SQL + " inner join am_apunit b on a.KodeSatuanKonv = b.KodeSatuan"
            SQL = SQL + " Where a.kdbrg = '" & grid1.TextMatrix(grid1.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST.EOF Then
                OBJ.Close
                MsgBox "Konversi satuan tidak ditemukan", vbCritical, AppName
                Exit Sub
            End If
            grid1.TextMatrix(grid1.Row, 5) = RST!kodesatuankonv
            grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
            grid1.TextMatrix(grid1.Row, 4) = Format(txtnilai.text * RST!nilai, "#,##0.000")
            konvToKg = Format(txtnilai.text * RST!nilai, "#,##0.000")
            grid1.TextMatrix(grid1.Row, 7) = Format(getHPP(grid1.TextMatrix(grid1.Row, 1), stokbahan, konvToKg), "##,####,###,##0.00")
            grid1.SetFocus
            txttotalproduksi = txttotalproduksi + (txtnilai.text * RST!nilai)
            MsgBox "Satuan telah dikonversi ke Kilogram", vbInformation, "AppName"
            OBJ.Close
        Else
            grid1.TextMatrix(grid1.Row, 4) = txtnilai.text
            konvToKg = txtnilai.text
            grid1.TextMatrix(grid1.Row, 7) = Format(getHPP(grid1.TextMatrix(grid1.Row, 1), stokbahan, konvToKg), "##,####,###,##0.00")
            grid1.SetFocus
            txttotalproduksi = txttotalproduksi + txtnilai
        End If
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub


Private Sub txtnilai1_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    'Dim kode_bahan As String
    'Dim nama_bahan As String
    'Dim nama_satuan As String
    Dim hppbahan As Double
    Dim konvToKg As Double
    If KeyAscii = 13 Then
        'GetStokBarang Format(datedone, "yyyyMMdd"), kode_bahan, nama_bahan, nama_satuan, stokbahan
        GetStokBarang Format(datebahan, "yyyyMMdd"), grid2.TextMatrix(grid2.Row, 1), , , stokbahan
        If stokbahan <= 0 Or stokbahan < txtnilai1.Value Then
            'MsgBox datebahan & grid2.TextMatrix(grid2.Row, 1) & stokbahan
            MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
            Exit Sub
        Else
            'cek konversi to kg unit
            OBJ1.Open dsn
                SQL1 = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    konvToKg = txtnilai1 / RST1!nilai
                Else
                    konvToKg = txtnilai1
                End If
            OBJ1.Close
        End If
        grid2.TextMatrix(grid2.Row, 4) = txtnilai1.text
        grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, konvToKg), "##,####,###,##0.00")

        totalg1
        totalg2
        gtotal
        grid2.SetFocus
    End If
End Sub

Private Sub txtnilai1_LostFocus()
    txtnilai1.Visible = False
End Sub

Private Sub txtnilai3_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim hppbahan As Double
    If KeyAscii = 13 Then
        GetStokBarang Format(datedone, "yyyyMMdd"), grid3.TextMatrix(grid3.Row, 1), , , stokbahan
        
        If stokbahan <= 0 Or stokbahan <= txtnilai3.Value Then
            MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
            Exit Sub
        End If
        grid3.TextMatrix(grid3.Row, 4) = txtnilai3.text
        grid3.TextMatrix(grid3.Row, 7) = Format(getHPP(grid3.TextMatrix(grid3.Row, 1), stokbahan, txtnilai3.Value), "##,###,###,##0.00")
        grid3.SetFocus
    End If
End Sub

Private Sub txtnilai3_LostFocus()
    txtnilai3.Visible = False
End Sub

Private Sub txtnilai4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid4.TextMatrix(grid4.Row, 4) = txtnilai4.text
        grid4.TextMatrix(grid4.Row, 8) = Val(txtnilai4.text) * Val(grid4.TextMatrix(grid4.Row, 7))
        grid4.SetFocus
    End If
End Sub

Private Sub txtnilai4_LostFocus()
    txtnilai4.Visible = False
End Sub

Private Sub txtnolot_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        opendata
        openDataUpdate
        gtotal
    End If
End Sub

Private Sub txtnolot1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtnolot1 = "" Then Exit Sub
        grid1.TextMatrix(grid1.Row, 3) = txtnolot1
        grid1.SetFocus
    End If
End Sub

Private Sub txtnolot1_LostFocus()
    txtnolot1.Visible = False
End Sub

Private Sub hapusrow1()
    grid1.TextMatrix(grid1.Row, 1) = ""
    grid1.TextMatrix(grid1.Row, 2) = ""
    grid1.TextMatrix(grid1.Row, 3) = ""
    grid1.TextMatrix(grid1.Row, 4) = ""
    grid1.TextMatrix(grid1.Row, 5) = ""
    grid1.TextMatrix(grid1.Row, 6) = ""
    grid1.TextMatrix(grid1.Row, 7) = ""
    grid1.TextMatrix(grid1.Row, 8) = ""
    
    Do While True
        If grid1.TextMatrix(grid1.Row + 1, 1) = "" Then
            grid1.TextMatrix(grid1.Row, 1) = ""
            grid1.TextMatrix(grid1.Row, 2) = ""
            grid1.TextMatrix(grid1.Row, 3) = ""
            grid1.TextMatrix(grid1.Row, 4) = ""
            grid1.TextMatrix(grid1.Row, 5) = ""
            grid1.TextMatrix(grid1.Row, 6) = ""
            grid1.TextMatrix(grid1.Row, 7) = ""
            grid1.TextMatrix(grid1.Row, 8) = ""
            Exit Do
        End If
        grid1.TextMatrix(grid1.Row, 1) = grid1.TextMatrix(grid1.Row + 1, 1)
        grid1.TextMatrix(grid1.Row, 2) = grid1.TextMatrix(grid1.Row + 1, 2)
        grid1.TextMatrix(grid1.Row, 3) = grid1.TextMatrix(grid1.Row + 1, 3)
        grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row + 1, 4)
        grid1.TextMatrix(grid1.Row, 5) = grid1.TextMatrix(grid1.Row + 1, 5)
        grid1.TextMatrix(grid1.Row, 6) = grid1.TextMatrix(grid1.Row + 1, 6)
        grid1.TextMatrix(grid1.Row, 7) = grid1.TextMatrix(grid1.Row + 1, 7)
        grid1.TextMatrix(grid1.Row, 8) = grid1.TextMatrix(grid1.Row + 1, 8)
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = grid1.Rows - 1
    grid1.Col = 0
    Set grid1.CellPicture = blank
End Sub

Private Sub hapusrow2()
    grid2.TextMatrix(grid2.Row, 1) = ""
    grid2.TextMatrix(grid2.Row, 2) = ""
    grid2.TextMatrix(grid2.Row, 3) = ""
    grid2.TextMatrix(grid2.Row, 4) = ""
    grid2.TextMatrix(grid2.Row, 5) = ""
    grid2.TextMatrix(grid2.Row, 6) = ""
    grid2.TextMatrix(grid2.Row, 7) = ""
    Do While True
        If grid2.TextMatrix(grid2.Row + 1, 1) = "" Then
            grid2.TextMatrix(grid2.Row, 1) = ""
            grid2.TextMatrix(grid2.Row, 2) = ""
            grid2.TextMatrix(grid2.Row, 3) = ""
            grid2.TextMatrix(grid2.Row, 4) = ""
            grid2.TextMatrix(grid2.Row, 5) = ""
            grid2.TextMatrix(grid2.Row, 6) = ""
            grid2.TextMatrix(grid2.Row, 7) = ""
            Exit Do
        End If
        grid2.TextMatrix(grid2.Row, 1) = grid2.TextMatrix(grid2.Row + 1, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row + 1, 2)
        grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row + 1, 3)
        grid2.TextMatrix(grid2.Row, 4) = grid2.TextMatrix(grid2.Row + 1, 4)
        grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row + 1, 5)
        grid2.TextMatrix(grid2.Row, 6) = grid2.TextMatrix(grid2.Row + 1, 6)
        grid2.TextMatrix(grid2.Row, 7) = grid2.TextMatrix(grid2.Row + 1, 7)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = blank
End Sub

Private Sub hapusrow3()
    grid3.TextMatrix(grid3.Row, 1) = ""
    grid3.TextMatrix(grid3.Row, 2) = ""
    grid3.TextMatrix(grid3.Row, 3) = ""
    grid3.TextMatrix(grid3.Row, 4) = ""
    grid3.TextMatrix(grid3.Row, 5) = ""
    grid3.TextMatrix(grid3.Row, 6) = ""
    grid3.TextMatrix(grid3.Row, 7) = ""
    Do While True
        If grid3.TextMatrix(grid3.Row + 1, 1) = "" Then
            grid3.TextMatrix(grid3.Row, 1) = ""
            grid3.TextMatrix(grid3.Row, 2) = ""
            grid3.TextMatrix(grid3.Row, 3) = ""
            grid3.TextMatrix(grid3.Row, 4) = ""
            grid3.TextMatrix(grid3.Row, 5) = ""
            grid3.TextMatrix(grid3.Row, 6) = ""
            grid3.TextMatrix(grid3.Row, 7) = ""
            Exit Do
        End If
        grid3.TextMatrix(grid3.Row, 1) = grid3.TextMatrix(grid3.Row + 1, 1)
        grid3.TextMatrix(grid3.Row, 2) = grid3.TextMatrix(grid3.Row + 1, 2)
        grid3.TextMatrix(grid3.Row, 3) = grid3.TextMatrix(grid3.Row + 1, 3)
        grid3.TextMatrix(grid3.Row, 4) = grid3.TextMatrix(grid3.Row + 1, 4)
        grid3.TextMatrix(grid3.Row, 5) = grid3.TextMatrix(grid3.Row + 1, 5)
        grid3.TextMatrix(grid3.Row, 6) = grid3.TextMatrix(grid3.Row + 1, 6)
        grid3.TextMatrix(grid3.Row, 7) = grid3.TextMatrix(grid3.Row + 1, 7)
        grid3.Row = grid3.Row + 1
    Loop
    grid3.Rows = grid3.Rows - 1
    grid3.Col = 0
    Set grid3.CellPicture = blank
End Sub

Private Sub hapusrow4()
    grid4.TextMatrix(grid4.Row, 1) = ""
    grid4.TextMatrix(grid4.Row, 2) = ""
    grid4.TextMatrix(grid4.Row, 3) = ""
    grid4.TextMatrix(grid4.Row, 4) = ""
    grid4.TextMatrix(grid4.Row, 5) = ""
    grid4.TextMatrix(grid4.Row, 6) = ""
    grid4.TextMatrix(grid4.Row, 7) = ""
    grid4.TextMatrix(grid4.Row, 8) = ""
    
    Do While True
        If grid4.TextMatrix(grid4.Row + 1, 1) = "" Then
            grid4.TextMatrix(grid4.Row, 1) = ""
            grid4.TextMatrix(grid4.Row, 2) = ""
            grid4.TextMatrix(grid4.Row, 3) = ""
            grid4.TextMatrix(grid4.Row, 4) = ""
            grid4.TextMatrix(grid4.Row, 5) = ""
            grid4.TextMatrix(grid4.Row, 6) = ""
            grid4.TextMatrix(grid4.Row, 7) = ""
            grid4.TextMatrix(grid4.Row, 8) = ""
            Exit Do
        End If
        grid4.TextMatrix(grid4.Row, 1) = grid4.TextMatrix(grid4.Row + 1, 1)
        grid4.TextMatrix(grid4.Row, 2) = grid4.TextMatrix(grid4.Row + 1, 2)
        grid4.TextMatrix(grid4.Row, 3) = grid4.TextMatrix(grid4.Row + 1, 3)
        grid4.TextMatrix(grid4.Row, 4) = grid4.TextMatrix(grid4.Row + 1, 4)
        grid4.TextMatrix(grid4.Row, 5) = grid4.TextMatrix(grid4.Row + 1, 5)
        grid4.TextMatrix(grid4.Row, 6) = grid4.TextMatrix(grid4.Row + 1, 6)
        grid4.TextMatrix(grid4.Row, 7) = grid4.TextMatrix(grid4.Row + 1, 7)
        grid4.TextMatrix(grid4.Row, 8) = grid4.TextMatrix(grid4.Row + 1, 8)
        grid4.Row = grid4.Row + 1
    Loop
    grid4.Rows = grid4.Rows - 1
    grid4.Col = 0
    Set grid4.CellPicture = blank
End Sub

Private Sub grid2_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid2.Col
        Case 1:
                    grid2.TextMatrix(grid2.Row, 1) = hasil
                    grid2.TextMatrix(grid2.Row, 2) = hasil2
                    grid2.TextMatrix(grid2.Row, 3) = ""
                    grid2.TextMatrix(grid2.Row, 4) = "0.0000"
                    grid2.TextMatrix(grid2.Row, 5) = hasil3
                    
                    
                    'cari satuan
                    SQL = "select  initial from am_apunit where kodesatuan ='" & hasil3 & "'"
                    OBJ.Open dsn
                    Set RST = OBJ.Execute(SQL)
                    
                    grid2.TextMatrix(grid2.Row, 6) = RST!initial
                    OBJ.Close
                    
                    grid2.Col = 0
                    Set grid2.CellPicture = uncheck
                    
                    namatabel = ""
                    carisql1 = ""
                    hasil = ""
                    hasil1 = ""
                    hasil2 = ""
                    hasil3 = ""
                    grid2.Rows = grid2.Rows + 1
                    grid2.Row = grid2.Row + 1
    End Select
End Sub

Private Sub txtnolot2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid2.TextMatrix(grid2.Row, 3) = txtnolot2
        grid2.SetFocus
    End If
End Sub

Private Sub txtnolot2_LostFocus()
    txtnolot2.Visible = False
End Sub

Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.TextMatrix(grid1.Row, 7) = ""
        grid1.TextMatrix(grid1.Row, 8) = ""

        grid1.Col = 0
        Set grid1.CellPicture = blank
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    setGrid1
End Sub

Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    setGrid2
End Sub

Private Sub hapusgrid3()
    grid3.Row = 1
    Do While True
        If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
        grid3.TextMatrix(grid3.Row, 1) = ""
        grid3.TextMatrix(grid3.Row, 2) = ""
        grid3.TextMatrix(grid3.Row, 3) = ""
        grid3.TextMatrix(grid3.Row, 4) = ""
        grid3.TextMatrix(grid3.Row, 5) = ""
        grid3.TextMatrix(grid3.Row, 6) = ""
        grid3.TextMatrix(grid3.Row, 7) = ""
        grid3.Col = 0
        Set grid3.CellPicture = blank
        grid3.Row = grid3.Row + 1
    Loop
    grid3.Rows = 2
    setGrid3
End Sub

Private Sub hapusgrid4()
    grid4.Row = 1
    Do While True
        If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
        grid4.TextMatrix(grid4.Row, 1) = ""
        grid4.TextMatrix(grid4.Row, 2) = ""
        grid4.TextMatrix(grid4.Row, 3) = ""
        grid4.TextMatrix(grid4.Row, 4) = ""
        grid4.TextMatrix(grid4.Row, 5) = ""
        grid4.TextMatrix(grid4.Row, 6) = ""
        grid4.TextMatrix(grid4.Row, 7) = ""
        grid4.TextMatrix(grid4.Row, 8) = ""
        grid4.Col = 0
        Set grid4.CellPicture = blank
        grid4.Row = grid4.Row + 1
    Loop
    grid4.Rows = 2
    setGrid4
End Sub


Private Sub openDataUpdate()
    On Error GoTo Err_handler:
    If txtnolot(0) = "" Then Exit Sub
    'Opendata am_bpbhdr
    OBJ.Open dsn
    SQL = "Select * From am_bpbhdr Where noref = '" & txtnolot(0) & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
    With RST
        txtnobpb = !nobpb
        datebahan = !dateentry
        datedone = !dateupdate
    End With
    End If
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub opendata()
    On Error GoTo Err_handler:
    If txtnolot(0) = "" Then Exit Sub

    OBJ.Open dsn
    totalg1
    totalg2
    edit_mode = True
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    
End Sub

Private Sub OpenDataBahanBakuUtama()
    On Error GoTo Err_handler:
    Dim stok_bahan_baku_utama As Double
    Dim hpp_bahan_baku_utama As Double
    Dim nama_bahan As String
    Dim nama_satuan As String
    Dim konvToKg As Double
    Dim d As Date
    
    
    SQL = "select a.*, b.namasatuan from list_produk_child a inner join  am_apunit b on a.kode_satuan= b.kodesatuan"
    SQL = SQL + " where a.kode_produk='" & txtkodeproduk & "' order by a.line"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    hapusgrid1
    grid1.Row = 1
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 0) = RST!Line
        grid1.TextMatrix(grid1.Row, 1) = RST!kode_bahan
        grid1.TextMatrix(grid1.Row, 2) = RST!inisial
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = Format(RST!qty, "##,###,###,##0.000")
        grid1.TextMatrix(grid1.Row, 5) = RST!KODE_SATUAN
        grid1.TextMatrix(grid1.Row, 6) = RST!namasatuan
        d = DateAdd("d", 1, datebahan)
        GetStokBarang Format(d, "yyyyMMdd"), RST!kode_bahan, nama_bahan, nama_satuan, stok_bahan_baku_utama
        If stok_bahan_baku_utama < 0 Or stok_bahan_baku_utama < Format(grid1.TextMatrix(grid1.Row, 4), "general number") Then
            MsgBox "Nama Bahan : " & nama_bahan & Chr(13) & _
            "Stok Terakhir : " & stok_bahan_baku_utama & " " & nama_satuan, vbCritical, "Peringatan STOK"
            btnsave.Enabled = False
            grid1.TextMatrix(grid1.Row, 7) = "0.00"
            btnsave.Enabled = False
            setAlternatingGrid1Red grid1.Row
        Else
            'cek konversi to kg unit
            OBJ1.Open dsn
                SQL1 = "Select * from am_apunit_konversi Where kdbrg ='" & grid1.TextMatrix(grid1.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    konvToKg = grid1.TextMatrix(grid1.Row, 4) / RST1!nilai
                Else
                    konvToKg = RST!qty
                End If
            OBJ1.Close
            grid1.TextMatrix(grid1.Row, 7) = Format(getHPP(RST!kode_bahan, stok_bahan_baku_utama, konvToKg), "##,###,###,###,##0.00")
            setAlternatingGrid1 grid1.Row
        End If
        grid1.TextMatrix(grid1.Row, 8) = RST!Line
        
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    totalg1
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub OpenDataHasilProduksi()
    On Error GoTo Err_handler:
    SQL = "select  a.kode_produk,a.kode_barang_jadi,a.kode_satuan,"
    SQL = SQL + "b.namabarang ,c.namasatuan  "
    SQL = SQL + "from list_produk_hasil a "
    SQL = SQL + "inner join am_itemdtl b on a.kode_barang_jadi= b.kodebarang and a.kode_satuan=b.kodesatuan "
    SQL = SQL + "inner join am_unit c on a.kode_satuan = c.kodesatuan "
    SQL = SQL + "where a.kode_produk ='" & txtkodeproduk & "'"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    hapusgrid4
    grid4.Row = 1
    Do While Not RST.EOF
        grid4.TextMatrix(grid4.Row, 1) = RST!kode_barang_jadi
        grid4.TextMatrix(grid4.Row, 2) = RST!namabarang
        grid4.TextMatrix(grid4.Row, 4) = "0.00"
        grid4.TextMatrix(grid4.Row, 5) = RST!KODE_SATUAN
        grid4.TextMatrix(grid4.Row, 6) = RST!namasatuan
        grid4.Rows = grid4.Rows + 1
        grid4.Row = grid4.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
    txttotalhasilproduksi = "0.0000"
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Function setAlternatingGrid1(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If (i Mod 2) = 0 Then
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HFFFFFF
        Next
    End If
End Function

Private Function setAlternatingGrid1Red(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = vbRed
    Next
End Function

Private Function setAlternatingGrid1Yelow(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To 7
        grid1.Col = j
        grid1.CellBackColor = vbYellow
    Next
End Function


Private Function setAlternatingGrid4(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If (i Mod 2) = 0 Then
        For j = 0 To grid4.Cols - 1
        grid4.Col = j
        grid4.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid4.Cols - 1
        grid4.Col = j
        grid4.CellBackColor = &HFFFFFF
        Next
    End If
End Function

Private Sub openDataPerolehan()
    On Error GoTo Err_handler:
    OBJ.Open dsn
    SQL = "select * from list_produk_hasil where kode_produk='" & txtkodeproduk & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        RST.MoveNext
    Loop
    
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub totalg1()
On Error Resume Next
'TOTAL GRID1
    grid1.Row = 1
    tg1 = 0
    Do While True
        DoEvents
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            tg1 = CDbl(Format(grid1.TextMatrix(grid1.Row, 4), "general number") + CDbl(tg1))
                grid1.Row = grid1.Row + 1
    Loop
        tg1 = Format(tg1, "##,###,##0.0000")
End Sub

Private Sub totalg2()
On Error Resume Next
'TOTAL GRID2
    grid2.Row = 1
    tg2 = 0
    Do While True
        DoEvents
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
            tg2 = CDbl(Format(grid2.TextMatrix(grid2.Row, 4), "general Number") + CDbl(tg2))
                grid2.Row = grid2.Row + 1
    Loop
        tg2 = Format(tg2, "##,###,##0.0000")
End Sub

Private Sub gtotal()
    txttotalproduksi = CDbl(Format(tg1, "##,###,##0.0000")) + CDbl(Format(tg2, "##,###,##0.0000"))
End Sub

Private Sub cekbase()
On Error Resume Next
    Dim kdlotstok As String
    grid1.Row = 1
    Do While True
        DoEvents
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
            If grid1.TextMatrix(grid1.Row, 3) <> "" Then
                If MsgBox("Penggunaan stok base terdeteksi" + vbCrLf + _
                "Klik OK lalu batalkan penggunaan base pada SOP ini", vbExclamation, AppName) = vbOK Then GoTo batal:
            End If
                grid1.Row = grid1.Row + 1
    Loop
    Unload Me
batal:
End Sub
Private Sub cetaksop()
    On Error GoTo Err_handler:
    Dim cetak_ke As Integer
    
    OBJ.Open dsn
    SQL = "Select count(nolot)as jml from list_historicetaksop where nolot ='" & txtnolot(0) & "'"
    Set RST = OBJ.Execute(SQL)
    cetak_ke = RST!jml
    cetak_ke = cetak_ke + 1
    
    SQL = "select * from list_historicetaksop where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot(0)
        !tanggal = Format(Date, "yyyy/MM/dd")
        !cetakan = cetak_ke
        !keterangan = txtnolot(0).text
        !UserName = nmuser
        .Update
    End With
    OBJ.Close
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetaksop1)"
    '==
    If Left(txtkodeproduk, 1) = "L" Then
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_sop.rpt"
    ElseIf Left(txtkodeproduk, 1) = "K" Then
        crystal.ReportFileName = AppPath & "\reports\produksi\cetak_sopk.rpt"
    End If
    crystal.ParameterFields(0) = "@nolot;" & txtnolot(0).text & ";true"
    crystal.ParameterFields(1) = "@username;" & nmuser & ";true"
    crystal.ParameterFields(2) = "@cetakan;" & "CETAKAN KE " & Str(cetak_ke) & ";true"
    crystal.ParameterFields(3) = "@kode;" & Cheap_Decrypt(txtnolot(0)) & Cheap_Decrypt(Str(Trim(cetak_ke))) & ";true"
    crystal.ParameterFields(4) = "@nolot2;" & txtnolot(0).text & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    btnNew_Click
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As Control
  
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSHFlexGrid Then
          If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        End If
    Next ctl
End Sub
