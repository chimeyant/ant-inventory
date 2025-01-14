VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frme_tol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-TOL CARD"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4665
      Left            =   -15
      TabIndex        =   0
      Top             =   -30
      Width           =   9390
      _Version        =   851970
      _ExtentX        =   16563
      _ExtentY        =   8229
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Transaksi E-TOL"
      Item(0).ControlCount=   28
      Item(0).Control(0)=   "cmbpilih"
      Item(0).Control(1)=   "txttol"
      Item(0).Control(2)=   "txtpemakai"
      Item(0).Control(3)=   "txtket"
      Item(0).Control(4)=   "txtncash"
      Item(0).Control(5)=   "cmdnomor"
      Item(0).Control(6)=   "date1"
      Item(0).Control(7)=   "Label7"
      Item(0).Control(8)=   "Label4"
      Item(0).Control(9)=   "Label3"
      Item(0).Control(10)=   "Label5"
      Item(0).Control(11)=   "Label6"
      Item(0).Control(12)=   "Label8"
      Item(0).Control(13)=   "lblterbilang"
      Item(0).Control(14)=   "Label10"
      Item(0).Control(15)=   "Label1"
      Item(0).Control(16)=   "txtnoknd"
      Item(0).Control(17)=   "txtnocard"
      Item(0).Control(18)=   "Label13"
      Item(0).Control(19)=   "Label14"
      Item(0).Control(20)=   "txtkdtrans"
      Item(0).Control(21)=   "dtpjam"
      Item(0).Control(22)=   "Label15"
      Item(0).Control(23)=   "Label16"
      Item(0).Control(24)=   "cbedit"
      Item(0).Control(25)=   "cbdelete"
      Item(0).Control(26)=   "txtsaldo"
      Item(0).Control(27)=   "Label17"
      Item(1).Caption =   "Master E-TOL"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "Label2"
      Item(1).Control(1)=   "Label9"
      Item(1).Control(2)=   "Label11"
      Item(1).Control(3)=   "txtcard"
      Item(1).Control(4)=   "Label12"
      Item(1).Control(5)=   "txtmobil"
      Item(1).Control(6)=   "txtuser"
      Item(1).Control(7)=   "grid"
      Item(2).Caption =   "Laporan Transaksi E-TOL"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "optcard"
      Item(2).Control(1)=   "optdate"
      Item(2).Control(2)=   "dtptgl1"
      Item(2).Control(3)=   "dtptgl2"
      Item(2).Control(4)=   "Shape1"
      Item(2).Control(5)=   "btnview"
      Item(2).Control(6)=   "txtkartu"
      Item(2).Control(7)=   "btnkartu"
      Item(2).Control(8)=   "Lfrom"
      Item(2).Control(9)=   "lto"
      Begin XtremeSuiteControls.CheckBox cbedit 
         Height          =   240
         Left            =   -64345
         TabIndex        =   45
         Top             =   795
         Visible         =   0   'False
         Width           =   600
         _Version        =   851970
         _ExtentX        =   1058
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Edit"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optcard 
         Height          =   345
         Left            =   255
         TabIndex        =   37
         Top             =   1245
         Width           =   1920
         _Version        =   851970
         _ExtentX        =   3387
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "By E-Tol Card"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optdate 
         Height          =   345
         Left            =   255
         TabIndex        =   36
         Top             =   810
         Width           =   1920
         _Version        =   851970
         _ExtentX        =   3387
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "By Date"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin VB.TextBox txtkdtrans 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   -62110
         TabIndex        =   30
         Top             =   750
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtuser 
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
         Height          =   300
         Left            =   -68485
         MaxLength       =   30
         TabIndex        =   26
         Top             =   1665
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txtcard 
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
         Height          =   300
         Left            =   -68485
         MaxLength       =   20
         TabIndex        =   22
         Top             =   855
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txtmobil 
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
         Height          =   300
         Left            =   -68500
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1245
         Visible         =   0   'False
         Width           =   1485
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
         Height          =   810
         Left            =   -68350
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2700
         Visible         =   0   'False
         Width           =   7530
      End
      Begin VB.TextBox txtpemakai 
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
         Height          =   300
         Left            =   -63025
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1875
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txtnocard 
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
         Height          =   300
         Left            =   -68350
         TabIndex        =   4
         Top             =   1185
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txttol 
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
         Height          =   300
         Left            =   -68350
         MaxLength       =   30
         TabIndex        =   3
         Top             =   2325
         Visible         =   0   'False
         Width           =   7530
      End
      Begin VB.TextBox txtnoknd 
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
         Height          =   300
         Left            =   -68350
         TabIndex        =   1
         Top             =   1575
         Visible         =   0   'False
         Width           =   1485
      End
      Begin XtremeSuiteControls.ComboBox cmbpilih 
         Height          =   315
         Left            =   -68350
         TabIndex        =   2
         Top             =   1950
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851970
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBNumber6Ctl.TDBNumber txtncash 
         Height          =   285
         Left            =   -68350
         TabIndex        =   7
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   503
         Calculator      =   "frme_tol.frx":0000
         Caption         =   "frme_tol.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frme_tol.frx":008C
         Keys            =   "frme_tol.frx":00AA
         Spin            =   "frme_tol.frx":00EC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483628
         BorderStyle     =   1
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin XtremeSuiteControls.PushButton cmdnomor 
         Height          =   300
         Left            =   -69835
         TabIndex        =   8
         Top             =   1170
         Visible         =   0   'False
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "No Kartu"
         BackColor       =   14737632
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
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   -62110
         TabIndex        =   9
         Top             =   1140
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
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
         Format          =   134283265
         CurrentDate     =   41743
      End
      Begin MSComCtl2.DTPicker dtpjam 
         Height          =   330
         Left            =   -62110
         TabIndex        =   31
         Top             =   1515
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
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
         Format          =   134283266
         CurrentDate     =   41743
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   3885
         Left            =   -65575
         TabIndex        =   33
         Top             =   660
         Visible         =   0   'False
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   6853
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483638
         BackColorBkg    =   5395026
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin Crystal.CrystalReport crystal 
         Left            =   120
         Top             =   4140
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSComCtl2.DTPicker dtptgl1 
         Height          =   285
         Left            =   3930
         TabIndex        =   38
         Top             =   795
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
         Format          =   134283267
         CurrentDate     =   37694
      End
      Begin MSComCtl2.DTPicker dtptgl2 
         Height          =   285
         Left            =   3930
         TabIndex        =   39
         Top             =   1125
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
         Format          =   134283267
         CurrentDate     =   37694
      End
      Begin XtremeSuiteControls.PushButton btnkartu 
         Height          =   300
         Left            =   2895
         TabIndex        =   41
         Top             =   1545
         Width           =   945
         _Version        =   851970
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "No Kartu"
         BackColor       =   14737632
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
         Appearance      =   1
      End
      Begin VB.TextBox txtkartu 
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
         Height          =   300
         Left            =   3930
         TabIndex        =   40
         Top             =   1545
         Width           =   2280
      End
      Begin XtremeSuiteControls.PushButton btnview 
         Height          =   405
         Left            =   8265
         TabIndex        =   44
         Top             =   4185
         Width           =   1065
         _Version        =   851970
         _ExtentX        =   1879
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Vew"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox cbdelete 
         Height          =   240
         Left            =   -65215
         TabIndex        =   46
         Top             =   795
         Visible         =   0   'False
         Width           =   825
         _Version        =   851970
         _ExtentX        =   1455
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Delete"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin TDBNumber6Ctl.TDBNumber txtsaldo 
         Height          =   285
         Left            =   -62560
         TabIndex        =   48
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   503
         Calculator      =   "frme_tol.frx":0114
         Caption         =   "frme_tol.frx":0134
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frme_tol.frx":01A0
         Keys            =   "frme_tol.frx":01BE
         Spin            =   "frme_tol.frx":0200
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483628
         BorderStyle     =   1
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64000
         TabIndex        =   49
         Top             =   3960
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lto 
         BackStyle       =   0  'Transparent
         Caption         =   "To date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   43
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Lfrom 
         BackStyle       =   0  'Transparent
         Caption         =   "From date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   42
         Top             =   825
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   975
         Left            =   150
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   2145
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -63865
         TabIndex        =   35
         Top             =   1545
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Transaksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -63865
         TabIndex        =   34
         Top             =   795
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keperluan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70000
         TabIndex        =   28
         Top             =   1980
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pengguna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70105
         TabIndex        =   25
         Top             =   1695
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70120
         TabIndex        =   24
         Top             =   855
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No.Kendaraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70105
         TabIndex        =   23
         Top             =   1260
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PT. SPARTA PRIMA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -63250
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PT. SPARTA PRIMA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -63250
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69790
         TabIndex        =   17
         Top             =   3975
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblterbilang 
         BackStyle       =   0  'Transparent
         Caption         =   "Nol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -68350
         TabIndex        =   16
         Top             =   3645
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terbilang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69790
         TabIndex        =   15
         Top             =   3645
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69760
         TabIndex        =   14
         Top             =   2685
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pengguna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64180
         TabIndex        =   13
         Top             =   1920
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Transaksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64045
         TabIndex        =   12
         Top             =   1185
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Tol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69775
         TabIndex        =   11
         Top             =   2310
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No.Kendaraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69970
         TabIndex        =   10
         Top             =   1605
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Height          =   2925
         Left            =   -69895
         TabIndex        =   29
         Top             =   690
         Visible         =   0   'False
         Width           =   9165
      End
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   405
      Left            =   8280
      TabIndex        =   18
      Top             =   4710
      Width           =   1065
      _Version        =   851970
      _ExtentX        =   1879
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   405
      Left            =   7150
      TabIndex        =   27
      Top             =   4710
      Width           =   1065
      _Version        =   851970
      _ExtentX        =   1879
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Save"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdClear 
      Height          =   405
      Left            =   6015
      TabIndex        =   32
      Top             =   4710
      Width           =   1065
      _Version        =   851970
      _ExtentX        =   1879
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Clear"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmddel 
      Height          =   405
      Left            =   4890
      TabIndex        =   47
      Top             =   4710
      Width           =   1065
      _Version        =   851970
      _ExtentX        =   1879
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Delete"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frme_tol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As New ADODB.Recordset
Private SQL As String

Dim trans_etol As Boolean
Dim kodetol, strtol As String


Private Sub btnkartu_Click()
    carisql1 = "select kdcard,username,noknd from t_etol"
    namatabel = "E-Tol"

    frmsearch.Show vbModal
End Sub

Private Sub btnkartu_GotFocus()
    If hasil = "" Then Exit Sub
    txtkartu = hasil
End Sub

Private Sub btnview_Click()
    If optdate.Value = True Then
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_etolbydate)"
        Crystal.ReportFileName = AppPath & "\reports\gl\ledger\etolbydate.rpt"
        Crystal.ParameterFields(0) = "@kode1;" & Format(dtptgl1, "yyyy/MM/dd") & ";true"
        Crystal.ParameterFields(1) = "@kode2;" & Format(dtptgl2, "yyyy/MM/dd") & ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    ElseIf optcard.Value = True Then
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_etolbycard)"
        Crystal.ReportFileName = AppPath & "\reports\gl\ledger\etolbycard.rpt"
        Crystal.ParameterFields(0) = "@nocard;" & txtkartu & ";true"
        Crystal.ParameterFields(1) = "@kode1;" & Format(dtptgl1, "yyyy/MM/dd") & ";true"
        Crystal.ParameterFields(2) = "@kode2;" & Format(dtptgl2, "yyyy/MM/dd") & ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    End If
End Sub

Private Sub cbAdd_Click()
    
End Sub

Private Sub cbdelete_Click()
    If cbdelete.Value = xtpChecked Then
        txtkdtrans.Enabled = True
        cbedit.Value = xtpUnchecked
        cmdsave.Enabled = False
    Else
        If cbdelete.Value = xtpUnchecked And cbedit.Value = xtpUnchecked Then txtkdtrans.Enabled = False: cmdsave.Enabled = True
    End If
End Sub

Private Sub cbedit_Click()
    If cbedit.Value = xtpChecked Then
        txtkdtrans.Enabled = True
        cbdelete.Value = xtpUnchecked
        cmdsave.Enabled = True
    Else
        If cbdelete.Value = xtpUnchecked And cbedit.Value = xtpUnchecked Then txtkdtrans.Enabled = False
    End If
End Sub

Private Sub cmdclear_Click()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    txtnocard = ""
    txtnoknd = ""
    cmbpilih = ""
    txttol = ""
    txtket = ""
    txtncash = "0.00"
    txtsaldo = "0.00"
    date1 = Date
    dtptgl1 = Date
    dtptgl2 = Date
    dtpjam = Format(Now, "HH:mm:ss")
    txtpemakai = ""
    lblterbilang = "Nol"

    cbedit.Value = xtpUnchecked
    cbdelete.Value = xtpUnchecked
    txtkdtrans.Enabled = False
    txtcard = ""
    txtmobil = ""
    txtuser = ""
    txtkartu = ""
    cmdsave.Enabled = True
    cmddel.Enabled = False
    If trans_etol = True Then
        OBJ.Open dsn
        SQL = "select top 1 kdetol from t_etol_detail where kdetol like '" + strformat + "%' order by kdetol desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            kodetol = Right(RST!kdetol, 3)
        Else
            If IsNull(kdetol) Then kodetol = 0
            kodetol = 0
        End If
        
        kodetol = kodetol + 1
        If Len(kodetol) = 1 Then strtol = strformat & Mid(RST!kdetol, 5, 2) & kodetol
        If Len(kodetol) = 2 Then strtol = strformat & Mid(RST!kdetol, 5, 1) & kodetol
        If Len(kodetol) = 3 Then strtol = strformat & kodetol
        txtkdtrans = strtol
        OBJ.Close
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddel_Click()
    If MsgBox("Are you sure, you want to remove this data", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "Delete From t_etol_detail Where kdetol='" & txtkdtrans & "'"
    OBJ.Execute SQL
    OBJ.Close
    MsgBox "Data is successfuly removed", vbInformation, AppName
    cmddel.Enabled = False
    cmdclear_Click
End Sub

Private Sub cmdnomor_Click()
    If cbdelete.Value = xtpChecked Then
        MsgBox "Your in Delete Mode", vbExclamation, AppName
        Exit Sub
    End If
    carisql1 = "select kdcard,username,noknd from t_etol"
    namatabel = "E-Tol"

    frmsearch.Show vbModal
End Sub

Private Sub cmdnomor_GotFocus()
    If hasil = "" Then Exit Sub
    txtnocard = hasil
    txtpemakai = hasil1
    txtnoknd = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    
    carisaldo
End Sub

Private Sub carisaldo()
    OBJ.Open dsn
    SQL = "Select kdcard,SUM(amount)'saldo' From t_etol_detail "
    SQL = SQL + "Where kdcard= '" & txtnocard & "' GROUP By kdcard"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        txtsaldo = "0.00"
    Else
        txtsaldo = RST!saldo
    End If
    OBJ.Close
End Sub
Private Sub cmdsave_Click()
    OBJ.Open dsn
    
    If trans_etol = True Then
        If cbedit.Value = xtpUnchecked Then
        'add transaksi e-tol
            If txtnocard = "" Or txtnoknd = "" Or txtpemakai = "" Or cmbpilih = "" Then
                OBJ.Close
                MsgBox "Data is not completed", vbCritical, AppName
                Exit Sub
            End If
            'CEK NOMOR TRANSAKSI
            SQL = "Select * From t_etol_detail Where kdetol = '" & txtkdtrans & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                OBJ.Close
                MsgBox "Kode Transaksi is already exist", vbCritical, AppName
                Exit Sub
            End If
            
            SQL = "Select * From t_etol_detail Where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                .AddNew
                !kdetol = txtkdtrans
                !kdcard = txtnocard
                !Status = cmbpilih.text
                !tolname = txttol
                If cmbpilih.text = "TOP-UP" Then
                    !amount = txtncash
                Else
                    !amount = txtncash * -1
                End If
                !tanggal = Format(date1, "yyyy/MM/dd")
                !jam = Format(dtpjam, "HH:mm:ss")
                !keterangan = txtket
                !noknd = txtnoknd
                !flag = "0"
                !pengguna = txtpemakai
                .Update
            End With
            MsgBox "Data berhasil disimpan", vbInformation, AppName
        ElseIf cbedit.Value = xtpChecked Then
        'update transaksi e-tol
            If txtnocard = "" Or txtnoknd = "" Or txtpemakai = "" Or cmbpilih = "" Then
                OBJ.Close
                MsgBox "Data is not completed", vbCritical, AppName
                Exit Sub
            End If
            SQL = "Select * From t_etol_detail Where kdetol='" & txtkdtrans & "'"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                !kdcard = txtnocard
                !Status = cmbpilih.text
                !tolname = txttol
                If cmbpilih.text = "TOP-UP" Then
                    !amount = txtncash
                Else
                    !amount = txtncash * -1
                End If
                !tanggal = Format(date1, "yyyy/MM/dd")
                !jam = Format(dtpjam, "HH:mm:ss")
                !keterangan = txtket
                !noknd = txtnoknd
                !flag = "0"
                !pengguna = txtpemakai
                .Update
            End With
            MsgBox "Data berhasil diupdate", vbInformation, AppName
        End If
    Else
        If txtcard = "" Then Exit Sub
        SQL = "Select * From t_etol Where kdcard= '" & txtcard & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If MsgBox("Apakah Anda ingin merubah data kartu : " & txtcard & ".", vbQuestion + vbYesNo) = vbNo Then OBJ.Close: Exit Sub
        'update master e-tol
            SQL = "Update t_etol set Username='" & txtuser & "',"
            SQL = SQL + " tanggal='" & Format(Date, "yyyy/MM/dd") & "',"
            SQL = SQL + " noknd='" & txtmobil & "'"
            SQL = SQL + " Where kdcard='" & txtcard & "'"
            Set RST = OBJ.Execute(SQL)
            MsgBox "Data berhasil diupdate", vbInformation, AppName
                
            SQL = "Select kdcard,username,noknd from t_etol"
            Set RST = OBJ.Execute(SQL)
            Set grid.DataSource = RST
            setgrid
            
            OBJ.Close
            Exit Sub
        End If
        'add master e-tol
        SQL = "Select * From t_etol Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        With RST
            .AddNew
            !kdcard = txtcard
            !UserName = txtuser
            !tanggal = Format(Date, "yyyy/MM/dd")
            !noknd = txtmobil
            !flag = "0"
            .Update
        End With
        MsgBox "Data berhasil disimpan", vbInformation, AppName
        
        SQL = "Select kdcard,username,noknd from t_etol"
        Set RST = OBJ.Execute(SQL)
        Set grid.DataSource = RST
        setgrid
        
    End If
    OBJ.Close
    cmdclear_Click
End Sub

Private Sub Form_Load()
    On Error GoTo err_handler:
    Dim strformat As String
    strformat = Format(Date, "yymm")
    TabControl1.SelectedItem = "0"
    trans_etol = True
    cmbpilih.AddItem "TOP-UP"
    cmbpilih.AddItem "BAYAR TOL"
    date1 = Date
    dtptgl1 = Date
    dtptgl2 = Date
    dtpjam = Format(Now, "HH:mm:ss")
    
    OBJ.Open dsn
    SQL = "select top 1 kdetol from t_etol_detail where kdetol like '" + strformat + "%' order by kdetol desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        kodetol = Right(RST!kdetol, 3)
    Else
        If IsNull(kdetol) Then kodetol = 0
        kodetol = 0
    End If
    
    kodetol = kodetol + 1
    If Len(kodetol) = 1 Then strtol = strformat & Mid(RST!kdetol, 5, 2) & kodetol
    If Len(kodetol) = 2 Then strtol = strformat & Mid(RST!kdetol, 5, 1) & kodetol
    If Len(kodetol) = 3 Then strtol = strformat & kodetol
    txtkdtrans = strtol
    
    OBJ.Close
    Exit Sub
err_handler:
    strtol = strformat & "001"
    txtkdtrans = strtol
    OBJ.Close
'    MsgBox Err.Description, AppName
End Sub

Private Sub setgrid()
    With grid
        .ColWidth(0) = 1700
        .ColWidth(1) = 1700
        .ColWidth(2) = 1000

        .Cols = 3
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "NO KARTU"
        .TextMatrix(0, 1) = "PENGGUNA KARTU"
        .TextMatrix(0, 2) = "NO.KEND"
        .Refresh
    End With
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0, 1, 2:
            txtcard = grid.TextMatrix(grid.Row, 0)
            txtuser = grid.TextMatrix(grid.Row, 1)
            txtmobil = grid.TextMatrix(grid.Row, 2)
    End Select
End Sub

Private Sub optcard_Click()
    btnkartu.Visible = True
    txtkartu.Visible = True
End Sub

Private Sub optdate_Click()
    btnkartu.Visible = False
    txtkartu.Visible = False
    dtptgl1.Visible = True
    dtptgl2.Visible = True
    Lfrom.Visible = True
    lto.Visible = True
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case TabControl1.SelectedItem
        Case 0: trans_etol = True
        
        Case 1: trans_etol = False
                OBJ.Open dsn
                SQL = "Select kdcard,username,noknd from t_etol"
                Set RST = OBJ.Execute(SQL)
                Set grid.DataSource = RST
                OBJ.Close
                setgrid
        Case 2:
                btnkartu.Visible = False
                txtkartu.Visible = False
    End Select
End Sub

Private Sub txtkdtrans_KeyPress(KeyAscii As Integer)
Dim strjam As String
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Select * From t_etol_detail Where kdetol='" & txtkdtrans & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Data is not found", vbCritical, AppName
            OBJ.Close
            cmdclear_Click
            Exit Sub
        End If
        txtnocard = RST!kdcard
        txtnoknd = RST!noknd
        cmbpilih = RST!Status
        txttol = RST!tolname
        txtket = RST!keterangan
        If RST!Status = "BAYAR TOL" Then
            txtncash = RST!amount * -1
        Else
            txtncash = RST!amount
        End If
        txtpemakai = RST!pengguna
        date1 = RST!tanggal
        strjam = Format(RST!jam, "hh:mm:ss")
        dtpjam = Format(Left(strjam, 8), "hh:mm:ss")
        cmddel.Enabled = True
        OBJ.Close
    End If
End Sub

Private Sub txtncash_Change()
    If txtncash = "" Then Exit Sub
    lblterbilang = ANGKAKEHURUF(Format(txtncash, "general number")) & " Rupiah"
End Sub
