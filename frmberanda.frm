VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmberanda 
   BorderStyle     =   0  'None
   Caption         =   "Beranda"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   6720
      Top             =   120
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   11895
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15375
      _Version        =   851970
      _ExtentX        =   27120
      _ExtentY        =   20981
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowReorder    =   -1  'True
      Appearance      =   3
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "OMSET PENJUALAN"
      Item(0).ControlCount=   23
      Item(0).Control(0)=   "Label1"
      Item(0).Control(1)=   "Label2"
      Item(0).Control(2)=   "Label3"
      Item(0).Control(3)=   "Label4"
      Item(0).Control(4)=   "Label5"
      Item(0).Control(5)=   "lblbrutto"
      Item(0).Control(6)=   "lbldiskon"
      Item(0).Control(7)=   "lbltanggal"
      Item(0).Control(8)=   "lblppn"
      Item(0).Control(9)=   "lblnetto"
      Item(0).Control(10)=   "Label6(0)"
      Item(0).Control(11)=   "Label6(1)"
      Item(0).Control(12)=   "Label6(2)"
      Item(0).Control(13)=   "Label6(3)"
      Item(0).Control(14)=   "Shape1(0)"
      Item(0).Control(15)=   "Shape1(1)"
      Item(0).Control(16)=   "Shape1(2)"
      Item(0).Control(17)=   "Shape1(3)"
      Item(0).Control(18)=   "Shape1(4)"
      Item(0).Control(19)=   "picarea"
      Item(0).Control(20)=   "picsales"
      Item(0).Control(21)=   "Gbjual"
      Item(0).Control(22)=   "Shape1(5)"
      Item(1).Caption =   "PRODUKSI"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "Shape1(7)"
      Item(1).Control(1)=   "Label10"
      Item(2).Caption =   "..."
      Item(2).ControlCount=   0
      Begin XtremeSuiteControls.GroupBox Gbjual 
         Height          =   3015
         Left            =   8400
         TabIndex        =   30
         Top             =   1080
         Width           =   3495
         _Version        =   851970
         _ExtentX        =   6165
         _ExtentY        =   5318
         _StockProps     =   79
         Caption         =   "MONTHLY"
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin MSComCtl2.DTPicker date_tahun 
            Height          =   405
            Left            =   1800
            TabIndex        =   42
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy"
            Format          =   143982595
            CurrentDate     =   37464
         End
         Begin XtremeSuiteControls.PushButton cmdview 
            Height          =   495
            Left            =   1680
            TabIndex        =   38
            Top             =   1920
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "View"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin VB.TextBox txtinv3 
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
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   36
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtinv4 
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
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   35
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(monthly) per Sales"
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
            Left            =   2040
            TabIndex        =   34
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(monthly) per Area Customer"
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
            Left            =   2040
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(monthly) per Customer"
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
            Left            =   2040
            TabIndex        =   32
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Penjualan"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Value           =   -1  'True
            Width           =   2895
         End
         Begin Crystal.CrystalReport crystal 
            Left            =   120
            Top             =   2160
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
         Begin Chameleon.chameleonButton cmdsearch3 
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "From"
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   99
            MICON           =   "frmberanda.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Chameleon.chameleonButton cmdsearch4 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "To"
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   99
            MICON           =   "frmberanda.frx":031A
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker date1 
            Height          =   285
            Left            =   2040
            TabIndex        =   41
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   143982595
            CurrentDate     =   37464
         End
         Begin MSComCtl2.DTPicker date2 
            Height          =   285
            Left            =   2040
            TabIndex        =   43
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   143982595
            CurrentDate     =   37464
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tahun"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   37
            Top             =   1250
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   615
            Index           =   6
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   1080
            Width           =   3135
         End
      End
      Begin VB.PictureBox picarea 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3855
         ScaleWidth      =   6495
         TabIndex        =   27
         Top             =   4440
         Width           =   6495
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
            Height          =   3015
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorBkg    =   -2147483632
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "OMSET AREA BULAN INI"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   5535
         End
      End
      Begin VB.PictureBox picsales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   6720
         ScaleHeight     =   3735
         ScaleWidth      =   6735
         TabIndex        =   24
         Top             =   4440
         Width           =   6735
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   3015
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorBkg    =   -2147483632
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "OMSET SALES BULAN INI"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   120
            Width           =   5415
         End
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TOP PRODUK BY Kg"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69640
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   7
         Left            =   -69880
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Visible         =   0   'False
         Width           =   7935
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   5
         Left            =   8160
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3000
         TabIndex        =   23
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   22
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   21
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblnetto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   19
         Top             =   3480
         Width           =   4215
      End
      Begin VB.Label lblppn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label lbltanggal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   17
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lbldiskon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   16
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label lblbrutto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   15
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "NETTO"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DISKON"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BRUTTO"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DARI TANGGAL"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   0
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   7695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   1
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   7695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   2
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   7695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   3
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   7695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   4
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   7935
      End
   End
   Begin VB.TextBox txtcust1 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox txtcust2 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtsales2 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtsales1 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtarea1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtarea2 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   1935
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   10920
      Width           =   13185
      _Version        =   851970
      _ExtentX        =   23257
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "C L O S E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   143982593
      CurrentDate     =   44056
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   143982593
      CurrentDate     =   44056
   End
End
Attribute VB_Name = "frmberanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim PARAMETER As New ADODB.PARAMETER
Dim SP As New ADODB.Command
Dim SQL As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Sub ambildataomset()
    OBJ.Open dsn
    'Omset Penjualan
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_telegram_omset"
    Set PARAMETER = SP.CreateParameter("@tanggal1", adVarChar, adParamInput, 10, Format(dtp1, "yyyyMMdd"))
    SP.Parameters.Append PARAMETER
    OBJ.Execute "am_telegram_omset '" & Format(dtp1, "yyyyMMdd") & "','" & Format(dtp2, "yyyyMMdd") & "','" & txtcust1 & "','" & txtcust2 & "'"
    Set SP = Nothing
    
    SQL = "Select * From am_telomset where flag='0.00'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lbltanggal = "01 - " & dtp2
        lblbrutto = Format(RST!brutto, "###,##0.00")
        lbldiskon = Format(RST!disc, "###,##0.00")
        lblppn = Format(RST!ppn, "###,##0.00")
        lblnetto = Format(RST!netto, "###,##0.00")
    End If

    SQL = "delete from am_telomset"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Sub ambildatasales()
    OBJ.Open dsn
    'Omset per sales
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_telegram_omsetsales"
    Set PARAMETER = SP.CreateParameter("@tanggal1", adVarChar, adParamInput, 10, Format(dtp1, "yyyyMMdd"))
    SP.Parameters.Append PARAMETER
    OBJ.Execute "am_telegram_omsetsales '" & Format(dtp1, "yyyyMMdd") & "','" & Format(dtp2, "yyyyMMdd") & "','" & txtsales1 & "','" & txtsales2 & "'"
    Set SP = Nothing

    SQL = "Select * From am_telomset Where flag ='1.00'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!bulan
        grid.TextMatrix(grid.Row, 2) = Format(RST!brutto, "###,##0.00")
        grid.TextMatrix(grid.Row, 3) = Format(RST!disc, "###,##0.00")
        grid.TextMatrix(grid.Row, 4) = Format(RST!ppn, "###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!netto, "###,##0.00")
        RST.MoveNext
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
    Loop
    
    SQL = "delete from am_telomset"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Sub ambildataarea()
    OBJ.Open dsn
    
    'Omset per Area
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_telegram_omsetarea"
    Set PARAMETER = SP.CreateParameter("@tanggal1", adVarChar, adParamInput, 10, Format(dtp1, "yyyyMMdd"))
    SP.Parameters.Append PARAMETER
    OBJ.Execute "am_telegram_omsetarea '" & Format(dtp1, "yyyyMMdd") & "','" & Format(dtp2, "yyyyMMdd") & "','" & txtarea1 & "','" & txtarea2 & "'"
    Set SP = Nothing
    
    SQL = "Select * From am_telomset Where flag ='2.00'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 0) = grid1.Row
        grid1.TextMatrix(grid1.Row, 1) = RST!bulan
        grid1.TextMatrix(grid1.Row, 2) = Format(RST!brutto, "###,##0.00")
        grid1.TextMatrix(grid1.Row, 3) = Format(RST!disc, "###,##0.00")
        grid1.TextMatrix(grid1.Row, 4) = Format(RST!ppn, "###,##0.00")
        grid1.TextMatrix(grid1.Row, 5) = Format(RST!netto, "###,##0.00")
        RST.MoveNext
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
    Loop
    
    SQL = "delete from am_telomset"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub caricustMonthly()
    OBJ.Open dsn
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust asc"
    Set RST = OBJ.Execute(SQL)
    txtinv3 = RST!kodecust
    
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust desc"
    Set RST = OBJ.Execute(SQL)
    txtinv4 = RST!kodecust
    OBJ.Close
End Sub

Private Sub cmdview_Click()
    If Option1.Value = True Then str1 = "mall"
    If Option2.Value = True Then str1 = "mcust"
    If Option3.Value = True Then str1 = "marea"
    If Option4.Value = True Then str1 = "msales"
    status = "P"
        
    crystal.reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    'crystal.Window = True
    
    crystal.Connect = dsnreport
    If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Then
        crystal.DataFiles(0) = "Proc(am_monthly)"
    ElseIf Option4.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualbulansale)"
    Else
    
    End If
    
    If Option1.Value = True Then
        crystal.ReportFileName = App.Path & "\reports\sale\inv\monthlypl_all.rpt"
    ElseIf Option2.Value = True Then
        crystal.ReportFileName = App.Path & "\reports\sale\inv\monthlypl.rpt"
    ElseIf Option3.Value = True Then
        crystal.ReportFileName = App.Path & "\reports\sale\inv\monthlypl_area.rpt"
    ElseIf Option4.Value = True Then
        crystal.ReportFileName = App.Path & "\reports\sale\inv\monthlypl_sales.rpt"
    End If
    crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyy-MM-dd") & ";true"
    crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
    crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
    
    If Option4.Value = True Then
        crystal.ParameterFields(4) = "@namauser ;" + nmuser + ";true"
        crystal.ParameterFields(5) = "@PL ;" + status + ";true"
    Else
        crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
        crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
        crystal.ParameterFields(6) = "@PL ;" + status + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    Dim h As Integer
    Dim t As Date
    Dim thn, bln, bln2, tgl, tgl2 As String
    
    h = Day(Date)
    h = h - 1
    t = DateAdd("d", -h, Date)
    dtp1 = t
    dtp2 = Date
    date_tahun = Date
    
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
    
    setgrid
    setgrid2
    caricust
    ambildataomset
    ambildatasales
    ambildataarea
    
    caricustMonthly
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
End Sub

Private Sub date_tahun_Change()
    Dim thn, bln, bln2, tgl, tgl2 As String
    thn = Year(date_tahun)
    bln = "01"
    bln2 = "12"
    tgl = "01"
    tgl2 = "31"
    date1 = thn & "-" & bln & "-" & tgl
    date2 = thn & "-" & bln2 & "-" & tgl2
End Sub

Private Sub caricust()
    On Error GoTo Err_handler:
    OBJ.Open dsn
    'customer
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust asc"
    Set RST = OBJ.Execute(SQL)
    txtcust1 = RST!kodecust
    
    SQL = "Select top 1 KodeCust from am_customer order by KodeCust desc"
    Set RST = OBJ.Execute(SQL)
    txtcust2 = RST!kodecust
    
    'sales
    SQL = "Select top 1 kodesales from am_salesman where idupdate<>'0' order by kodesales asc"
    Set RST = OBJ.Execute(SQL)
    txtsales1 = RST!kodesales
    
    SQL = "Select top 1 KodeSales from am_salesman where idupdate<>'0' order by KodeSales desc"
    Set RST = OBJ.Execute(SQL)
    txtsales2 = RST!kodesales
    
    'area
    SQL = "Select top 1 kode from am_area order by kode asc"
    Set RST = OBJ.Execute(SQL)
    txtarea1 = RST!kode
    
    SQL = "Select top 1 Kode from am_area order by Kode desc"
    Set RST = OBJ.Execute(SQL)
    txtarea2 = RST!kode
    
    SQL = "Select bot from am_telid"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Bot Telegram tidak ditemukan", vbCritical, "WARNING"
        End
    End If
    Bot = RST!Bot
    OBJ.Close
    Exit Sub
Err_handler:
    MsgBox "Tidak Dapat terkoneksi, periksa koneksi internet Anda", vbCritical, "ERROR CONECTION"
    End
End Sub

Private Sub Form_Resize()
    TabControl1.Move 0, 0, Me.Width - 50, Me.Height - cmdclose.Height
    cmdclose.Move 0, TabControl1.Height, TabControl1.Width
    picarea.Width = TabControl1.Width - 240
    picarea.Height = (TabControl1.Height - picarea.Top - 50) / 2
    picsales.Move 120, picarea.Height + picarea.Top, TabControl1.Width - 240, TabControl1.Height - (picarea.Height + picarea.Top + 150)
    
    Gbjual.Move 8400, 1080
    
    grid1.Width = picarea.Width - 240
    grid1.Height = picarea.Height - 840
    grid.Width = picsales.Width - 240
    grid.Height = picsales.Height - 840
    
    grid1.ColWidth(0) = grid1.Width * 0.05
    grid1.ColWidth(1) = grid1.Width * 0.18
    grid1.ColWidth(2) = grid1.Width * 0.18
    grid1.ColWidth(3) = grid1.Width * 0.18
    grid1.ColWidth(4) = grid1.Width * 0.18
    grid1.ColWidth(5) = grid1.Width * 0.18
    
    grid.ColWidth(0) = grid.Width * 0.05
    grid.ColWidth(1) = grid.Width * 0.18
    grid.ColWidth(2) = grid.Width * 0.18
    grid.ColWidth(3) = grid.Width * 0.18
    grid.ColWidth(4) = grid.Width * 0.18
    grid.ColWidth(5) = grid.Width * 0.18
End Sub
Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.Col = 5
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.Col = 5
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
End Sub

Private Sub setgrid()
    grid.TextMatrix(0, 1) = "Sales"
    grid.TextMatrix(0, 2) = "Brutto"
    grid.TextMatrix(0, 3) = "Diskon"
    grid.TextMatrix(0, 4) = "PPN"
    grid.TextMatrix(0, 5) = "Netto"
    
    
    grid.ColAlignment(0) = flexAlignCenterCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignRightCenter
    grid.ColAlignment(3) = flexAlignRightCenter
    grid.ColAlignment(4) = flexAlignRightCenter
    grid.ColAlignment(5) = flexAlignRightCenter
    grid.Rows = 2
    
End Sub

Private Sub setgrid2()
    grid1.TextMatrix(0, 1) = "Area"
    grid1.TextMatrix(0, 2) = "Brutto"
    grid1.TextMatrix(0, 3) = "Diskon"
    grid1.TextMatrix(0, 4) = "PPN"
    grid1.TextMatrix(0, 5) = "Netto"
    
    
    
    grid1.ColAlignment(0) = flexAlignCenterCenter
    grid1.ColAlignment(1) = flexAlignLeftCenter
    grid1.ColAlignment(2) = flexAlignRightCenter
    grid1.ColAlignment(3) = flexAlignRightCenter
    grid1.ColAlignment(4) = flexAlignRightCenter
    grid1.ColAlignment(5) = flexAlignRightCenter
    grid1.Rows = 2
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As control
  
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSHFlexGrid Then
          If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        End If
    Next ctl
End Sub

Private Sub Timer1_Timer() 'refresh tiap 10 detik
    Dim h As Integer
    Dim t As Date
    h = Day(Date)
    h = h - 1
    t = DateAdd("d", -h, Date)
    dtp1 = t
    dtp2 = Date
    
    setgrid
    setgrid2
    caricust
    ambildataomset
    ambildatasales
    ambildataarea
End Sub
