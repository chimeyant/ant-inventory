VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAdjbase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Stock Base WIP"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnolot 
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
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtkode 
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
      Left            =   7200
      MaxLength       =   15
      TabIndex        =   15
      Top             =   600
      Width           =   1335
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1680
      Width           =   7695
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   5160
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
      MICON           =   "frmAdjbase.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpmut 
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Top             =   240
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
      Format          =   142278659
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1275
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmAdjbase.frx":031A
      Caption         =   "frmAdjbase.frx":033A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdjbase.frx":03A6
      Keys            =   "frmAdjbase.frx":03C4
      Spin            =   "frmAdjbase.frx":0406
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   5160
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
      MICON           =   "frmAdjbase.frx":042E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483631
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
   Begin Chameleon.chameleonButton cmdkode 
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Kode/Item"
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
      MICON           =   "frmAdjbase.frx":0748
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   8895
      _Version        =   851970
      _ExtentX        =   15690
      _ExtentY        =   450
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   5160
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
      MICON           =   "frmAdjbase.frx":0A62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txthppkg 
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmAdjbase.frx":0D7C
      Caption         =   "frmAdjbase.frx":0D9C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdjbase.frx":0E08
      Keys            =   "frmAdjbase.frx":0E26
      Spin            =   "frmAdjbase.frx":0E68
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txttotalhpp 
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmAdjbase.frx":0E90
      Caption         =   "frmAdjbase.frx":0EB0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdjbase.frx":0F1C
      Keys            =   "frmAdjbase.frx":0F3A
      Spin            =   "frmAdjbase.frx":0F7C
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtqtysisa 
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmAdjbase.frx":0FA4
      Caption         =   "frmAdjbase.frx":0FC4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdjbase.frx":1030
      Keys            =   "frmAdjbase.frx":104E
      Spin            =   "frmAdjbase.frx":1090
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode/Item"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "No. Lot"
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
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblrow 
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
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   4335
   End
   Begin VB.Label lblkdsat 
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblbrg 
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
      Left            =   7200
      TabIndex        =   17
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Qty"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1275
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Mutasi"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "No Mutasi"
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
      Left            =   120
      TabIndex        =   13
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Desc/Reference"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label lbltype 
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
      Left            =   2160
      TabIndex        =   11
      Top             =   150
      Width           =   2895
   End
   Begin MSForms.ComboBox cmbtype 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   735
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1296;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      Caption         =   "Type Mutasi"
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
      Left            =   120
      TabIndex        =   9
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdjbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim jumlah As Integer
Dim TotNet As Double

Private Sub cmbtype_Change()
    txtnobukti = ""
    Date1 = Date
    txtnobukti.SetFocus
    
    If cmbtype = "01" Then lbltype = "Pinjaman (In)"
    If cmbtype = "02" Then lbltype = "Pinjaman (Out)"
    If cmbtype = "03" Then lbltype = "KeBarang Jadi (Out)"
End Sub

Private Sub cmdclear_Click()
    cmbtype = ""
    lbltype = ""
    txtnobukti = ""
    txtkode = ""
    txtnolot = ""
    lblbrg = ""
    txtnilai = "0.00"
    txthppkg = "0.00"
    txttotalhpp = "0.00"
    txtqtysisa = "0.00"
    txtket = ""
    dtpmut = Date
    lblkdsat = ""
    Call openallbase
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdkode_Click()
    namatabel = "item base"

    carisql1 = "Select distinct a.kodebahan,b.NamaBarang,b.KodeSatuan From am_stoklot a"
    carisql1 = carisql1 + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang "
    frmsearch.Show vbModal
End Sub

Private Sub cmdkode_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    lblbrg = hasil1
    lblkdsat = hasil2
    hapusgrid
    OBJ.Open dsn
    SQL = "SELECT COUNT(*)'jml' FROM (SELECT nolot,SUM(qtybahan)'qty'"
    SQL = SQL + " From am_stoklot Where kodebahan='" & hasil & "' and flag='0' group by nolot) AS subquery"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "select a.nolot,a.kodebahan,b.NamaBarang,SUM(a.qtybahan)'qty',SUM(a.hpp)'hpp' from am_stoklot a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang"
    SQL = SQL + " Where a.kodebahan='" & hasil & "' and a.flag='0' group by a.nolot,a.kodebahan,b.NamaBarang"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!kodebahan
        grid.TextMatrix(grid.Row, 3) = RST!namabarang
        grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!hpp, "##,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        RST.MoveNext
    Loop
    OBJ.Close
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    Pg.Value = 0
    Call Totalkg
    lblrow = jumlah & " Lot available on WIP : " & Format(TotNet, "##,###,##0.00") & " Kg"
End Sub

Private Sub opendata()
    hapusgrid
    OBJ.Open dsn
    SQL = "SELECT COUNT(*)'jml' FROM (SELECT nolot,SUM(qtybahan)'qty'"
    SQL = SQL + " From am_stoklot Where kodebahan='" & txtkode & "' and flag='0' group by nolot) AS subquery"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "select a.nolot,a.kodebahan,b.NamaBarang,SUM(a.qtybahan)'qty',SUM(a.hpp)'hpp' from am_stoklot a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang"
    SQL = SQL + " Where a.kodebahan='" & txtkode & "' and a.flag='0' group by a.nolot,a.kodebahan,b.NamaBarang"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!kodebahan
        grid.TextMatrix(grid.Row, 3) = RST!namabarang
        grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!hpp, "##,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        RST.MoveNext
    Loop


    'cari sisa qty pada lot
    SQL = "select nolot,SUM(qtybahan)'qty',SUM(hpp)'hpp' from am_stoklot"
    SQL = SQL + " Where flag='0' and nolot='" & txtnolot & "' group by nolot"
    Set RST = OBJ.Execute(SQL)
    
    txtqtysisa = RST!qty
    
    OBJ.Close
    Pg.Value = 0
    Call Totalkg
    lblrow = jumlah & " Lot available on WIP : " & Format(TotNet, "##,###,##0.00") & " Kg"
End Sub
Private Sub openallbase()
    hapusgrid
    OBJ.Open dsn
    
    SQL = "SELECT COUNT(*)'jml' FROM (SELECT nolot,SUM(qtybahan)'qty'"
    SQL = SQL + " From am_stoklot Where flag='0' group by nolot) AS subquery"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "select a.nolot,a.kodebahan,b.NamaBarang,SUM(a.qtybahan)'qty',SUM(a.hpp)'hpp' from am_stoklot a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebahan = b.KodeBarang"
    SQL = SQL + " Where a.flag='0' group by a.nolot,a.kodebahan,b.NamaBarang"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!kodebahan
        grid.TextMatrix(grid.Row, 3) = RST!namabarang
        grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = Format(RST!hpp, "##,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    
    Call Totalkg
    lblrow = jumlah & " Lot available on WIP : " & Format(TotNet, "##,###,##0.00") & " Kg"
End Sub
Private Sub cmdsave_Click()
    Dim nomutlot As String
    If cmbtype = "" Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    If txtkode = "" Then Exit Sub
    If txtnolot = "" Then
        MsgBox "No Lot Belum Diisi.", vbCritical, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select nolot,SUM(qtybahan)'qty' from am_stoklot Where nolot='" & txtnolot & "'"
    SQL = SQL + " and kodebahan='" & txtkode & "' group by nolot"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF And RST!qty = "0.00" Then
        MsgBox "Lot telah habis terpakai", vbExclamation, AppName
        'update flag 0=1
        
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    'Simpan ke am_muthdr , am_mutlin , am_stoklot
    nomutlot = getnomut
    
    OBJ1.Open dsn
    SQL1 = "Select * From am_muthdr Where nomut='" & txtnobukti & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    
    OBJ.Open dsn
    If Not RST1.EOF Then GoTo nextline:
    
    SQL = "Select * From am_muthdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    RST.AddNew
    RST!nomut = txtnobukti
    RST!tglmut = Format(dtpmut, "yyyy/MM/dd")
    RST!Type = cmbtype
    RST!keterangan = txtket
    RST.Update
    
nextline:
    OBJ1.Close
    SQL = "Select * From am_mutlin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    RST.AddNew
    RST!nomut = txtnobukti
    RST!Type = cmbtype
    RST!kodebarang = txtkode
    RST!qty = txtnilai
    RST!kodesatuan = lblkdsat
    RST!lineitem = "1"
    RST.Update
    
    SQL = "Select * From am_stoklot Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    RST.AddNew
    RST!lotstok = nomutlot
    RST!lotsop = txtnobukti
    RST!nolot = txtnolot
    RST!kodebahan = txtkode
    If cmbtype = "01" Then
        RST!qtybahan = txtnilai
        RST!hpp = "0.00"
    ElseIf cmbtype = "02" Then
        RST!qtybahan = txtnilai * -1
        RST!hpp = txttotalhpp * -1
    End If
    RST!kodesatuan = lblkdsat
    RST!flag = "0"
    RST.Update
    
    If cmbtype = "02" And txtqtysisa - txtnilai <= "0.00" Then
    'Jika qty pada lotbase sudah habis update flag=1
        SQL = "Update am_stoklot set flag='1' Where nolot='" & txtnolot & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    If MsgBox("Lanjutkan dengan nomor mutasi yang sama", vbQuestion + vbYesNo, "Konfirmasi nomor mutasi") = vbYes Then
        txtkode = ""
        txtnolot = ""
        lblbrg = ""
        txtnilai = "0.00"
        txthppkg = "0.00"
        txttotalhpp = "0.00"
        txtqtysisa = "0.00"
        lblkdsat = ""
        Call openallbase
    Else
        cmdclear_Click
    End If
End Sub

Private Sub Form_Load()
    grid.Cols = 6
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "No Lot"
    grid.TextMatrix(0, 2) = "Kode"
    grid.TextMatrix(0, 3) = "Item"
    grid.TextMatrix(0, 4) = "Qty"
    grid.TextMatrix(0, 5) = "Hpp"
    grid.ColWidth(0) = 450
    grid.ColWidth(1) = 1800
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 2000
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 0
    grid.ColAlignment(0) = flexAlignLeftCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignmentFixed(4) = flexAlignCenterCenter
    grid.ColAlignmentFixed(5) = flexAlignCenterCenter

    cmbtype.Clear
    cmbtype.ColumnCount = 2
    cmbtype.ListWidth = "6 cm"
    cmbtype.ColumnWidths = "2 cm; 4 cm"
    
    cmbtype.AddItem "01"
    cmbtype.AddItem "02"
    cmbtype.List(0, 1) = "Pinjaman (In)"
    cmbtype.List(1, 1) = "Pinjaman (Out)"
    
    dtpmut = Date
    
    Call openallbase
End Sub
Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Function getnomut() As String    '2016060001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yyyymm")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(lotstok)as kr from am_stoklot where lotstok like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getnomut = strnumber + "0001"
    Else
        kode = CLng(Mid(RST!kr, 7, 4)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "000" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 4) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getnomut = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getnomut = strnumber + "0001"
End Function

Private Sub cmbtype_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtnobukti.SetFocus
    KeyAscii = 0
End Sub
Private Sub Totalkg()
On Error Resume Next
    grid.Row = 1
    tkg = 0
    Do While True
        DoEvents
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        tkg = CDbl(Format(grid.TextMatrix(grid.Row, 4), "general number") + CDbl(tkg))
        grid.Row = grid.Row + 1
    Loop
    tkg = Format(tkg, "##,###,##0.00")
    TotNet = Format(tkg, "##,###,##0.00")
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 1:
            'txtnolot = grid.TextMatrix(grid.Row, 1)
    End Select
End Sub

Private Sub txtnilai_Change()
txttotalhpp = txthppkg * txtnilai
End Sub

Private Sub txtnolot_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnolot_LostFocus()
    'cek nolot ada atau tidak
    OBJ.Open dsn
    SQL = "Select * From am_stoklot where nolot = '" & txtnolot & "' and lotsop is null"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Lot tidak ditemukan", vbCritical, AppName
        txtnolot = ""
        OBJ.Close
        Exit Sub
    ElseIf RST!flag = "1" Then
        MsgBox "Lot sudah Close", vbCritical, AppName
        OBJ.Close
        Exit Sub
    End If
    txtkode = RST!kodebahan
    lblkdsat = RST!kodesatuan
    txthppkg = RST!hpp / RST!qtybahan
    
    SQL = "Select * From am_apitemmst Where kodebarang='" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    lblbrg = RST!namabarang
    
    OBJ.Close
    
    Call opendata
End Sub
