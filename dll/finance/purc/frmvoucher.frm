VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmvoucher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create / Change Voucher"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8040
      Top             =   105
   End
   Begin VB.CheckBox chknonpajak 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Non Pajak"
      Height          =   300
      Left            =   6000
      TabIndex        =   29
      Top             =   150
      Visible         =   0   'False
      Width           =   285
   End
   Begin XtremeSuiteControls.PushButton cmdnovoucher 
      Height          =   315
      Left            =   180
      TabIndex        =   26
      Top             =   240
      Width           =   1110
      _Version        =   851970
      _ExtentX        =   1958
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "No Voucher :"
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
      EnableMarkup    =   -1  'True
   End
   Begin VB.TextBox txttext 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6450
      MaxLength       =   60
      TabIndex        =   16
      Top             =   765
      Visible         =   0   'False
      Width           =   2070
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   6735
      TabIndex        =   15
      Top             =   420
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmvoucher.frx":0000
      Caption         =   "frmvoucher.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmvoucher.frx":008C
      Keys            =   "frmvoucher.frx":00AA
      Spin            =   "frmvoucher.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(-##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,###,##0.00;(-##,###,###,##0.00)"
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
      ValueVT         =   1179649
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
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
      Left            =   4815
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   480
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
      Left            =   5295
      Picture         =   "frmvoucher.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   480
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
      Left            =   5055
      Picture         =   "frmvoucher.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   2895
      TabIndex        =   14
      Top             =   615
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   393216
      Format          =   120717313
      CurrentDate     =   41396
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   420
      Left            =   7905
      TabIndex        =   12
      Top             =   6165
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   -2147483633
      Appearance      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2625
      Left            =   105
      TabIndex        =   11
      Top             =   3495
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   4630
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   -2147483636
      BackColorBkg    =   16777215
      GridColorFixed  =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtalamat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   1965
      Width           =   7230
   End
   Begin VB.TextBox txtnpwp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1620
      Width           =   7230
   End
   Begin VB.TextBox txtkepada 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1275
      Width           =   7230
   End
   Begin VB.TextBox txtnota 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   930
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1335
      TabIndex        =   2
      Top             =   615
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      _Version        =   393216
      Format          =   120717313
      CurrentDate     =   41396
   End
   Begin VB.TextBox txtnovoucher 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   420
      Left            =   5865
      TabIndex        =   13
      Top             =   6165
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Save"
      BackColor       =   -2147483633
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   420
      Left            =   4845
      TabIndex        =   20
      Top             =   6165
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Clear"
      BackColor       =   -2147483633
      Appearance      =   2
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs 
      Height          =   315
      Left            =   1320
      TabIndex        =   23
      Top             =   2715
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   556
      Calculator      =   "frmvoucher.frx":0744
      Caption         =   "frmvoucher.frx":0764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmvoucher.frx":07D0
      Keys            =   "frmvoucher.frx":07EE
      Spin            =   "frmvoucher.frx":0830
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
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
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   315
      Left            =   3285
      TabIndex        =   25
      Top             =   2715
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   556
      Calculator      =   "frmvoucher.frx":0858
      Caption         =   "frmvoucher.frx":0878
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmvoucher.frx":08E4
      Keys            =   "frmvoucher.frx":0902
      Spin            =   "frmvoucher.frx":0944
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
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
      ValueVT         =   1806434309
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdkurs 
      Height          =   330
      Left            =   270
      TabIndex        =   27
      Top             =   2325
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Currency :"
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
      EnableMarkup    =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdhapus 
      Height          =   420
      Left            =   6885
      TabIndex        =   28
      Top             =   6165
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Delete"
      BackColor       =   -2147483633
      Appearance      =   5
   End
   Begin VB.TextBox txtkurs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   21
      Top             =   2325
      Width           =   870
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs2 
      Height          =   315
      Left            =   1320
      TabIndex        =   30
      Top             =   3075
      Visible         =   0   'False
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   556
      Calculator      =   "frmvoucher.frx":096C
      Caption         =   "frmvoucher.frx":098C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmvoucher.frx":09F8
      Keys            =   "frmvoucher.frx":0A16
      Spin            =   "frmvoucher.frx":0A58
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
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
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1806434309
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdkpd 
      Height          =   270
      Left            =   330
      TabIndex        =   33
      Top             =   1290
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   "Kepada :"
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
      EnableMarkup    =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type Voucher :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4635
      TabIndex        =   32
      Top             =   180
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Kurs Beli :"
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
      Left            =   150
      TabIndex        =   31
      Top             =   3090
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PPn"
      Height          =   180
      Left            =   2835
      TabIndex        =   24
      Top             =   2760
      Width           =   450
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Kurs  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   22
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   9
      Top             =   2025
      Width           =   1125
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N.P.W.P :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   7
      Top             =   1650
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kepada :"
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
      Left            =   150
      TabIndex        =   5
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No Nota :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   975
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   645
      Width           =   1125
   End
End
Attribute VB_Name = "frmvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private posrow As Integer
Private poscol As Integer

Private edit As Boolean

Private Sub chknonpajak_Click()
    Timer1.Enabled = True
End Sub

Private Sub cmdclear_Click()
    txtnovoucher = ""
    date1 = Date
    date2 = Date
    txtkepada = ""
    txtnpwp = ""
    txtalamat = ""
    txtnota = ""
    txtkurs = ""
    txtnilaikurs.text = "1.00"
    txtppn.Value = 0
    hapusgrid
    edit = False
    txtnovoucher = GetNoVoucher
    txtnovoucher.SetFocus
End Sub

Private Sub printdoc()
    Dim nilai_rupiah As Double
    Dim nilai_pnn As Double
    Dim nilai_hutang As Double
    
    SQL = "select sum(jumlah) as jml from am_voucherin where novoucher='" + txtnovoucher + "'"

    
    Set RST = OBJ.Execute(SQL)
    
    nilai_rupiah = RST!jml * txtnilaikurs
    If txtppn.Value > 0 Then
        nilai_pnn = SpyRoundUp(nilai_rupiah * (txtppn.Value / 100))
    End If
    
    nilai_hutang = nilai_rupiah + nilai_pnn

    
    SQL = "select * from am_voucherin where novoucher='" + txtnovoucher + "'"
    
    With rptprintvoucher_1
        .lblsupp = txtkepada
        .lblnpwp = txtnpwp
        .lblalamat = txtalamat
        .lblnovoucher = ": " + txtnovoucher
        .lbltanggal = ": " + Format(date1, "dd/MM/yyyy")
        .lblkurs = txtkurs
        .lblnilaikurs = txtnilaikurs.text
        .lblppn = Format(nilai_pnn, "###,###,##0.00")
        .lbljumlah = Format(nilai_rupiah, "###,###,##0.00")
        .lblhutang = Format(nilai_hutang, "###,###,##0.00")
        
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdhapus_Click()
    If txtnovoucher = "" Then Exit Sub
    If MsgBox("Are you sure to delete ...?", vbQuestion + vbYesNo, AppName) = vbYes Then
         
        OBJ.Open dsn
    
        SQL = "delete  from am_voucherhdr where novoucher='" + txtnovoucher + "'"
        OBJ.Execute SQL
    
        SQL = "delete from am_voucherin where novoucher='" + txtnovoucher + "'"
        OBJ.Execute SQL
        OBJ.Close
        MsgBox "Delete Succesed...", vbInformation, AppName
        
        cmdclear_Click
    End If
End Sub

Private Sub cmdkpd_Click()
    carisql1 = "Select NamaSupp,AlamatSupp1,AlamatSupp2,kodesupp From am_supplier"
    namatabel = "Supplier"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkpd_GotFocus()
    If hasil = "" Then Exit Sub
    txtkepada = hasil
    txtalamat = hasil1
    hasil = ""
    hasil1 = ""
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub cmdkurs_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkurs_GotFocus()
    If hasil = "" Then Exit Sub
    txtkurs = hasil
    hasil = ""
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err_msg
    Dim ppn As Double
    
    If txtnovoucher = "" Then
        MsgBox "Data tidak lengkap", vbCritical, AppName
        Exit Sub
    End If
    If txtkurs = "" Then
        MsgBox "Kolom Currency kosong", vbCritical, AppName
        Exit Sub
    End If
    If edit = False Then
        txtnovoucher = GetNoVoucher
    End If
    If OBJ.State = 1 Then OBJ.Close
    OBJ.Open dsn
    
    
    SQL = "select * from am_voucherhdr where novoucher='" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Data Telah Ada Apakah anda akan melakukan update data tersebut...?", vbQuestion + vbYesNo, AppName) = vbNo Then
            OBJ.Close
            Exit Sub
        End If
    End If

    SQL = "delete  from am_voucherhdr where novoucher='" + txtnovoucher + "'"
    OBJ.Execute SQL
    
    SQL = "delete from am_voucherin where novoucher='" + txtnovoucher + "'"
    OBJ.Execute SQL
    
    SQL = "Insert into am_voucherhdr("
    SQL = SQL + "novoucher,tgl,kepada,npwp,alamat,kdkurs,nilai,ppn,username,ispajak) VALUES('"
    SQL = SQL + txtnovoucher + "',"
    SQL = SQL + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'"
    SQL = SQL + txtkepada + "','"
    SQL = SQL + txtnpwp + "','"
    SQL = SQL + txtalamat + "','"
    SQL = SQL + txtkurs + "',"
    SQL = SQL + "convert(money,'" + Format(txtnilaikurs, "general number") + "'),"
    SQL = SQL + "convert(money,'" + Format(txtppn, "general number") + "'),'"
    SQL = SQL + nmuser + "','"
    If chknonpajak.Value = 0 Then
        SQL = SQL + "1" + "')"
    End If
    If chknonpajak.Value = 1 Then
        SQL = SQL + "0" + "')"
    End If
    
    OBJ.Execute SQL
    
    
    Dim jmlnilaikurs As Double
    Dim nilaikurs As Double
    Dim nilaibaris As Double
    
    grid.Row = 1
    Do While True
        With grid
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "insert into am_voucherin ("
            SQL = SQL + "novoucher,nonota,tgl,keterangan,perkiraan,jumlah) VALUES('"
            SQL = SQL + txtnovoucher + "','"
            SQL = SQL + grid.TextMatrix(.Row, 2) + "',"
            'SQL = SQL + "convert(datetime,'" + Format(date2, "MM/dd/yyyy") + "'),'"
            SQL = SQL + "convert(datetime, '" + Format(grid.TextMatrix(.Row, 1), "MM/dd/yyyy") + "'),'"
            SQL = SQL + grid.TextMatrix(.Row, 3) + "','"
            SQL = SQL + grid.TextMatrix(.Row, 4) + "',"
            nilaibaris = Format(grid.TextMatrix(.Row, 5), "general number")
            nilaikurs = nilaibaris
            SQL = SQL + "convert(money,'" + Str(nilaikurs) + "'))"
            OBJ.Execute (SQL)
            .Row = .Row + 1
        End With
    Loop
    
    printdoc
    OBJ.Close
    If edit = False Then
        MsgBox "Data is Saved...", vbInformation, AppName
    Else
        MsgBox "Update is Saved...", vbInformation, AppName
    End If
    cmdclear_Click
    Exit Sub
err_msg:
     MsgBox Err.Description
     OBJ.Close
End Sub

Private Sub date2_Change()
    grid.TextMatrix(posrow, poscol) = date2.Value
    grid.Col = 0
    Set grid.CellPicture = uncheck
    grid.Rows = grid.Rows + 1
    grid.SetFocus
End Sub

Private Sub date2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       grid.TextMatrix(posrow, poscol) = date2.Value
        grid.Col = 0
        Set grid.CellPicture = uncheck
        grid.Rows = grid.Rows + 1
        'grid.Rows = grid.Row + 2
        If grid.Row = grid.Rows - 1 Then Exit Sub
        grid.Row = grid.Row + 1
        grid.SetFocus
    End If
End Sub

Private Sub date2_LostFocus()
    date2.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 113 And chknonpajak.Value = 0 Then
        'chknonpajak.Value = 1
        'MsgBox "ok"
    'ElseIf KeyCode = 113 And chknonpajak.Value = 1 Then
        'chknonpajak.Value = 0
        'MsgBox "yes"
    'End If
    If KeyCode = 113 Then
        If chknonpajak.Value = Unchecked Then
            chknonpajak.Value = Checked
        Else
            chknonpajak.Value = Unchecked
        End If
    End If
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    With grid
        .Cols = 6
        .TextMatrix(0, 1) = "Tanggal Nota"
        .TextMatrix(0, 2) = "No Nota"
        .TextMatrix(0, 3) = "Keterangan"
        .TextMatrix(0, 4) = "Perkiraan"
        .TextMatrix(0, 5) = "Jumlah"
    End With
    
    setgrid
    Timer1.Enabled = True
    date1 = Date
End Sub

Private Sub setgrid()
    With grid
        .RowHeightMin = 350
        .ColWidth(0) = 500
        .ColWidth(1) = 1300
        .ColWidth(2) = 1500
        .ColWidth(3) = 2500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1500
    End With
End Sub

Private Sub grid_Click()
    On Error Resume Next
    If grid.MouseRow = 0 Then Exit Sub
    
    posrow = grid.Row
    poscol = grid.Col
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
            With date2
                .Width = grid.ColWidth(grid.Col) - 40
                If grid.TextMatrix(grid.Row, 1) = "" Then
                    .Value = Date
                Else
                    .Value = grid.TextMatrix(grid.Row, 1)
                End If
                .Left = grid.Left + grid.CellLeft
                .Top = grid.Top + grid.CellTop + 20
                .Height = grid.CellHeight - 40
                .Visible = True
                .SetFocus
            End With
        Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 2, 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            With txttext
                .Width = grid.ColWidth(grid.Col) - 40
                .text = grid.TextMatrix(grid.Row, grid.Col)
                .Left = grid.Left + grid.CellLeft
                .Top = grid.Top + grid.CellTop + 20
                .Height = grid.CellHeight - 40
                .Visible = True
                .SetFocus
            End With
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            carisql1 = "select noac, nmac from gl_masterac"
            namatabel = "Master Account"
            frmsearch.Show vbModal
    End Select
    
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
        
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    
    grid.Rows = 2
    setgrid
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture
    grid.Col = 1
End Sub

Private Sub grid_EnterCell()
    On Error Resume Next
    poscol = grid.Col
    posrow = grid.Row
         
     Select Case grid.Col
        
     Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai.Value = Format(grid.TextMatrix(grid.Row, grid.Col), "general number")
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 2, 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            With txttext
                .Width = grid.ColWidth(grid.Col) - 40
                .text = grid.TextMatrix(grid.Row, grid.Col)
                .Left = grid.Left + grid.CellLeft
                .Top = grid.Top + grid.CellTop + 20
                .Height = grid.CellHeight - 40
                .Visible = True
                .SetFocus
            End With
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    grid.TextMatrix(posrow, 4) = hasil
    carisql1 = ""
    hasil = ""
    hasil1 = ""
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If edit = False Then
    txtnovoucher = GetNoVoucher
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(posrow, poscol) = Format(txtnilai, "###,###,##0.00")
        grid.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnovoucher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtnovoucher = "" Then Exit Sub
        OpenVoucher
    End If
End Sub

Private Sub OpenVoucher()
    On Error GoTo Err_handler
    SQL = "select * from am_voucherhdr where novoucher='" + txtnovoucher + "'"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    
    edit = True
    
    date1.Value = RST!tgl
    txtkepada = RST!kepada
    txtnpwp = RST!npwp
    txtalamat = RST!alamat
    If RST!ispajak = "1" Then chknonpajak.Value = 0
    If RST!ispajak = "0" Then chknonpajak.Value = 1
    If Not IsNull(RST!kdkurs) Then
        txtkurs = RST!kdkurs
    End If
    If Not IsNull(RST!nilai) Then
        txtnilaikurs = RST!nilai
    End If
    If Not IsNull(RST!ppn) Then
        txtppn = RST!ppn
    End If
    
    SQL = "Select * from am_voucherin where novoucher ='" + txtnovoucher + "'"
    Set RST = OBJ.Execute(SQL)
    
    hapusgrid
    grid.Row = 1
    Do While Not RST.EOF
        grid.Col = 0
        Set grid.CellPicture = uncheck
        grid.TextMatrix(grid.Row, 1) = Format(RST!tgl, "dd/MM/yyyy")
        grid.TextMatrix(grid.Row, 2) = RST!nonota
        grid.TextMatrix(grid.Row, 3) = RST!keterangan
        grid.TextMatrix(grid.Row, 4) = RST!perkiraan
        grid.TextMatrix(grid.Row, 5) = Format(RST!jumlah, "##,###,###,##0.00")
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
        DoEvents
    Loop
    OBJ.Close
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub txtnovoucher_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        If chknonpajak.Value = Unchecked Then
            chknonpajak.Value = Checked
        Else
            chknonpajak.Value = Unchecked
        End If
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub txtnovoucher_LostFocus()
    If txtnovoucher = "" Then Exit Sub
    OpenVoucher
End Sub

Private Sub txttext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(posrow, poscol) = txttext
        grid.SetFocus
    End If
End Sub

Private Sub txttext_LostFocus()
    txttext.Visible = False
End Sub

Private Function GetNoVoucher() As String
    If chknonpajak.Value = 0 Then
        SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='1'"
    End If
    If chknonpajak.Value = 1 Then
        SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='0'"
    End If
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    GetNoVoucher = Trim(Str(RST!maxno + 1))
    OBJ.Close
Exit Function


    Dim tempyear As String
    Dim temp_kode As String
    Dim int_kode As Long
    tempyear = Format(Date, "yy") & "-"
    
    If chknonpajak.Value = 0 Then
        OBJ.Open dsn
        SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='1' and novoucher like '" & tempyear & "%'"
        Set RST = OBJ.Execute(SQL)
        If RST!maxno = "" Or IsNull(RST!maxno) Then
            temp_kode = "0001"
        End If
        
        If RST!maxno <> "" Then
            int_kode = Right(RST!maxno, 4)
            int_kode = int_kode + 1
            
            If (Len(Trim(Str(Right(int_kode, 4)))) = 1) Then
                temp_kode = "000" + Trim(Str(Right(int_kode, 1)))
            End If
            If (Len(Trim(Str(Right(int_kode, 4)))) = 2) Then
                temp_kode = "00" + Trim(Str(Right(int_kode, 2)))
            End If
            If (Len(Trim(Str(Right(int_kode, 4)))) = 3) Then
                temp_kode = "0" + Trim(Str(Right(int_kode, 3)))
            End If
            If (Len(Trim(Str(Right(int_kode, 4)))) = 4) Then
                temp_kode = Trim(Str(Right(int_kode, 4)))
            End If
        End If
        GetNoVoucher = Format(Date, "yy") & "-" & temp_kode
        OBJ.Close
    End If
    If chknonpajak.Value = 1 Then
        SQL = "select max(novoucher)as maxno from am_voucherhdr where ispajak='0'"
        OBJ.Open dsn
        Set RST = OBJ.Execute(SQL)
        GetNoVoucher = Trim(Str(RST!maxno + 1))
        OBJ.Close
    End If
End Function

Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function
Private Function SpyRoundUp(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.1) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRoundUp = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRoundUp = Val(arVal(0)) Else: SpyRoundUp = Val(arVal(0)) + 1
End Function

