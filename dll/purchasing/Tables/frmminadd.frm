VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmminadd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Minimum Stock"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4005
      Left            =   45
      TabIndex        =   5
      Top             =   660
      Width           =   8505
      _Version        =   851970
      _ExtentX        =   15002
      _ExtentY        =   7064
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   10
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Add Minimum Stock"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "txtnobukti"
      Item(0).Control(1)=   "cmdsearch1"
      Item(0).Control(2)=   "cmbbulan"
      Item(0).Control(3)=   "txtnilai"
      Item(0).Control(4)=   "txtmin"
      Item(0).Control(5)=   "Label2"
      Item(0).Control(6)=   "Label3"
      Item(0).Control(7)=   "Label6"
      Item(0).Control(8)=   "Shape1"
      Item(0).Control(9)=   "Shape2"
      Item(0).Control(10)=   "lblsat"
      Item(0).Control(11)=   "lblsatuan"
      Item(0).Control(12)=   "lblnamabarang"
      Item(1).Caption =   "List Minimum Stock"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "grid"
      Begin VB.ComboBox cmbbulan 
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
         ItemData        =   "frmminadd.frx":0000
         Left            =   2580
         List            =   "frmminadd.frx":0028
         TabIndex        =   9
         Top             =   2220
         Width           =   1290
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   6
         Top             =   735
         Width           =   1275
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   750
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Kode Barang"
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
         MICON           =   "frmminadd.frx":009B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   3135
         Left            =   -69910
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
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
         _Band(0).Cols   =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   1755
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         Calculator      =   "frmminadd.frx":03B5
         Caption         =   "frmminadd.frx":03D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmminadd.frx":0441
         Keys            =   "frmminadd.frx":045F
         Spin            =   "frmminadd.frx":04A1
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,##0.000;(##,###,##0.000)"
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
      Begin TDBNumber6Ctl.TDBNumber txtmin 
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   2730
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         Calculator      =   "frmminadd.frx":04C9
         Caption         =   "frmminadd.frx":04E9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmminadd.frx":0555
         Keys            =   "frmminadd.frx":0573
         Spin            =   "frmminadd.frx":05B5
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,##0.000;(##,###,##0.000);0"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,##0.000;(##,###,##0.000)"
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
      Begin VB.Label lblnamabarang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   255
         TabIndex        =   17
         Top             =   1245
         Width           =   7845
      End
      Begin VB.Label lblsatuan 
         BackColor       =   &H8000000E&
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
         Left            =   3945
         TabIndex        =   16
         Top             =   2730
         Width           =   570
      End
      Begin VB.Label lblsat 
         BackColor       =   &H8000000E&
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
         Left            =   3945
         TabIndex        =   15
         Top             =   1755
         Width           =   570
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   360
         Left            =   2565
         Shape           =   4  'Rounded Rectangle
         Top             =   2685
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   360
         Left            =   2565
         Shape           =   4  'Rounded Rectangle
         Top             =   1710
         Width           =   1995
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Stock"
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
         Left            =   1290
         TabIndex        =   14
         Top             =   2730
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
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
         Left            =   2145
         TabIndex        =   13
         Top             =   2250
         Width           =   270
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian rata-rata / bulan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   12
         Top             =   1785
         Width           =   2160
      End
   End
   Begin VB.ComboBox cmbkode 
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
      Left            =   1410
      TabIndex        =   3
      Top             =   240
      Width           =   1305
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7635
      TabIndex        =   0
      Top             =   4740
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
      MICON           =   "frmminadd.frx":05DD
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
      Left            =   6675
      TabIndex        =   1
      Top             =   4740
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
      MICON           =   "frmminadd.frx":08F7
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
      Left            =   5715
      TabIndex        =   2
      Top             =   4740
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
      MICON           =   "frmminadd.frx":0C11
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Sub Divisi"
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
      Left            =   285
      TabIndex        =   4
      Top             =   270
      Width           =   1050
   End
End
Attribute VB_Name = "frmminadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim str1, str2 As String

Private Sub cmbbulan_Click()
    If cmbbulan = "1 Bulan" Then txtmin = txtnilai: str2 = "1"
    If cmbbulan = "2 Bulan" Then txtmin = txtnilai * 2: str2 = "2"
    If cmbbulan = "3 Bulan" Then txtmin = txtnilai * 3: str2 = "3"
    If cmbbulan = "4 Bulan" Then txtmin = txtnilai * 4: str2 = "4"
    If cmbbulan = "5 Bulan" Then txtmin = txtnilai * 5: str2 = "5"
    If cmbbulan = "6 Bulan" Then txtmin = txtnilai * 6: str2 = "6"
    If cmbbulan = "7 Bulan" Then txtmin = txtnilai * 7: str2 = "7"
    If cmbbulan = "8 Bulan" Then txtmin = txtnilai * 8: str2 = "8"
    If cmbbulan = "9 Bulan" Then txtmin = txtnilai * 9: str2 = "9"
    If cmbbulan = "10 Bulan" Then txtmin = txtnilai * 10: str2 = "10"
    If cmbbulan = "11 Bulan" Then txtmin = txtnilai * 11: str2 = "11"
    If cmbbulan = "12 Bulan" Then txtmin = txtnilai * 12: str2 = "12"
End Sub

Private Sub cmbkode_Click()
    hapusgrid
    OBJ.Open dsn
    SQL = "Select a.kodebarang,b.NamaBarang,a.usageAvrg,a.range,c.NamaSatuan "
    SQL = SQL + "From am_stokminimum a left join am_apitemmst b on a.kodebarang=b.KodeBarang "
    SQL = SQL + "left join am_apunit c on a.kodesatuan = c.KodeSatuan Where a.kodeProduk = '" & cmbkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do Until RST.EOF
        With grid
            .TextMatrix(.Row, 0) = .Row
            .TextMatrix(.Row, 1) = RST!kodebarang
            .TextMatrix(.Row, 2) = RST!namabarang
            .TextMatrix(.Row, 3) = Format(RST!UsageAvrg, "###,###,##0.000")
            .TextMatrix(.Row, 4) = RST!range & " Bulan"
            .TextMatrix(.Row, 5) = Format((RST!UsageAvrg * RST!range), " ###,###,##0.000")
            .TextMatrix(.Row, 6) = RST!namasatuan
            .Rows = .Rows + 1
            .Row = .Row + 1
        End With
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdadd_Click()
    If cmbkode = "" Then Exit Sub
    If txtnobukti = "" Or txtnilai = "0.000" Then
        MsgBox "Data Entry Not Complete..", vbExclamation, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From am_stokminimum where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kodebarang = txtnobukti
        !UsageAvrg = txtnilai
        !range = str2
        !kodesatuan = str1
        !kodeproduk = cmbkode
        !DateEntry = Date
        !IdEntry = nmuser
        !DateUpdate = "1900-01-01"
        !IdUpdate = ""
        !flag = "0"
        .Update
    End With
    
    SQL = "Select * From am_invmin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kodebarang = txtnobukti
        !minstock = txtmin
        .Update
    End With
    
    OBJ.Close
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, AppName
    cmdclear_Click
    cmbkode_Click
End Sub

Private Sub cmdclear_Click()
    lblnamabarang = ""
    cmbbulan = ""
    txtnobukti = ""
    txtnilai = "0.000"
    txtmin = "0.000"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    If cmbkode = "" Then Exit Sub
    carisql1 = "select a.KodeBarang,a.NamaBarang,a.KodeSatuan,b.NamaSatuan,c.flag from am_apitemmst a"
    carisql1 = carisql1 + " left join am_apunit b on a.KodeSatuan = b.KodeSatuan"
    carisql1 = carisql1 + " left join am_stokminimum c on a.KodeBarang = c.KodeBarang where a.KodeProduk = '" & cmbkode & "'"
    
    namatabel = "Barang per Divisi."
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    lblnamabarang = hasil1
    lblsat = hasil2
    lblsatuan = hasil2
    str1 = hasil3
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    cekminstok
End Sub

Private Sub cekminstok()
    OBJ.Open dsn
    SQL = "Select * From am_stokminimum Where KodeBarang='" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Minimum Stock is Already Exist", vbCritical, AppName
        cmdclear_Click
    End If
    OBJ.Close
End Sub
Private Sub Form_Load()
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
    
    grid.Cols = 7
    grid.TextMatrix(0, 1) = "Kode"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Usage/Month"
    grid.TextMatrix(0, 4) = "Range"
    grid.TextMatrix(0, 5) = "Qty Minimum"
    grid.TextMatrix(0, 6) = "Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1350
    grid.ColWidth(4) = 800
    grid.ColAlignmentFixed(4) = flexAlignRightCenter
    grid.ColWidth(5) = 1200
    grid.ColAlignment(5) = flexAlignRightCenter
    grid.ColWidth(6) = 800
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
        grid.TextMatrix(grid.Row, 6) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.TextMatrix(1, 0) = ""
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1350
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1200
    grid.ColWidth(6) = 800
End Sub

Private Sub txtnilai_Change()
    If cmbbulan = "" Then Exit Sub
    If cmbbulan = "1 Bulan" Then txtmin = txtnilai
    If cmbbulan = "2 Bulan" Then txtmin = txtnilai * 2
    If cmbbulan = "3 Bulan" Then txtmin = txtnilai * 3
    If cmbbulan = "4 Bulan" Then txtmin = txtnilai * 4
    If cmbbulan = "5 Bulan" Then txtmin = txtnilai * 5
    If cmbbulan = "6 Bulan" Then txtmin = txtnilai * 6
    If cmbbulan = "7 Bulan" Then txtmin = txtnilai * 7
    If cmbbulan = "8 Bulan" Then txtmin = txtnilai * 8
    If cmbbulan = "9 Bulan" Then txtmin = txtnilai * 9
    If cmbbulan = "10 Bulan" Then txtmin = txtnilai * 10
    If cmbbulan = "11 Bulan" Then txtmin = txtnilai * 11
    If cmbbulan = "12 Bulan" Then txtmin = txtnilai * 12
End Sub
