VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmminedit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Minimum Stok"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcari 
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
      Left            =   5175
      TabIndex        =   15
      Top             =   540
      Width           =   4635
   End
   Begin VB.ComboBox cmbbulan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmminedit.frx":0000
      Left            =   7125
      List            =   "frmminedit.frx":0028
      TabIndex        =   4
      Top             =   2415
      Width           =   1755
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
      Left            =   1110
      TabIndex        =   1
      Top             =   135
      Width           =   1305
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3360
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Click Row to Update"
      Top             =   525
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   5927
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   3495
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
      MICON           =   "frmminedit.frx":009B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   315
      Left            =   7125
      TabIndex        =   5
      Top             =   2040
      Width           =   2220
      _Version        =   65536
      _ExtentX        =   3916
      _ExtentY        =   556
      Calculator      =   "frmminedit.frx":03B5
      Caption         =   "frmminedit.frx":03D5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmminedit.frx":0441
      Keys            =   "frmminedit.frx":045F
      Spin            =   "frmminedit.frx":04A1
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      Height          =   315
      Left            =   7125
      TabIndex        =   6
      Top             =   2820
      Width           =   2220
      _Version        =   65536
      _ExtentX        =   3916
      _ExtentY        =   556
      Calculator      =   "frmminedit.frx":04C9
      Caption         =   "frmminedit.frx":04E9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmminedit.frx":0555
      Keys            =   "frmminedit.frx":0573
      Spin            =   "frmminedit.frx":05B5
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   8325
      TabIndex        =   16
      Top             =   3495
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
      MICON           =   "frmminedit.frx":05DD
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
      Left            =   7395
      TabIndex        =   17
      Top             =   3495
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
      MICON           =   "frmminedit.frx":08F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbltotalrow 
      Alignment       =   1  'Right Justify
      Caption         =   "0 Records"
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
      Left            =   3150
      TabIndex        =   18
      Top             =   180
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   4605
      Picture         =   "frmminedit.frx":0C11
      Stretch         =   -1  'True
      Top             =   375
      Width           =   585
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang :"
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
      Left            =   6090
      TabIndex        =   14
      Top             =   1305
      Width           =   1050
   End
   Begin VB.Label lblsatuan 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9360
      TabIndex        =   13
      Top             =   2850
      Width           =   660
   End
   Begin VB.Label lblsat 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9360
      TabIndex        =   12
      Top             =   2055
      Width           =   660
   End
   Begin VB.Label lblkode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   7125
      TabIndex        =   11
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemakaian rata-rata / bulan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4575
      TabIndex        =   10
      Top             =   2070
      Width           =   2610
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6825
      TabIndex        =   9
      Top             =   2445
      Width           =   270
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Stock :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5670
      TabIndex        =   8
      Top             =   2850
      Width           =   1500
   End
   Begin VB.Label lblnamabarang 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   4605
      TabIndex        =   7
      Top             =   1665
      Width           =   5385
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
      Left            =   120
      TabIndex        =   2
      Top             =   165
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00404040&
      Height          =   2340
      Left            =   4485
      Shape           =   4  'Rounded Rectangle
      Top             =   1050
      Width           =   5610
   End
End
Attribute VB_Name = "frmminedit"
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
    txtcari = ""
    grid.Rows = 2

    OBJ.Open dsn
    SQL = "Select a.kodebarang,b.NamaBarang'NAMA BARANG',a.usageAvrg,a.range, "
    SQL = SQL + "CAST((a.usageAvrg*a.range)as numeric(10,3))'QTY MINIMUM',a.KodeSatuan,c.NamaSatuan "
    SQL = SQL + "From am_stokminimum a left join am_apitemmst b on a.kodebarang=b.KodeBarang "
    SQL = SQL + "left join am_apunit c on a.kodesatuan = c.KodeSatuan Where a.kodeProduk = '" & cmbkode & "'"
    
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(0) = 0
        grid.TextMatrix(0, 1) = "NAMA BARANG"
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
        grid.TextMatrix(0, 4) = "QTY MINIMUM"
        grid.ColWidth(4) = 1500
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 0
        lbltotalrow = "0 Records"
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    grid.ColAlignmentFixed(4) = flexAlignRightCenter
    grid.ColAlignment(4) = flexAlignRightCenter
    
    SQL = "Select COUNT(kodebarang)'jml' From am_stokminimum Where KodeProduk ='" & cmbkode & "'"
    Set RST = OBJ.Execute(SQL)
    lbltotalrow = RST!jml & " Records"
    OBJ.Close
End Sub

Private Sub cmdclear_Click()
    txtcari = ""
    lblkode = ""
    lblnamabarang = ""
    txtnilai = "0.000"
    cmbbulan = ""
    txtmin = "0.000"
    lblsat = ""
    lblsatuan = ""
    cmbkode = ""
    lbltotalrow = "0 Records"
    hapusgrid
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdupdate_Click()
    If lblkode = "" Then Exit Sub
        
    If MsgBox("Are you sure, you want to update this data...", vbQuestion + vbYesNo, "KONFIRMASI UPDATE DATA") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "Select * From am_stokminimum Where KodeBarang='" & lblkode & "'"
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        !UsageAvrg = txtnilai
        !range = str2
        !DateUpdate = Date
        !IdUpdate = nmuser
        !flag = "1"
        .Update
    End With
    
    SQL = "Select * From am_invmin Where kodebarang='" & lblkode & "'"
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        !minstock = txtmin
        .Update
    End With
    
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    txtcari = ""
    lblkode = ""
    lblnamabarang = ""
    txtnilai = "0.000"
    txtmin = "0.000"
    cmbbulan = ""
    lblsat = ""
    lblsatuan = ""
    cmbkode_Click
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
    grid.TextMatrix(0, 0) = "Kode"
    grid.TextMatrix(0, 1) = "NAMA BARANG"
    grid.TextMatrix(0, 2) = "Usage/Month"
    grid.TextMatrix(0, 3) = "Range"
    grid.TextMatrix(0, 4) = "QTY MINIMUM"
    grid.TextMatrix(0, 5) = "KodeSatuan"
    grid.TextMatrix(0, 6) = "Satuan"
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 0
    grid.ColWidth(4) = 1500
    grid.ColAlignmentFixed(4) = flexAlignRightCenter
    grid.ColAlignment(4) = flexAlignRightCenter
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 0
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        If grid.Rows = grid.Row + 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 0
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 0
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    lblkode = grid.TextMatrix(grid.Row, 0)
    lblnamabarang = grid.TextMatrix(grid.Row, 1)
    txtnilai = grid.TextMatrix(grid.Row, 2)
    cmbbulan = grid.TextMatrix(grid.Row, 3) & " Bulan"
    str2 = grid.TextMatrix(grid.Row, 3)
    txtmin = grid.TextMatrix(grid.Row, 4)
    str1 = grid.TextMatrix(grid.Row, 5)
    lblsat = grid.TextMatrix(grid.Row, 6)
    lblsatuan = grid.TextMatrix(grid.Row, 6)
End Sub

Private Sub txtcari_Change()
    OBJ.Open dsn
    SQL = "Select a.kodebarang,b.NamaBarang'NAMA BARANG',a.usageAvrg,a.range,"
    SQL = SQL + "CAST((a.usageAvrg*a.range)as numeric(10,3))'QTY MINIMUM',a.KodeSatuan,c.NamaSatuan "
    SQL = SQL + "From am_stokminimum a left join am_apitemmst b on a.kodebarang=b.KodeBarang "
    SQL = SQL + "left join am_apunit c on a.kodesatuan = c.KodeSatuan Where a.kodeProduk = '" & cmbkode & "' "
    SQL = SQL + "and b.NamaBarang like '" & txtcari & "%'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(0) = 0
        grid.TextMatrix(0, 1) = "NAMA BARANG"
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
        grid.TextMatrix(0, 4) = "QTY MINIMUM"
        grid.ColWidth(4) = 1500
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 0
        lbltotalrow = "0 Records"
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    grid.ColAlignmentFixed(4) = flexAlignRightCenter
    grid.ColAlignment(4) = flexAlignRightCenter
    
    SQL = "Select COUNT(b.NamaBarang)'jml' From am_stokminimum a "
    SQL = SQL + "left join am_apitemmst b on a.kodebarang=b.KodeBarang "
    SQL = SQL + "left join am_apunit c on a.kodesatuan = c.KodeSatuan "
    SQL = SQL + "Where a.kodeProduk = '" & cmbkode & "' and b.NamaBarang like '" & txtcari & "%'"
    Set RST = OBJ.Execute(SQL)
    lbltotalrow = RST!jml & " Records"
    
    OBJ.Close
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
