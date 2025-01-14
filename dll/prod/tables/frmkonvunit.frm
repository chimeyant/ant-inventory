VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmkonvunit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Konversi Satuan"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbkode 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      TabIndex        =   13
      Top             =   720
      Width           =   1305
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5295
      _Version        =   851970
      _ExtentX        =   9340
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Konversi"
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
      Begin VB.TextBox txtunitkonv 
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
         Left            =   4080
         TabIndex        =   9
         ToolTipText     =   "Click here to show unit"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtunit 
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
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Click here to show unit"
         Top             =   480
         Width           =   1455
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Calculator      =   "frmkonvunit.frx":0000
         Caption         =   "frmkonvunit.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmkonvunit.frx":008C
         Keys            =   "frmkonvunit.frx":00AA
         Spin            =   "frmkonvunit.frx":00EC
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "="
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
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "KONVERSI SATUAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "SATUAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.TextBox txtnmbrg 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Click here to show item"
      Top             =   360
      Width           =   4095
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   5640
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
      MICON           =   "frmkonvunit.frx":0114
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
      Height          =   2955
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Click here to update or delete"
      Top             =   2640
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   5212
      _Version        =   393216
      BackColor       =   -2147483628
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1800
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
      MICON           =   "frmkonvunit.frx":042E
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
      Left            =   6120
      TabIndex        =   15
      Top             =   5640
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
      MICON           =   "frmkonvunit.frx":0748
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmddelete 
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Delete"
      ENAB            =   0   'False
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
      MICON           =   "frmkonvunit.frx":0A62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Daftar Konversi Satuan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   7815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      Left            =   360
      TabIndex        =   14
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
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
      Left            =   360
      TabIndex        =   1
      Top             =   375
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmkonvunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private str_kode, str_kode1, str_kode2 As String

Private Sub cmdclear_Click()
    txtnmbrg = ""
    txtunit = ""
    txtunitkonv = ""
    cmbkode = ""
    txtnilai = 0
    str_kode = ""
    str_kode1 = ""
    str_kode2 = ""
    cmdsave.Caption = "Save"
    cmddelete.Enabled = False
    txtnmbrg.Enabled = True
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddelete_Click()
    OBJ.Open dsn
    SQL = "Delete from am_apunit_konversi Where kdbrg = '" & str_kode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "Unit convertion deleted successfuly", vbInformation, AppName
    cmdclear_Click
    Call showdata
End Sub

Private Sub cmdsave_Click()
    If txtnmbrg = "" Then Exit Sub
    If cmbkode = "" Then
        MsgBox "Sub Divisi tidak boleh kosong", vbExclamation, AppName
        Exit Sub
    End If
    If txtunit = "" Or txtunitkonv = "" Then
        MsgBox "Satuan/Satuan Konversi harus diisi terlebih dahulu", vbExclamation, AppName
        Exit Sub
    End If
    
    If txtnilai = "" Or IsNull(txtnilai) Or txtnilai = "0" Then
        MsgBox "Convertion Value isnull", vbExclamation, AppName
        Exit Sub
    End If
        
    OBJ.Open dsn
    SQL = "Select * From am_apunit_konversi Where kdbrg='" & str_kode & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    If cmdsave.Caption = "Save" Then
        If Not RST.EOF Then
            MsgBox "Unit already has a conversion", vbCritical, AppName
        Else
            With RST
                .AddNew
                !kdbrg = str_kode
                !kodesatuan = str_kode1
                !KodeSatuanKonv = str_kode2
                !nilai = txtnilai
                !DateCreate = Date
                !DateUpdate = "1900-01-01"
                !UserCreate = nmuser
                !UserUpdate = ""
                !Divisi = cmbkode
                .Update
            End With
            MsgBox "Unit saved successfully", vbInformation, AppName
        End If
    Else
        With RST
                !kodesatuan = str_kode1
                !KodeSatuanKonv = str_kode2
                !nilai = txtnilai
                !DateUpdate = Date
                !UserUpdate = nmuser
                !Divisi = cmbkode
                .Update
            End With
            MsgBox "Unit updated successfully", vbInformation, AppName
    End If
    OBJ.Close
    cmdclear_Click
    Call showdata
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
    setGrid
    initGrid
    showdata
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 500
        .ColWidth(4) = 700
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
End Sub

Private Sub initGrid()
    With grid
        .Cols = 10
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "SATUAN"
        .TextMatrix(0, 5) = ""
        .TextMatrix(0, 6) = "KONVERSI"
        .TextMatrix(0, 7) = "KDSATUAN"
        .TextMatrix(0, 8) = "KDKONVERSI"
        .TextMatrix(0, 9) = "DIVISI"
    End With
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
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub

Private Sub showdata()
    OBJ.Open dsn
    SQL = "Select a.kdbrg,d.NamaBarang,a.KodeSatuan,b.NamaSatuan,a.Nilai,a.KodeSatuanKonv,c.NamaSatuan 'konv',a.Divisi"
    SQL = SQL + " From am_apunit_konversi a inner join am_apunit b on a.KodeSatuan = b.KodeSatuan"
    SQL = SQL + " inner join am_apunit c on a.KodeSatuanKonv = c.KodeSatuan"
    SQL = SQL + " inner join am_apitemmst d on a.kdbrg = d.KodeBarang"
    Set RST = OBJ.Execute(SQL)
    hapusgrid
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!kdbrg
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = "1"
        grid.TextMatrix(grid.Row, 4) = RST!namasatuan
        grid.TextMatrix(grid.Row, 5) = RST!nilai
        grid.TextMatrix(grid.Row, 6) = RST!konv
        grid.TextMatrix(grid.Row, 7) = RST!kodesatuan
        grid.TextMatrix(grid.Row, 8) = RST!KodeSatuanKonv
        grid.TextMatrix(grid.Row, 9) = RST!Divisi
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    str_kode = grid.TextMatrix(grid.Row, 1)
    txtnmbrg = grid.TextMatrix(grid.Row, 2)
    str_kode1 = grid.TextMatrix(grid.Row, 7)
    txtunit = grid.TextMatrix(grid.Row, 4)
    str_kode2 = grid.TextMatrix(grid.Row, 8)
    txtunitkonv = grid.TextMatrix(grid.Row, 6)
    txtnilai = grid.TextMatrix(grid.Row, 5)
    cmbkode = grid.TextMatrix(grid.Row, 9)
    cmdsave.Caption = "Update"
    cmddelete.Enabled = True
    txtnmbrg.Enabled = False
End Sub

Private Sub txtnmbrg_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    frmsearch.Show vbModal
End Sub

Private Sub txtnmbrg_GotFocus()
    If hasil = "" Then Exit Sub
    str_kode = hasil
    txtnmbrg = hasil1
    hasil = ""
    hasil1 = ""
    
    OBJ.Open dsn
    'SQL = "Select * From am_apitemmst Where kodebarang = '" & str_kode & "'"
    SQL = "Select a.*,b.NamaSatuan from am_apitemmst a inner join am_apunit b"
    SQL = SQL + " on a.KodeSatuan = b.KodeSatuan Where kodebarang='" & str_kode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        cmbkode = RST!KodeProduk
        txtunit = RST!namasatuan
        str_kode1 = RST!kodesatuan
    End If
    OBJ.Close
End Sub

Private Sub txtunit_Click()
    If txtnmbrg = "" Then Exit Sub
    carisql1 = "select kodesatuan, namasatuan from am_apunit"
    namatabel = "Satuan Bahan Baku"
    
    frmsearch.Show
End Sub

Private Sub txtunit_GotFocus()
    If hasil = "" Then Exit Sub
    str_kode1 = hasil
    txtunit = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub txtunitkonv_Click()
    If txtnmbrg = "" Then Exit Sub
    carisql1 = "select kodesatuan, namasatuan from am_apunit"
    namatabel = "Satuan Bahan Baku"
    
    frmsearch.Show
End Sub

Private Sub txtunitkonv_GotFocus()
    If hasil = "" Then Exit Sub
    str_kode2 = hasil
    txtunitkonv = hasil1
    hasil = ""
    hasil1 = ""
End Sub
