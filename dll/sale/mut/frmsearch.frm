VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmsearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "d"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "View Stock"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Single Click To Choose"
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      GridLines       =   0
      SelectionMode   =   1
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin Chameleon.chameleonButton cmdcancel 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Cancel"
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
      MICON           =   "frmsearch.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmsearch.frx":031A
      Caption         =   "frmsearch.frx":033A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsearch.frx":03A6
      Keys            =   "frmsearch.frx":03C4
      Spin            =   "frmsearch.frx":0406
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3780
      TabIndex        =   5
      Top             =   2850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2490
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   150
      Width           =   735
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnusort1 
         Caption         =   "Sort Coloumn 1"
      End
      Begin VB.Menu mnusort2 
         Caption         =   "Sort Coloumn 2"
      End
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private m_SortColumn As Integer
Private m_SortAscending As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    If KeyAscii = 13 Then
        hasil = grid.TextMatrix(grid.Row, 0)
        hasil1 = grid.TextMatrix(grid.Row, 1)
        hasil2 = ""
        
        Unload Me
    End If
End Sub

Private Sub showtran()
    OBJ.Open dsn
    SQL = carisql1
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        
        If namatabel = "Surat Jalan " Or _
        namatabel = "Penjualan" Or _
        namatabel = "Retur Penjualan" Or _
        namatabel = "Pembayaran" Or _
        namatabel = "Koreksi" Or _
        namatabel = "Mutasi Barang" Or _
        namatabel = "Sales Order" Or _
        namatabel = "Desc/Referance" Or _
        namatabel = "Ganti Giro" Or _
        namatabel = "Pindah Gudang" Or _
        namatabel = "Susut Barang" Or _
        namatabel = "Request for Stock" Or _
        namatabel = "Sales Order  " Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Surat Jalan" Or _
        namatabel = "Surat Jalan  " Or _
        namatabel = "Apply to..." Then
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = grid.ColWidth(0)
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Sales Order " Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 2500
            grid.ColWidth(3) = 0
        Else
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 0
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        End If
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
        
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    Label2 = grid.Rows - 1 & " Records"
    grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
    Label4 = Mid$(grid.TextMatrix(0, 0), 3)
    m_SortColumn = 0
    Label3 = 0
    grid.Col = 0
    grid.Sort = flexSortStringAscending
    
    If namatabel = "Surat Jalan " Or _
    namatabel = "Penjualan" Or _
    namatabel = "Retur Penjualan" Or _
    namatabel = "Pembayaran" Or _
    namatabel = "Koreksi" Or _
    namatabel = "Mutasi Barang" Or _
    namatabel = "Pindah Gudang" Or _
    namatabel = "Susut Barang" Or _
    namatabel = "Request for Stock" Or _
    namatabel = "Desc/Referance" Or _
    namatabel = "Sales Order" Or _
    namatabel = "Ganti Giro" Or _
    namatabel = "Sales Order  " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Surat Jalan" Or _
    namatabel = "Surat Jalan  " Or _
    namatabel = "Apply to..." Then
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = grid.ColWidth(0)
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Sales Order " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 2500
        grid.ColWidth(3) = 0
    Else
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 0
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    End If
End Sub

Private Sub showtabel()
    If namatabel = "Company" _
    Or namatabel = "Company Account" _
    Or namatabel = "Company Type" _
    Or namatabel = "Acc Sparta" _
    Or namatabel = "Master Account" Then OBJ.Open dsn1 Else OBJ.Open dsn
    
    If namatabel = "Item on Sales Order" Then
        SQL = carisql1 + " order by a.lineitem"

    ElseIf namatabel = "User Level" Then
        SQL = carisql1 + " group by kode,keterangan"
    ElseIf namatabel = "Item" Then
        SQL = carisql1 + " where len(kodebarang)=8"
    ElseIf namatabel = "nolot" Then
        SQL = carisql1 + " Where flagprint = '4' order by tanggal desc"
    ElseIf namatabel = "Lot Packaging" Then
        SQL = carisql1 + " Where flag = '2' order by nolot desc"
    ElseIf namatabel = "Lot Packpending" Then
        SQL = carisql1 + " Where flag = '1' order by nolot desc"
    ElseIf namatabel = "Lot Item Gudang" Then
        SQL = carisql1 + " WHERE gudang = 'G1' order by nolot desc"
    ElseIf namatabel = "Base Bahan" Then
        SQL = carisql1 + " order by kodebarang asc"
    ElseIf namatabel = "Daftar Lot" Or namatabel = "Daftar Nomor Lot" Then
        SQL = carisql1 + " Where kodebarang='" & hasil8 & "' and flag='0'"
    ElseIf namatabel = "Item Lot" Then
        SQL = carisql1 + " Where nolot='" & hasil8 & "' and flag='0' group by kodebarang,namabarang,satuan"
    ElseIf namatabel = "Item over zak" Then
        SQL = carisql1 + " Where kodebarang like 'L%' order by kodebarang asc"
    Else
        SQL = carisql1
    End If
    
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Satuan" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 1000
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Customer" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 2000
            grid.ColWidth(2) = 4000
            grid.ColWidth(3) = 1500
        ElseIf namatabel = "Item on Sales Order" Then
            grid.ColWidth(1) = grid.ColWidth(0)
            grid.ColWidth(2) = 1500
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Lot Packaging" Or namatabel = "Lot Packpending" Or namatabel = "Lot Delete" Then
            grid.ColWidth(0) = 3000
        ElseIf namatabel = "Lot Item Gudang" Then
            grid.ColWidth(0) = 1600
            grid.ColWidth(1) = 450
            grid.ColWidth(2) = 1000
            grid.ColWidth(3) = 2800
            grid.ColWidth(4) = 970
            grid.ColWidth(5) = 0
            grid.ColWidth(6) = 800
            grid.ColWidth(7) = 0
            grid.ColWidth(8) = 0
            grid.ColWidth(9) = 0
            grid.ColWidth(10) = 0
        ElseIf namatabel = "Daftar Lot" Or namatabel = "Daftar Nomor Lot" Then
            grid.ColWidth(0) = 1600
            grid.ColWidth(1) = 1600
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
            grid.ColWidth(4) = 0
        ElseIf namatabel = "Item Lot" Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 2200
            grid.ColWidth(2) = 800
            grid.ColWidth(3) = 1000
        ElseIf namatabel = "Item over zak" Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 2200
        ElseIf namatabel = "Satuan." Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 2200
        Else
            grid.ColWidth(1) = 2940
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        End If
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If namatabel = "Company" _
    Or namatabel = "Company Account" _
    Or namatabel = "Company Type" _
    Or namatabel = "Acc Sparta" _
    Or namatabel = "Master Account" Then OBJ.Open dsn1 Else OBJ.Open dsn
    
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    Label2 = grid.Rows - 1 & " Records"
    grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
    Label4 = Mid$(grid.TextMatrix(0, 0), 3)
    m_SortColumn = 0
    Label3 = 0
    grid.Col = 0
    grid.Sort = flexSortStringAscending
        
    If namatabel = "Satuan" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Customer" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 2000
        grid.ColWidth(2) = 4000
        grid.ColWidth(3) = 1500
    ElseIf namatabel = "Item on Sales Order" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
        grid.ColWidth(3) = 0
    ElseIf namatabel = "nolot" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 0
        grid.ColWidth(2) = 1500
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Lot Packaging" Or namatabel = "Lot Packpending" Or namatabel = "Lot Delete" Then
        grid.ColWidth(0) = 1000
    ElseIf namatabel = "Lot Item Gudang" Then
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 450
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 2800
        grid.ColWidth(4) = 970
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 800
        grid.ColWidth(7) = 0
        grid.ColWidth(8) = 0
        grid.ColWidth(9) = 0
        grid.ColWidth(10) = 0
    ElseIf namatabel = "Daftar Lot" Or namatabel = "Daftar Nomor Lot" Then
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 1600
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
        grid.ColWidth(4) = 0
    ElseIf namatabel = "Item Lot" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
        grid.ColWidth(2) = 800
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Item over zak" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
    ElseIf namatabel = "Satuan." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    m_SortColumn = -1
    m_SortAscending = -1
    
    If namatabel = "Item on Sales Order" Then Check1.Visible = True
    
    If namatabel = "Penjualan" Or _
    namatabel = "Retur Penjualan" Or _
    namatabel = "Pembayaran" Or _
    namatabel = "Apply to..." Or _
    namatabel = "Koreksi" Or _
    namatabel = "Mutasi Barang" Or _
    namatabel = "Faktur Pajak Standar" Or _
    namatabel = "Surat Jalan " Or _
    namatabel = "Surat Jalan" Or _
    namatabel = "Surat Jalan  " Or _
    namatabel = "Sales Order" Or _
    namatabel = "Sales Order " Or _
    namatabel = "Ganti Giro" Or _
    namatabel = "Desc/Referance" Or _
    namatabel = "Pindah Gudang" Or _
    namatabel = "Susut Barang" Or _
    namatabel = "Request for Stock" Or _
    namatabel = "Sales Order  " Then
        Label1.Visible = False
        txtsearch.Visible = False
        grid.Top = 120
        grid.Height = 2295
        Me.Caption = "Searching Transaksi " & namatabel
        If namatabel = "Sales Order " Then
            Me.Width = 6000
            grid.Width = 5655
            cmdcancel.Width = 5655
        End If
        showtran
        Exit Sub
    End If
    
    Me.Caption = "Searching Tabel " & namatabel
    If namatabel = "Customer" Or namatabel = "Lot Item Gudang" Then
        Me.Width = 9600
        grid.Width = 9255
        cmdcancel.Width = 9255
    End If
    showtabel
    
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
End Sub

Private Sub grid_Click()
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    'left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and
    If namatabel = "Item on Sales Order" And Check1.Value = 1 Then
        OBJ.Open dsn
        SQL = "select isnull(sum(a.qty),0)'stok' from am_bpblin a left join am_bpbhdr b on a.nobpb=b.nobpb where a.kodebarang = '" & grid.TextMatrix(grid.Row, 0) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "' and b.kodegudang = '" & setup3 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtnil1 = RST!stok Else txtnil1 = 0
        
        SQL = "select isnull(sum(a.qty),0)'stok' from am_sjlin a left join am_sjhdr b on a.nosj=b.nosj and a.tglsj=b.tglsj left join am_sjapp c on a.nosj=c.nosj and a.kodebarang=c.kodebarang and a.kodesatuan=c.kodesatuan where (c.flag2 is null or c.flag2 <> '9') and b.kodegudang = '" & setup3 & "' and a.kodebarang = '" & grid.TextMatrix(grid.Row, 0) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtnil1 = txtnil1 - RST!stok
        OBJ.Close
        
        grid.ToolTipText = "Stock " & grid.TextMatrix(grid.Row, 0) & " on " & setup3 & " : " & Format(txtnil1, "general number")
        Exit Sub
    End If
    
    Label3 = grid.MouseCol
    If grid.MouseRow > 0 Then
        If namatabel = "nolot" Then
            hasil = grid.TextMatrix(grid.Row, 1)
            hasil1 = grid.TextMatrix(grid.Row, 2)
            hasil2 = ""
            Unload Me
            Exit Sub
        ElseIf namatabel = "Lot Packaging" Or namatabel = "Lot Packpending" Or namatabel = "Lot Delete" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            Unload Me
            Exit Sub
        ElseIf namatabel = "Lot Item Gudang" Then
            hasil = grid.TextMatrix(grid.Row, 2)    'kode barang
            hasil1 = grid.TextMatrix(grid.Row, 5)   'kd satuan
            hasil2 = grid.TextMatrix(grid.Row, 6)   'qty
            hasil3 = grid.TextMatrix(grid.Row, 0)   'nolot
            hasil4 = grid.TextMatrix(grid.Row, 1)   'palet
            hasil5 = grid.TextMatrix(grid.Row, 7)   'isi
            hasil6 = grid.TextMatrix(grid.Row, 8)   'kg
            hasil7 = grid.TextMatrix(grid.Row, 9)   'hpp
            hasil8 = grid.TextMatrix(grid.Row, 10)  'total pack
            Unload Me
            Exit Sub
        ElseIf namatabel = "Daftar Lot" Or namatabel = "Daftar Nomor Lot" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = Format(grid.TextMatrix(grid.Row, 2), "#,##0.00")
            hasil3 = Format(grid.TextMatrix(grid.Row, 3), "#,##0.00")
            hasil4 = Format(grid.TextMatrix(grid.Row, 4), "#,##0.00")
            Unload Me
            Exit Sub
        ElseIf namatabel = "Item Lot" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = Format(grid.TextMatrix(grid.Row, 2), "#,##0.00")
            hasil3 = grid.TextMatrix(grid.Row, 3)
            Unload Me
            Exit Sub
        ElseIf namatabel = "Item over zak" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            Unload Me
            Exit Sub
        ElseIf namatabel = "Satuan." Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            Unload Me
            Exit Sub
        End If
        hasil = grid.TextMatrix(grid.Row, 0)
        hasil1 = grid.TextMatrix(grid.Row, 1)
        hasil2 = ""
        
        Unload Me
        Exit Sub
    End If
    If grid.MouseCol <> m_SortColumn Then
        If m_SortColumn >= 0 Then
            grid.TextMatrix(0, m_SortColumn) = _
                Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
        m_SortColumn = grid.MouseCol
        
        m_SortAscending = True
        grid.TextMatrix(0, m_SortColumn) = _
            "> " & grid.TextMatrix(0, m_SortColumn)
    Else
        m_SortAscending = Not m_SortAscending
        
        If m_SortAscending Then
            grid.TextMatrix(0, m_SortColumn) = _
                "> " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        Else
            grid.TextMatrix(0, m_SortColumn) = _
                "< " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
    End If
    
    Label4 = Mid$(grid.TextMatrix(0, Label3), 3)
    grid.Row = 1
    grid.RowSel = grid.Rows - 1
    grid.Col = m_SortColumn

    If m_SortAscending Then
        grid.Sort = flexSortStringAscending
    Else
        grid.Sort = flexSortStringDescending
    End If
    
    If txtsearch.Visible = True Then txtsearch.SetFocus
End Sub

Private Sub txtsearch_Change()
    If namatabel = "Company" _
    Or namatabel = "Company Account" _
    Or namatabel = "Company Type" _
    Or namatabel = "Acc Sparta" _
    Or namatabel = "Master Account" Then OBJ.Open dsn1 Else OBJ.Open dsn
    
    If namatabel = "User Level" Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%' group by kode,keterangan"
    ElseIf namatabel = "Satuan " Or namatabel = "Company Account" Then
        SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Cek/Giro" Or namatabel = "Faktur" Or namatabel = "Barang " Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Item on Sales Order" Then
        If Label4 = "namabarang" Then
            SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%' order by a.lineitem"
        Else
            SQL = carisql1 + " and a." + Label4 + " like '" + txtsearch + "%' order by a.lineitem"
        End If
    ElseIf namatabel = "ItemG3" Then
        If Label4 = "namabarang" Then
            SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%'"
        Else
            SQL = carisql1 + " and a." + Label4 + " like '" + txtsearch + "%'"
        End If
    ElseIf namatabel = "Item" Then
        SQL = carisql1 + " where len(kodebarang)=8 and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Item " Then
        SQL = carisql1 + " and len(kodebarang)=8 and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "nolot" Then
        SQL = carisql1 + " Where flagprint = '4' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Lot Packaging" Then
        SQL = carisql1 + " Where flag = '2' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Lot Packpending" Then
        SQL = carisql1 + " Where flag = '1' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Lot Item Gudang" Then
        SQL = carisql1 + " WHERE gudang = 'G1' and " + Label4 + " like '" + txtsearch + "%' order by nolot desc"
    ElseIf namatabel = "Base Bahan" Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%' order by kodebarang asc"
    ElseIf namatabel = "Daftar lot" Then
        'SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%'"
        SQL = carisql1 + " Where kodebarang='" + hasil8 + "' and flag='0' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Daftar Nomor Lot" Then
        SQL = carisql1 + " Where kodebarang='" + hasil8 + "' and flag='0' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Item Lot" Then
        SQL = carisql1 + " Where nolot='" + hasil8 + "' and flag='0' and " + Label4 + " like '" + txtsearch + "%' Group By kodebarang,namabarang,satuan"
    ElseIf namatabel = "Item over zak" Then
        SQL = carisql1 + " Where kodebarang like 'L%' and " + Label4 + " like '" + txtsearch + "%' Order By kodebarang asc"
    ElseIf namatabel = "Satuan." Then
        OBJ.Close
        Exit Sub
    Else
        SQL = carisql1 + " where " + Label4 + " like '" + txtsearch + "%'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Satuan" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 1000
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Customer" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 2000
            grid.ColWidth(2) = 4000
            grid.ColWidth(3) = 1500
        ElseIf namatabel = "Item on Sales Order" Then
            grid.ColWidth(1) = grid.ColWidth(0)
            grid.ColWidth(2) = 1500
            grid.ColWidth(3) = 0
        Else
            grid.ColWidth(1) = 2940
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        End If
        OBJ.Close
        Label2 = ""
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    Label2 = grid.Rows - 1 & " Records"
    OBJ.Close
    grid.TextMatrix(0, Label3) = _
            "> " & grid.TextMatrix(0, Label3)
    grid.Sort = flexSortStringAscending
        
    If namatabel = "Satuan" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Customer" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 2000
        grid.ColWidth(2) = 4000
        grid.ColWidth(3) = 1500
    ElseIf namatabel = "Item on Sales Order" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
        grid.ColWidth(3) = 0
    ElseIf namatabel = "nolot" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 0
        grid.ColWidth(2) = 1500
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Lot Packaging" Or namatabel = "Lot Packpending" Or namatabel = "Lot Delete" Then
        grid.ColWidth(0) = 4000
    ElseIf namatabel = "Lot Item Gudang" Then
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 450
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 2800
        grid.ColWidth(4) = 970
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 800
        grid.ColWidth(7) = 0
        grid.ColWidth(8) = 0
        grid.ColWidth(9) = 0
        grid.ColWidth(10) = 0
    ElseIf namatabel = "Daftar Lot" Or namatabel = "Daftar Nomor Lot" Then
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 1600
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
        grid.ColWidth(4) = 0
    ElseIf namatabel = "Item Lot" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
        grid.ColWidth(2) = 800
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Item over zak" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
    ElseIf namatabel = "Satuan." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2200
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
    If KeyAscii = 13 And grid.Rows = 2 Then
        hasil = grid.TextMatrix(1, 0)
        hasil1 = grid.TextMatrix(1, 1)
        hasil2 = ""
        Unload Me
        Exit Sub
    End If
    If Label3 = "" Then KeyAscii = 0
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As Control
  
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSHFlexGrid Then
          If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        End If
    Next ctl
End Sub
