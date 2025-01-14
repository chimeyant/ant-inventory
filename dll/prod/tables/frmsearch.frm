VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "m"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
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
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4560
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox cmbkode 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Sub Divisi"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin Chameleon.chameleonButton cmdcancel 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3960
      _ExtentX        =   6985
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
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3780
      TabIndex        =   4
      Top             =   2490
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2490
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
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
      Left            =   240
      TabIndex        =   1
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

Private Sub cmbkode_Click()
    OBJ.Open dsn
    SQL = carisql1 + " and nopo like '%" + cmbkode + "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
        
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
    
    grid.ColWidth(0) = 1600
    grid.ColWidth(1) = 1325
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 0
End Sub

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

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
        If namatabel = "Faktur" Then
            hasil1 = ""
        Else
            hasil1 = grid.TextMatrix(grid.Row, 1)
        End If
        If namatabel = "Supplier" Then
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Barang per Divisi" Then
            hasil2 = grid.TextMatrix(grid.Row, 5)
        Else
            hasil2 = ""
        End If
        
        Unload Me
    End If
End Sub

Private Sub showtran()
    If namatabel = "Purchase Order " Then
        Frame1.Visible = True
        cmbkode.Clear
        
        OBJ.Open dsn
        SQL = "select * from am_kode"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            
            RST.MoveNext
        Loop
        OBJ.Close
        
        grid.Top = 480
        grid.Height = 1935
    End If
    
    OBJ.Open dsn
    SQL = carisql1
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Transaction" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 1325
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Purchase Order" Or _
        namatabel = "Penerimaan Barang" Or _
        namatabel = "Pemakaian Barang" Or _
        namatabel = "Pengiriman Barang" Or _
        namatabel = "Mutasi" Or _
        namatabel = "Bayar Hutang" Or _
        namatabel = "Koreksi" Or _
        namatabel = "Purchase Order " Or _
        namatabel = "Produksi Harian" Then
            grid.ColWidth(0) = 1600
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Apply to..." Then
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = grid.ColWidth(0)
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Produk" Then
            grid.ColWidth(0) = 0
            grid.ColWidth(1) = 1000
            grid.ColWidth(2) = 2000
        ElseIf namatabel = "konversilevel" Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 3000
            grid.ColWidth(2) = 0
        ElseIf namatabel = "Barang Jadi." Then
            grid.ColWidth(0) = 800
            grid.ColWidth(1) = 3700
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 1000
        ElseIf namatabel = "Rencana Produksi" Then
            grid.ColWidth(0) = 1500
            grid.ColWidth(1) = 1200
            grid.ColWidth(2) = 1200
            grid.ColWidth(3) = 1200
        Else
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 0
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
            MsgBox "1"
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
    
    If namatabel = "Purchase Order" Or _
    namatabel = "Purchase Order " Or _
    namatabel = "Penerimaan Barang" Or _
    namatabel = "Produksi Harian" Or _
    namatabel = "Mutasi" Or _
    namatabel = "Pengiriman Barang" Or _
    namatabel = "Bayar Hutang" Or _
    namatabel = "Koreksi" Or _
    namatabel = "Pemakaian Barang" Then
        grid.ColWidth(0) = 1600
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Apply to..." Then
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = grid.ColWidth(0)
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Produk" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 2000
    ElseIf namatabel = "Produk." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2000
    ElseIf namatabel = "konversilevel" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 3000
        grid.ColWidth(2) = 0
    ElseIf namatabel = "Barang Jadi." Then
        grid.ColWidth(0) = 800
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Rencana Produksi" Then
        grid.ColWidth(0) = 1500
        grid.ColWidth(1) = 1200
        grid.ColWidth(2) = 1200
        grid.ColWidth(3) = 1200
    Else
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 0
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    End If
End Sub

Private Sub showtabel()
    OBJ.Open dsn
    If namatabel = "Item on PO" Then
        SQL = carisql1 + " order by a.lineitem"
    ElseIf namatabel = "Rencana Produksi" Then
        SQL = carisql1 + " group By KD_RCN,TGL1,TGL2 Order By KD_RCN desc"
    Else
        SQL = carisql1
    End If
    Set RST = OBJ.Execute(SQL)
  
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Stock Bahan Baku" Then
            grid.ColWidth(0) = 1200
            grid.ColWidth(1) = 1750
            grid.ColWidth(2) = 700
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Satuan Bahan Baku" Or namatabel = "Satuan Bahan Baku " Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 1000
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Item on PO" Then
            grid.ColWidth(1) = grid.ColWidth(0)
            grid.ColWidth(2) = 1500
        ElseIf namatabel = "Bahan Baku " Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 1000
            grid.ColWidth(2) = 1650
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Barang per Divisi" Then
            grid.ColWidth(0) = 0
            grid.ColWidth(1) = 2500
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 1000
            grid.ColWidth(4) = 1000
            grid.ColWidth(5) = 0
            grid.ColWidth(6) = 2000
        ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Then
            grid.ColWidth(0) = 2500
            grid.ColWidth(1) = 3700
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        ElseIf namatabel = "Produk" Then
            grid.ColWidth(0) = 0
            grid.ColWidth(1) = 1000
            grid.ColWidth(2) = 2000
        ElseIf namatabel = "Produk." Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 2000
        ElseIf namatabel = "konversilevel" Then
            grid.ColWidth(0) = 1000
            grid.ColWidth(1) = 3000
            grid.ColWidth(2) = 0
        ElseIf namatabel = "Barang Jadi." Then
            grid.ColWidth(0) = 800
            grid.ColWidth(1) = 3700
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 1000
        ElseIf namatabel = "Rencana Produksi" Then
            grid.ColWidth(0) = 1500
            grid.ColWidth(1) = 1200
            grid.ColWidth(2) = 1200
            grid.ColWidth(3) = 1200
        Else
            grid.ColWidth(1) = 2940
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
    If namatabel = "Barang per Divisi" Then
        grid.TextMatrix(0, 1) = "> " & grid.TextMatrix(0, 1)
        Label4 = Mid$(grid.TextMatrix(0, 1), 3)
        m_SortColumn = 1
        Label3 = 1
        grid.Col = 1
    Else
        grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
        Label4 = Mid$(grid.TextMatrix(0, 0), 3)
        m_SortColumn = 0
        Label3 = 0
        grid.Col = 0
    End If
    grid.Sort = flexSortStringAscending
        
    If namatabel = "Stock Bahan Baku" Then
        grid.ColWidth(0) = 1200
        grid.ColWidth(1) = 1750
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Satuan Bahan Baku" Or namatabel = "Satuan Bahan Baku " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Item on PO" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
    ElseIf namatabel = "Bahan Baku " Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 1650
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Barang per Divisi" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
        grid.ColWidth(4) = 1000
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 2000
    ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Then
        grid.ColWidth(0) = 2500
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Produk" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 2000
    ElseIf namatabel = "Produk." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2000
    ElseIf namatabel = "konversilevel" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 3000
        grid.ColWidth(2) = 0
    ElseIf namatabel = "Barang Jadi." Then
        grid.ColWidth(0) = 800
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Rencana Produksi" Then
        grid.ColWidth(0) = 1500
        grid.ColWidth(1) = 1200
        grid.ColWidth(2) = 1200
        grid.ColWidth(3) = 1200
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1500
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    m_SortColumn = -1
    m_SortAscending = -1
    
    If namatabel = "Penerimaan Barang" Or _
    namatabel = "Pemakaian Barang" Or _
    namatabel = "Produksi Harian" Or _
    namatabel = "Pengiriman Barang" Or _
    namatabel = "Mutasi" Or _
    namatabel = "Bayar Hutang" Or _
    namatabel = "Koreksi" Or _
    namatabel = "Hutang" Or _
    namatabel = "Apply to..." Or _
    namatabel = "Purchase Order" Or _
    namatabel = "Purchase Order " Then
        Label1.Visible = False
        txtsearch.Visible = False
        grid.Top = 120
        grid.Height = 2295
        Me.Caption = "Searching Transaksi " & namatabel
        showtran
        Exit Sub
    End If
    
    Me.Caption = "Searching Tabel " & namatabel
    If namatabel = "Barang per Divisi" Then
        Me.Width = 7545
        grid.Width = 7200
        cmdcancel.Width = 7200
    ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Or namatabel = "Rencana Produksi" Then
        Me.Width = 6945
        grid.Width = 6615
        cmdcancel.Width = 6615
    End If
    Adodc1.ConnectionString = dsn
    showdata
    
End Sub

Private Sub grid_Click()
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    Label3 = grid.MouseCol
    If grid.MouseRow > 0 Then
        hasil = grid.TextMatrix(grid.Row, 0)
        If namatabel = "Faktur" Then
            hasil1 = ""
        Else
            hasil1 = grid.TextMatrix(grid.Row, 1)
        End If
        If namatabel = "Apply to..." Or namatabel = "Supplier" Or namatabel = "Supplier " Then
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Barang per Divisi" Then
            hasil2 = grid.TextMatrix(grid.Row, 5)
        ElseIf namatabel = "Produk" Then
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Produk." Then
            hasil1 = grid.TextMatrix(grid.Row, 0)
            hasil2 = grid.TextMatrix(grid.Row, 1)
        ElseIf namatabel = "konversilevel" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Barang Jadi." Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
            hasil3 = grid.TextMatrix(grid.Row, 3)
        ElseIf namatabel = "Rencana Produksi" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
            hasil3 = grid.TextMatrix(grid.Row, 3)
        ElseIf namatabel = "Barang Jadi" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "WIP Jadi" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
            hasil3 = grid.TextMatrix(grid.Row, 3)
        ElseIf namatabel = "Packaging 1" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
            hasil3 = grid.TextMatrix(grid.Row, 3)
        ElseIf namatabel = "Satuan Bahan Baku" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
        Else
            hasil2 = ""
        End If
        
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
    OBJ.Open dsn
    If namatabel = "Satuan Bahan Baku " Then
        SQL = carisql1 + " and c." + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Item on PO" Then
        If Label4 = "kodebarang" Or Label4 = "kodesatuan" Then
            SQL = carisql1 + " and a." + Label4 + " like '" + txtsearch + "%' order by a.lineitem"
        Else
            SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%' order by a.lineitem"
        End If
    ElseIf namatabel = "Stock Bahan Baku" Then
        If Label4 = "kodebarang" Or Label4 = "namabarang" Then
            SQL = carisql1 + " where a." + Label4 + " like '" + txtsearch + "%'"
        Else
            SQL = carisql1
        End If
    ElseIf namatabel = "Barang per Divisi" Then
        If Label4 = "kodebarang" Or Label4 = "kodesatuan" Then SQL = carisql1 + " and a." + Label4 + " like '" + txtsearch + "%'"
        If Label4 = "namabarang" Then SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%'"
        If Label4 = "namasupp" Then SQL = carisql1 + " and c." + Label4 + " like '" + txtsearch + "%'"
        If Label4 = "lastupdate" Or Label4 = "price" Then SQL = carisql1
    ElseIf namatabel = "Supplier " Or namatabel = "Satuan" Or namatabel = "Company Account" Then
        SQL = carisql1 + " and b." + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Faktur" Or namatabel = "Cek/Giro" Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Produk" Then
        SQL = carisql1 + " and " & Label4 & " like '" & txtsearch & "%'"
    ElseIf namatabel = "Produk." Then
        SQL = carisql1 + " where " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "konversilevel" Then
        SQL = carisql1 + " and " & Label4 & " Like '" & txtsearch & "%'"
    ElseIf namatabel = "Barang " Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Rencana Produksi" Then
        SQL = carisql1 + " Where " + Label4 + " like '" + txtsearch + "%' Group By KD_RCN,TGL1,TGL2 Order By KD_RCN desc"
    Else
        SQL = carisql1 + " where " + Label4 + " like '" + txtsearch + "%'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Stock Bahan Baku" Then
        grid.ColWidth(0) = 1200
        grid.ColWidth(1) = 1750
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Satuan Bahan Baku" Or namatabel = "Satuan Bahan Baku " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Item on PO" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
    ElseIf namatabel = "Bahan Baku " Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 1650
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Barang per Divisi" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
        grid.ColWidth(4) = 1000
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 2000
    ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Then
        grid.ColWidth(0) = 2500
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Produk" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 2000
    ElseIf namatabel = "Produk." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2000
    ElseIf namatabel = "konversilevel" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 3000
        grid.ColWidth(2) = 0
    ElseIf namatabel = "Barang Jadi." Then
        grid.ColWidth(0) = 800
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Rencana Produksi" Then
        grid.ColWidth(0) = 1500
        grid.ColWidth(1) = 1200
        grid.ColWidth(2) = 1200
        grid.ColWidth(3) = 1200
    ElseIf namatabel = "Bahan Baku" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 4500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        cmdcancel.Width = grid.Width
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        cmdcancel.Width = grid.Width
        'grid.ColWidth(1) = 2940
        'grid.ColWidth(2) = 0
        'grid.ColWidth(3) = 0
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

    If namatabel = "Stock Bahan Baku" Then
        grid.ColWidth(0) = 1200
        grid.ColWidth(1) = 1750
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Satuan Bahan Baku" Or namatabel = "Satuan Bahan Baku " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Item on PO" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
    ElseIf namatabel = "Bahan Baku " Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 1650
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Barang per Divisi" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
        grid.ColWidth(4) = 1000
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 2000
    ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Then
        grid.ColWidth(0) = 2500
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Produk" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 2000
    ElseIf namatabel = "Produk." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2000
    ElseIf namatabel = "konversilevel" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 3000
        grid.ColWidth(2) = 0
    ElseIf namatabel = "Barang Jadi." Then
        grid.ColWidth(0) = 800
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Rencana Produksi" Then
        grid.ColWidth(0) = 1500
        grid.ColWidth(1) = 1200
        grid.ColWidth(2) = 1200
        grid.ColWidth(3) = 1200
    ElseIf namatabel = "Bahan Baku" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 4500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        cmdcancel.Width = grid.Width
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        'grid.ColWidth(1) = 2940
        'grid.ColWidth(2) = 0
        'grid.ColWidth(3) = 0
    End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
    If KeyAscii = 13 And grid.Rows = 2 Then
        hasil = grid.TextMatrix(1, 0)
        If namatabel = "Faktur" Then
            hasil1 = ""
        Else
            hasil1 = grid.TextMatrix(1, 1)
        End If
        If namatabel = "Apply to..." Or namatabel = "Supplier" Or namatabel = "Supplier " Then
            hasil2 = grid.TextMatrix(1, 2)
        ElseIf namatabel = "Barang per Divisi" Then
            hasil2 = grid.TextMatrix(1, 5)
        Else
            hasil2 = ""
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If Label3 = "" Then KeyAscii = 0
End Sub

Private Sub showdata()
    If namatabel = "Item on PO" Then
        SQL = carisql1 + " order by a.lineitem"
    ElseIf namatabel = "Rencana Produksi" Then
        SQL = carisql1 + " group By KD_RCN,TGL1,TGL2 Order By KD_RCN desc"
    Else
        SQL = carisql1
    End If
    Adodc1.RecordSource = SQL
    Set grid.DataSource = Adodc1

    Label2 = grid.Rows - 1 & " Records"
    If namatabel = "Barang per Divisi" Then
        grid.TextMatrix(0, 1) = "> " & grid.TextMatrix(0, 1)
        Label4 = Mid$(grid.TextMatrix(0, 1), 3)
        m_SortColumn = 1
        Label3 = 1
        grid.Col = 1
    Else
        grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
        Label4 = Mid$(grid.TextMatrix(0, 0), 3)
        m_SortColumn = 0
        Label3 = 0
        grid.Col = 0
    End If
    grid.Sort = flexSortStringAscending
        
    If namatabel = "Stock Bahan Baku" Then
        grid.ColWidth(0) = 1200
        grid.ColWidth(1) = 1750
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Satuan Bahan Baku" Or namatabel = "Satuan Bahan Baku " Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Item on PO" Then
        grid.ColWidth(1) = grid.ColWidth(0)
        grid.ColWidth(2) = 1500
    ElseIf namatabel = "Bahan Baku " Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 1650
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Barang per Divisi" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 2500
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
        grid.ColWidth(4) = 1000
        grid.ColWidth(5) = 0
        grid.ColWidth(6) = 2000
    ElseIf namatabel = "Supplier" Or namatabel = "Supplier " Then
        grid.ColWidth(0) = 2500
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Produk" Then
        grid.ColWidth(0) = 0
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 2000
    ElseIf namatabel = "Produk." Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 2000
    ElseIf namatabel = "konversilevel" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 3000
        grid.ColWidth(2) = 0
    ElseIf namatabel = "Barang Jadi." Then
        grid.ColWidth(0) = 800
        grid.ColWidth(1) = 3700
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1000
    ElseIf namatabel = "Rencana Produksi" Then
        grid.ColWidth(0) = 1500
        grid.ColWidth(1) = 1200
        grid.ColWidth(2) = 1200
        grid.ColWidth(3) = 1200
    ElseIf namatabel = "Satuan Bahan Baku" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1000
        grid.ColWidth(3) = 0
    ElseIf namatabel = "Bahan Baku" Then
        grid.ColWidth(0) = 1000
        grid.ColWidth(1) = 4500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        cmdcancel.Width = grid.Width
    Else
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 1500
        Me.Width = 6615
        grid.Width = Me.Width - 500
        cmdcancel.Width = grid.Width
    End If
End Sub
