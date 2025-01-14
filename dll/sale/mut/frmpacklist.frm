VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpacklist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kontrol Lot Pengambilan Kemasan"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker Date2 
      Height          =   315
      Left            =   9120
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   134873089
      CurrentDate     =   42039
   End
   Begin VB.TextBox txtnolot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   6060
      ItemData        =   "frmpacklist.frx":0000
      Left            =   120
      List            =   "frmpacklist.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Klik Nomor Lot untuk confirm"
      Top             =   480
      Width           =   2655
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   12960
      TabIndex        =   0
      Top             =   3480
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
      MICON           =   "frmpacklist.frx":0004
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
      Height          =   3015
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483642
      ScrollBars      =   0
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
      _Band(0).Cols   =   4
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   134873089
      CurrentDate     =   42039
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   3015
      Left            =   8040
      TabIndex        =   11
      Top             =   840
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483642
      ScrollBars      =   0
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
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
      Height          =   2175
      Left            =   8040
      TabIndex        =   12
      Top             =   4320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483642
      ScrollBars      =   2
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
   Begin Chameleon.chameleonButton cmdrefresh 
      Height          =   375
      Left            =   12960
      TabIndex        =   13
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Refresh"
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
      MICON           =   "frmpacklist.frx":031E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   2175
      Left            =   2880
      TabIndex        =   8
      Top             =   4320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483642
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
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pemakaian di SOP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selisih Pemakaian Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jumlah Pengambilan Kemasan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   480
      Width           =   10000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "No. LOT "
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
      Left            =   3000
      TabIndex        =   7
      Top             =   150
      Width           =   1335
   End
   Begin VB.Label lblTgl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   10320
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pending LOT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmpacklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdrefresh_Click()
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE flag = '1'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
    hapusgrid
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 0) = "X"
    grid.TextMatrix(0, 1) = "Kode"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Request"
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    
    grid1.TextMatrix(0, 0) = "Kode"
    grid1.TextMatrix(0, 1) = "Nama Barang"
    grid1.TextMatrix(0, 2) = "Rusak/Hilang/Retur"
    grid1.ColWidth(0) = 0
    grid1.ColWidth(1) = 3000
    grid1.ColWidth(2) = 1800
    
    grid2.TextMatrix(0, 0) = "Kode"
    grid2.TextMatrix(0, 1) = "Nama Barang"
    grid2.TextMatrix(0, 2) = "Pemakaian"
    grid2.ColWidth(0) = 1000
    grid2.ColWidth(1) = 3000
    grid2.ColWidth(2) = 1100
    
    grid3.TextMatrix(0, 0) = "Kode"
    grid3.TextMatrix(0, 1) = "Nama Barang"
    grid3.TextMatrix(0, 2) = "Gudang"
    grid3.TextMatrix(0, 3) = "SOP"
    grid3.TextMatrix(0, 4) = "Selisih"
    grid3.ColWidth(0) = 0
    grid3.ColWidth(1) = 3000
    grid3.ColWidth(2) = 750
    grid3.ColWidth(3) = 750
    grid3.ColWidth(4) = 700
    
    Date2 = Date
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE flag = '1'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub List1_Click()
    If List1.text = "" Then Exit Sub
    txtnolot = List1.text
    hapusgrid
    
    OBJ.Open dsn
    'Pengambilan Kemasan di gudang
    SQL = "Select a.kodebarang,b.NamaBarang,a.tgl,SUM(a.qty_confirmed)'Request' From am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang"
    SQL = SQL + " Where a.nolot = '" & txtnolot & "' and a.status ='1' and a.flag = '1'"
    SQL = SQL + " group by a.kodebarang,b.NamaBarang,a.tgl order by a.kodebarang Asc"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
    
        lblTgl = "Tanggal : " & Format(RST!tgl, "dd MMMM yyyy")
        date1 = RST!tgl
        Do Until RST.EOF
            grid.Col = 0
            'Set grid.CellPicture = uncheck
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
            If RST!request = "" Or IsNull(RST!request) Or RST!request = "0" Then
                grid.TextMatrix(grid.Row, 3) = ""
            Else
                grid.TextMatrix(grid.Row, 3) = Format(RST!request, "###,###,##0") * -1
            End If
    
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Request packaging data not Found.", vbExclamation, AppName
    End If
    
    'Retur/Rusak/Hilang Kemasan kegudang
    SQL = "Select a.kodebarang,b.NamaBarang,sum(a.qty_confirmed)'Return' From am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang=b.kodebarang"
    SQL = SQL + " Where a.nolot = '" & txtnolot & "' and a.status in('5','6','7') and a.flag = '1'"
    SQL = SQL + " group by a.kodebarang,b.NamaBarang order by a.kodebarang asc"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        Do Until RST.EOF
            grid1.Col = 0
            grid1.TextMatrix(grid1.Row, 0) = RST!kodebarang
            grid1.TextMatrix(grid1.Row, 1) = RST!NamaBarang
            If RST!return = "" Or IsNull(RST!return) Or RST!return = "0" Then
                grid1.TextMatrix(grid1.Row, 2) = ""
            Else
                grid1.TextMatrix(grid1.Row, 2) = Format(RST!return, "###,###,##0") * -1
            End If
            grid1.Rows = grid1.Rows + 1
            grid1.Row = grid1.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Data Return not Found.", vbExclamation, AppName
    End If
    
    'Pemakaian Kemasan di SOP
    SQL = "Select a.kode_bahan,b.NamaBarang,SUM(a.qty_bahan)'Used' From list_produksi_kemasan a"
    SQL = SQL + " inner join am_apitemmst b on a.kode_bahan = b.KodeBarang"
    SQL = SQL + " Where a.nolot = '" & txtnolot & "' group by a.kode_bahan,b.NamaBarang order by a.kode_bahan asc"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        Do Until RST.EOF
            grid2.Col = 0
            grid2.TextMatrix(grid2.Row, 0) = RST!kode_bahan
            grid2.TextMatrix(grid2.Row, 1) = RST!NamaBarang
            grid2.TextMatrix(grid2.Row, 2) = Format(RST!used, "###,###,##0")
    
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Packaging used data not Found.", vbExclamation, AppName
    End If
    
    'Selisih Permintaan - Pemakaian di SOP
    SQL = "Select a.kodebarang,b.NamaBarang,SUM(a.qty_confirmed * -1)'add' From am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang"
    SQL = SQL + " Where a.nolot = '" & txtnolot & "'"
    SQL = SQL + " group by a.kodebarang,b.NamaBarang order by a.kodebarang"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        grid2.Row = 1
        Do Until RST.EOF
            grid3.Col = 0
            grid3.TextMatrix(grid3.Row, 0) = RST!kodebarang
            grid3.TextMatrix(grid3.Row, 1) = RST!NamaBarang
            If RST!Add = "" Or IsNull(RST!Add) Or RST!Add = "0" Then
                grid3.TextMatrix(grid3.Row, 2) = ""
            Else
                grid3.TextMatrix(grid3.Row, 2) = Format(RST!Add, "###,###,##0")
                grid3.TextMatrix(grid3.Row, 2) = Format(grid3.TextMatrix(grid3.Row, 2), "###,###,##0")
            End If
            grid3.TextMatrix(grid3.Row, 3) = Format(grid2.TextMatrix(grid2.Row, 2), "###,###,##0")
            If grid3.TextMatrix(grid3.Row, 3) = "" Then
                grid3.TextMatrix(grid3.Row, 4) = grid3.TextMatrix(grid3.Row, 2)
            Else
                'grid3.TextMatrix(grid3.Row, 4) = grid3.TextMatrix(grid3.Row, 2) - grid3.TextMatrix(grid3.Row, 3)
                If grid3.TextMatrix(grid3.Row, 2) = "" Then
                    grid3.TextMatrix(grid3.Row, 4) = 0 - grid3.TextMatrix(grid3.Row, 3)
                Else
                    grid3.TextMatrix(grid3.Row, 4) = grid3.TextMatrix(grid3.Row, 2) - grid3.TextMatrix(grid3.Row, 3)
                End If
            End If
    
            grid3.Rows = grid3.Rows + 1
            grid3.Row = grid3.Row + 1
            grid2.Rows = grid2.Rows + 1
            grid2.Row = grid2.Row + 1
            RST.MoveNext
        Loop
        grid2.Rows = grid2.Row + 1
        OBJ.Close
        Call compare
    Else
        'MsgBox "Packaging used data not Found.", vbExclamation, AppName
        txtnolot = ""
        OBJ.Close
    End If
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.Col = 0
        'Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 0) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    grid1.ColWidth(0) = 0
    grid1.ColWidth(1) = 3000
    grid1.ColWidth(2) = 1800
    
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 0) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    grid2.ColWidth(0) = 1000
    grid2.ColWidth(1) = 3000
    grid2.ColWidth(2) = 1100
    
    grid3.Row = 1
    Do While True
        If grid3.TextMatrix(grid3.Row, 0) = "" Then Exit Do
        grid3.TextMatrix(grid3.Row, 0) = ""
        grid3.TextMatrix(grid3.Row, 1) = ""
        grid3.TextMatrix(grid3.Row, 2) = ""
        grid3.TextMatrix(grid3.Row, 3) = ""
        grid3.TextMatrix(grid3.Row, 4) = ""
        grid3.Row = grid3.Row + 1
    Loop
    grid3.Rows = 2
    grid3.ColWidth(0) = 0
    grid3.ColWidth(1) = 3000
    grid3.ColWidth(2) = 750
    grid3.ColWidth(3) = 750
    grid3.ColWidth(4) = 700
End Sub

Function tanggalconfirm()
    tanggalconfirm = Month(Date2) & "/" & Day(Date2) & "/" & Year(Date2)
End Function

Private Sub compare()
On Error Resume Next
Dim tg3 As Double
'Total pemakaian - total pengambilan
    grid3.Row = 1
    tg3 = 0
    Do While True
        DoEvents
        If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
        If grid3.TextMatrix(grid3.Row, 4) < 0 Then
            tg3 = (CDbl(Format(grid3.TextMatrix(grid3.Row, 4), "general number") * -1) + CDbl(tg3))
        Else
            tg3 = CDbl(Format(grid3.TextMatrix(grid3.Row, 4), "general number") + CDbl(tg3))
        End If
        grid3.Row = grid3.Row + 1
    Loop
        tg3 = Format(tg3, "##,###,##0.00")
        If tg3 = 0 Then
            MsgBox "Close SOP", vbInformation
            OBJ.Open dsn
            SQL = "Insert into am_gudang_perclose(nolot,tgl)"
            SQL = SQL + " Values('" & txtnolot & "',convert(datetime,'" & tanggalconfirm & "'))"
            Set RST = OBJ.Execute(SQL)
            
            'Close Permintaan
            SQL = "update am_gudang_permintaan set flag = '2' Where nolot = '" & txtnolot & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            cmdrefresh_Click
        End If
End Sub
