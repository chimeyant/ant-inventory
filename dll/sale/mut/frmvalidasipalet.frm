VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmvalidasipalet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validasi Data"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   10200
      TabIndex        =   0
      Top             =   7440
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
      MICON           =   "frmvalidasipalet.frx":0000
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
      Height          =   6420
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11324
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   -2147483632
      BackColorBkg    =   8421504
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
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   10935
      _Version        =   851970
      _ExtentX        =   19288
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
   Begin Chameleon.chameleonButton btnshow 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show"
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
      MICON           =   "frmvalidasipalet.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton btnval 
      Height          =   375
      Left            =   10200
      TabIndex        =   6
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Validasi"
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
      MICON           =   "frmvalidasipalet.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdreturn 
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   7440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Return pending"
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
      MICON           =   "frmvalidasipalet.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdperiksa 
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View data Null"
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
      MICON           =   "frmvalidasipalet.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblvalid 
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
      TabIndex        =   5
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Label lblrow 
      Caption         =   "0 Palet."
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
      TabIndex        =   4
      Top             =   7320
      Width           =   3735
   End
End
Attribute VB_Name = "frmvalidasipalet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL As String
Dim SQL1 As String

Dim baris, i, jumlah As Integer

Private Sub btnshow_Click()
    opendatahpp
End Sub

Private Sub Validasi()
    OBJ.Open dsn
    SQL = "Select a.noref,a.tanggal,b.kodebarang,b.kg,isnull(c.pack,0)'pack',e.thppbahan,e.perkilo,isnull(g.thpppack,0)'thpppack',g.thasil,(a.qty_bahan*b.kg)'hasil',"
    SQL = SQL + " e.thppbahan +isnull(g.thpppack,0)'brutto',(g.thasil*e.perkilo)+isnull(g.thpppack,0)'tjadi',"
    SQL = SQL + " (e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0))'loss',"
    SQL = SQL + " (((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil)'lossperkg',"
    SQL = SQL + " (isnull(c.pack,0)/(a.qty_bahan*b.kg))'packperkg',"
    SQL = SQL + " (e.perkilo + (isnull(c.pack,0)/(a.qty_bahan*b.kg))+(((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil))'hppperkg'"
    SQL = SQL + " From list_produksi_hasil a"
    SQL = SQL + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "',3,1)='A' then kg1"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "',3,1)='B' then kg2"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='C' then kg3"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='D' then kg4"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='E' then kg5"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='F' then kg6"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='G' then kg7"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='H' then kg8"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='J' then kg9"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='K' then kg10"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='L' then kg11"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
    SQL = SQL + " b on a.kode_bahan = b.kodebarang and a.kode_satuan = b.kodesatuan"
    SQL = SQL + " left join (select noref,isnull(SUM(hpp),0)'pack' from list_produksi_kemasan where nolot = '" & grid.TextMatrix(grid.Row, 2) & "' group by noref) c on a.noref = c.noref"
    SQL = SQL + " left join list_produksi_child d on a.nolot = d.nolot"
    SQL = SQL + " inner join (Select x.nolot,y.noref,SUM(x.hpp)'thppbahan',SUM(x.hpp)/SUM(x.qty_bahan)'perkilo'"
    SQL = SQL + " from list_produksi_child x left join list_produksi_hasil y on x.nolot = y.nolot where x.nolot ='" & grid.TextMatrix(grid.Row, 2) & "'"
    SQL = SQL + " group by x.nolot,y.noref) e on a.noref = e.noref"
    SQL = SQL + " left join (Select m.nolot,o.thpppack,SUM(m.qty_bahan * n.kg)'thasil' From list_produksi_hasil m"
    SQL = SQL + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "',3,1)='A' then kg1"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "',3,1)='B' then kg2"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='C' then kg3"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='D' then kg4"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='E' then kg5"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='F' then kg6"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='G' then kg7"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='H' then kg8"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='J' then kg9"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='K' then kg10"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='L' then kg11"
    SQL = SQL + " when SUBSTRING('" & grid.TextMatrix(grid.Row, 2) & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
    SQL = SQL + " n on m.kode_bahan = n.kodebarang and m.kode_satuan = n.kodesatuan"
    SQL = SQL + " left join (Select nolot,isnull(SUM(hpp),0)'thpppack' from list_produksi_kemasan Where nolot = '" & grid.TextMatrix(grid.Row, 2) & "' group by nolot)"
    SQL = SQL + " o on m.nolot=o.nolot"
    SQL = SQL + " Where m.nolot = '" & grid.TextMatrix(grid.Row, 2) & "' and m.proses_ke = '2' and n.tahun = '20' + LEFT('" & grid.TextMatrix(grid.Row, 2) & "',2) group by m.nolot,o.thpppack) g on a.nolot = g.nolot"
    SQL = SQL + " Where a.nolot = '" & grid.TextMatrix(grid.Row, 2) & "' and b.tahun = '20' + LEFT('" & grid.TextMatrix(grid.Row, 2) & "',2)and a.noref = '" & grid.TextMatrix(grid.Row, 3) & "'"
    SQL = SQL + " group by a.noref,a.tanggal,b.kodebarang,b.kg,c.pack,e.thppbahan,e.perkilo,g.thpppack,g.thasil,a.qty_bahan order by a.noref asc"
    Set RST = OBJ.Execute(SQL)
        
    If grid.TextMatrix(grid.Row, 8) = Format(RST!hppperkg, "##,###,##0.000") Then
        grid.TextMatrix(grid.Row, 9) = Format(RST!hppperkg, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 10) = "OKE"
        'update flag = '1' validasi hpp oke
        'OBJ1.Open dsn
        'SQL1 = "Update list_hpp_produksi set flag='2' Where palet='" & grid.TextMatrix(grid.Row, 3) & "'"
        'Set RST1 = OBJ1.Execute(SQL1)
        
        'SQL1 = "Update list_produksi_master set flagprint='4' Where nolot='" & grid.TextMatrix(grid.Row, 2) & "'"
        'Set RST1 = OBJ1.Execute(SQL1)
        'OBJ1.Close
        i = i + 1
        lblvalid = "Palet ready: " & i
    ElseIf grid.TextMatrix(grid.Row, 8) >= Format(RST!hppperkg + 0.001, "##,###,##0.000") Then
        grid.TextMatrix(grid.Row, 9) = Format(RST!hppperkg, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 10) = "OKE"
        i = i + 1
        lblvalid = "Palet ready: " & i
    ElseIf grid.TextMatrix(grid.Row, 8) <> Format(RST!hppperkg, "##,###,##0.000") Then
        grid.TextMatrix(grid.Row, 9) = Format(RST!hppperkg, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 10) = "Pending"
        setAlternatingGridYelow grid.Row
    End If

    OBJ.Close
End Sub

Private Sub btnval_Click()
    i = 0
    Pg.Max = baris
    Pg.Value = 0
    Pg.Visible = True
    Screen.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        Validasi
        DoEvents
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
    Loop
    Pg.Value = 0
    lblvalid = "Palet ready: " & i
    If baris <> i Then cmdreturn.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub opendatahpp()
    hapusgrid
    OBJ.Open dsn
    
    SQL = "Select COUNT(y.nolot)'jml'"
    SQL = SQL + " From(Select a.tanggal,a.nolot,a.palet,a.kodebarang,b.NamaBarang,c.NamaSatuan,a.hppperkg From list_hpp_produksi a "
    SQL = SQL + " inner join am_itemdtl b on a.kodebarang = b.KodeBarang"
    SQL = SQL + " inner join am_unit c on b.KodeSatuan = c.KodeSatuan Where a.flag = '1' and b.Level_ = '0') y"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select a.tanggal,a.nolot,a.palet,a.kodebarang,b.NamaBarang,(a.kgperpalet/a.kg)'qty',c.NamaSatuan,a.hppperkg From list_hpp_produksi a "
    SQL = SQL + " inner join am_itemdtl b on a.kodebarang = b.KodeBarang"
    SQL = SQL + " inner join am_unit c on b.KodeSatuan = c.KodeSatuan"
    SQL = SQL + " Where a.flag = '1' and b.Level_ = '0' order by a.nolot desc"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!tanggal
        grid.TextMatrix(grid.Row, 2) = RST!nolot
        grid.TextMatrix(grid.Row, 3) = RST!palet
        grid.TextMatrix(grid.Row, 4) = RST!kodebarang
        grid.TextMatrix(grid.Row, 5) = RST!NamaBarang
        grid.TextMatrix(grid.Row, 6) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 7) = RST!namasatuan
        grid.TextMatrix(grid.Row, 8) = Format(RST!hppperkg, "##,###,##0.000") 'hpp hasil scan diproduksi
        grid.TextMatrix(grid.Row, 9) = ""   'validasi hpp
        grid.TextMatrix(grid.Row, 10) = ""   'status (pending / OKE)
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        lblrow = Pg.Value & " Palet"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    baris = grid.Row - 1
    lblrow = baris & " Palet"
End Sub

Private Sub cmdperiksa_Click()
    frmdatanull.Show vbModal
End Sub

Private Sub cmdreturn_Click()
    Me.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        If grid.TextMatrix(grid.Row, 10) = "Pending" Then
            OBJ.Open dsn
            SQL = "Update list_produksi_master set flagprint= '3' Where nolot='" & grid.TextMatrix(grid.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "Update list_hpp_produksi set flag= '1' Where nolot='" & grid.TextMatrix(grid.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
        grid.Row = grid.Row + 1
    Loop
    Me.MousePointer = vbDefault
    cmdreturn.Enabled = False
    btnshow_Click
End Sub

Private Sub Form_Load()
    grid.Cols = 11
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "Nolot"
    grid.TextMatrix(0, 3) = "Palet"
    grid.TextMatrix(0, 4) = "Kode"
    grid.TextMatrix(0, 5) = "Item"
    grid.TextMatrix(0, 6) = "Qty"
    grid.TextMatrix(0, 7) = "Satuan"
    grid.TextMatrix(0, 8) = "Hpp"
    grid.TextMatrix(0, 9) = "Validasi"
    grid.TextMatrix(0, 10) = "Status"
    
    grid.ColWidth(0) = 600
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1600
    grid.ColWidth(3) = 1800
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 2500
    grid.ColWidth(6) = 800
    grid.ColWidth(7) = 800
    If nmuser = "Creator" Then
        grid.ColWidth(8) = 1500
        grid.ColWidth(9) = 1500
    Else
        grid.ColWidth(8) = 0
        grid.ColWidth(9) = 0
    End If
    grid.ColWidth(10) = 800
    grid.ColAlignment(0) = flexAlignLeftCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignLeftCenter
    grid.ColAlignment(6) = flexAlignRightCenter
    grid.ColAlignmentFixed(5) = flexAlignCenterCenter
    grid.ColAlignmentFixed(10) = flexAlignCenterCenter
    If nmuser = "Creator" Then cmdperiksa.Visible = True
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
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
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.TextMatrix(grid.Row, 10) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Function setAlternatingGridYelow(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To 10
        grid.Col = j
        grid.CellBackColor = vbYellow
    Next
End Function
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As Control
  
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSHFlexGrid Then
          If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        End If
    Next ctl
End Sub

Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function

