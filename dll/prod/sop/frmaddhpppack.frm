VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmaddhpppack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Hpp Packaging dan Hpp Kg Barang Jadi"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl Tab 
      Height          =   6015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9735
      _Version        =   851970
      _ExtentX        =   17171
      _ExtentY        =   10610
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   10
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Edit Hpp Packaging"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "grid"
      Item(0).Control(1)=   "btnupdate"
      Item(0).Control(2)=   "lblrow"
      Item(0).Control(3)=   "date1"
      Item(0).Control(4)=   "date2"
      Item(0).Control(5)=   "btnproses"
      Item(0).Control(6)=   "btnshow"
      Item(0).Control(7)=   "Label1"
      Item(0).Control(8)=   "Label2"
      Item(1).Caption =   "Edit Hpp per Kg"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "grid2"
      Item(1).Control(1)=   "btnshow2"
      Item(1).Control(2)=   "btnproses2"
      Item(1).Control(3)=   "lblrow2"
      Item(1).Control(4)=   "btnupdate2"
      Item(1).Control(5)=   "lblsinc"
      Item(1).Control(6)=   "Label4"
      Item(1).Control(7)=   "Label5"
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   4140
         Left            =   -69880
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7303
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
      Begin Chameleon.chameleonButton btnupdate 
         Height          =   375
         Left            =   -61240
         TabIndex        =   4
         Top             =   5400
         Visible         =   0   'False
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
         MICON           =   "frmaddhpppack.frx":0000
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
         Height          =   4140
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7303
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
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   -68920
         TabIndex        =   7
         Top             =   720
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
         Format          =   135135233
         CurrentDate     =   42039
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   315
         Left            =   -66640
         TabIndex        =   8
         Top             =   720
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
         Format          =   135135233
         CurrentDate     =   42039
      End
      Begin Chameleon.chameleonButton btnproses 
         Height          =   375
         Left            =   -61240
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses"
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
         MICON           =   "frmaddhpppack.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton btnshow 
         Height          =   375
         Left            =   -65080
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
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
         MICON           =   "frmaddhpppack.frx":0634
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton btnshow2 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
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
         MICON           =   "frmaddhpppack.frx":094E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton btnproses2 
         Height          =   375
         Left            =   8760
         TabIndex        =   14
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses"
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
         MICON           =   "frmaddhpppack.frx":0C68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton btnupdate2 
         Height          =   375
         Left            =   8760
         TabIndex        =   16
         Top             =   5520
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
         MICON           =   "frmaddhpppack.frx":0F82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hpp Produksi"
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
         Left            =   4400
         TabIndex        =   20
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hpp Gudang"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label lblsinc 
         BackStyle       =   0  'Transparent
         Caption         =   "Data hpp (per kilo) gudang barang jadi yang tidak sinkron: 0"
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
         Left            =   240
         TabIndex        =   17
         Top             =   5640
         Width           =   7695
      End
      Begin VB.Label lblrow2 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Lot"
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
         Left            =   240
         TabIndex        =   15
         Top             =   5400
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   -69880
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   -67360
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblrow 
         BackStyle       =   0  'Transparent
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
         Left            =   -69760
         TabIndex        =   5
         Top             =   5400
         Visible         =   0   'False
         Width           =   4335
      End
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   9735
      _Version        =   851970
      _ExtentX        =   17171
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
   Begin Chameleon.chameleonButton btnclose 
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   6360
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
      MICON           =   "frmaddhpppack.frx":129C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "* Data tidak sinkron disebabkan karena ada perubahan data SPK (SOP) yang sudah selesai/close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6480
      Width           =   8535
   End
End
Attribute VB_Name = "frmaddhpppack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim baris, baris2, jumlah As Integer

Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub opendata()
    hapusgrid
    OBJ.Open dsn
    
    SQL = "Select COUNT(palet)'jml' From am_stok Where tanggal >='" & Format(date1, "yyyy/MM/dd") & "'"
    SQL = SQL + " and tanggal <= '" & Format(date2, "yyyy/MM/dd") & "' and hpp_totpack is null"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select tanggal,nolot,palet,namabarang,hpp,hpp_totpack From am_stok"
    SQL = SQL + " Where tanggal >='" & Format(date1, "yyyy/MM/dd") & "'"
    SQL = SQL + " and tanggal <= '" & Format(date2, "yyyy/MM/dd") & "' and hpp_totpack is null"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!tanggal
        grid.TextMatrix(grid.Row, 2) = RST!nolot
        grid.TextMatrix(grid.Row, 3) = RST!palet
        grid.TextMatrix(grid.Row, 4) = RST!namabarang
        grid.TextMatrix(grid.Row, 5) = Format(RST!hpp, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 6) = Format(RST!hpp_totpack, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 7) = "0"
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        lblrow = Pg.Value & " Baris"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    baris = grid.Row - 1
    lblrow = baris & " Baris"
End Sub

Private Sub opendataperkg()
    hapusgrid2
    OBJ.Open dsn
    
    SQL = "Select COUNT(y.nolot)'jml'"
    SQL = SQL + " From (Select a.nolot,a.hpp,SUM(a.qtyin * a.kg)'totalhasil',SUM(b.hpp)'hppbahan',SUM(b.hpp)/SUM(b.qty_bahan)'perkg'"
    SQL = SQL + " From am_stok a inner join list_produksi_child b"
    SQL = SQL + " on a.nolot = b.nolot group by a.nolot,a.hpp) y"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select a.nolot,a.hpp,SUM(a.qtyin * a.kg)'totalhasil',SUM(b.hpp)'hppbahan',SUM(b.hpp)/SUM(b.qty_bahan)'perkg'"
    SQL = SQL + " From am_stok a inner join list_produksi_child b "
    SQL = SQL + " on a.nolot = b.nolot group by a.nolot,a.hpp order by a.nolot desc"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid2.TextMatrix(grid2.Row, 0) = grid2.Row
        grid2.TextMatrix(grid2.Row, 1) = RST!nolot
        grid2.TextMatrix(grid2.Row, 2) = Format(RST!hpp, "##,###,##0.00")   'hpp gudang
        grid2.TextMatrix(grid2.Row, 3) = Format(RST!perkg, "##,###,##0.00") 'hpp produksi
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        Pg.Value = Pg.Value + 1
        lblrow2 = Pg.Value & " Lot"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    baris2 = grid2.Row - 1
    lblrow2 = baris2 & " Lot"
End Sub

Private Sub btnproses_Click()
    Pg.Max = baris
    Pg.Value = 0
    Pg.Visible = True
    
    OBJ.Open dsn
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        'cek hpp bahan jika kosong akan diisi
        If grid.TextMatrix(grid.Row, 5) = "0.00" Then
            SQL = "Select nolot,SUM(hpp)/SUM(qty_bahan)'perkg' From list_produksi_child Where nolot = '" & grid.TextMatrix(grid.Row, 2) & "' group by nolot"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 5) = Format(RST!perkg, "##,###,##0.00")
                grid.TextMatrix(grid.Row, 7) = "1"
            End If
        End If
        
        'cek hpp kemasan
        SQL = "Select noref,SUM(hpp)'hpp_totpack' From list_produksi_kemasan Where noref = '" & grid.TextMatrix(grid.Row, 3) & "' Group By noref"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            grid.TextMatrix(grid.Row, 6) = "0.00"
        Else
            grid.TextMatrix(grid.Row, 6) = Format(RST!hpp_totpack, "##,###,##0.00")
        End If
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
    Loop
    OBJ.Close
    Pg.Value = 0
End Sub

Private Sub btnproses2_Click()
    Dim i As Integer
    i = 0
    Pg.Max = baris2
    Pg.Value = 0
    Pg.Visible = True
    
    OBJ.Open dsn
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        If grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row, 3) Then
            grid2.TextMatrix(grid2.Row, 4) = "0"
        Else
            grid2.TextMatrix(grid2.Row, 4) = "1"
            i = i + 1
        End If
        grid2.Row = grid2.Row + 1
        Pg.Value = Pg.Value + 1
    Loop
    OBJ.Close
    Pg.Value = 0
    lblsinc = "Data hpp (per kilo) gudang barang jadi yang tidak sinkron: " & i
End Sub

Private Sub btnshow_Click()
    Call opendata
End Sub

Private Sub btnshow2_Click()
    Call opendataperkg
End Sub

Private Sub btnupdate_Click()
    If grid.TextMatrix(1, 1) = "" Then Exit Sub
    If grid.TextMatrix(1, 6) = "" Then
        MsgBox "The data has not been processed" & vbLf & "Click the process button first", vbCritical, AppName
        Exit Sub
    End If
    If MsgBox("Are you sure you want to update this data", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    Pg.Max = baris
    Pg.Value = 0
    OBJ.Open dsn
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "Update am_stok set hpp_totpack='" & grid.TextMatrix(grid.Row, 6) & "'"
        SQL = SQL + " Where palet='" & grid.TextMatrix(grid.Row, 3) & "'"
        Set RST = OBJ.Execute(SQL)
        
        'update hpp bahan bernilai 0
        If grid.TextMatrix(grid.Row, 7) = "1" Then
            SQL = "Update am_stok set hpp='" & grid.TextMatrix(grid.Row, 5) & "'"
            SQL = SQL + " Where palet='" & grid.TextMatrix(grid.Row, 3) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        DoEvents
    Loop
    OBJ.Close
    Pg.Value = 0
    MsgBox "Data is successfuly updated", vbInformation, AppName
    date1 = Date
    date2 = Date
    hapusgrid
    lblrow = "0 Baris"
End Sub

Private Sub btnupdate2_Click()
    If grid2.TextMatrix(1, 1) = "" Then Exit Sub
    If grid2.TextMatrix(1, 4) = "" Then
        MsgBox "The data has not been processed" & vbLf & "Click the process button first", vbCritical, AppName
        Exit Sub
    End If
    If MsgBox("Are you sure you want to update this data", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    Pg.Max = baris2
    Pg.Value = 0
    OBJ.Open dsn
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        'update hpp bahan bernilai 0
        If grid2.TextMatrix(grid2.Row, 4) = "1" Then
            SQL = "Update am_stok set hpp='" & grid2.TextMatrix(grid2.Row, 3) & "'"
            SQL = SQL + " Where nolot='" & grid2.TextMatrix(grid2.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        grid2.Row = grid2.Row + 1
        Pg.Value = Pg.Value + 1
        DoEvents
    Loop
    OBJ.Close
    Pg.Value = 0
    MsgBox "Data is successfuly updated", vbInformation, AppName
    hapusgrid2
    lblsinc = "0"
    lblrow2 = "0 Baris"
End Sub

Private Sub Form_Load()
    grid.Cols = 8
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "Nolot"
    grid.TextMatrix(0, 3) = "Palet"
    grid.TextMatrix(0, 4) = "Item"
    grid.TextMatrix(0, 5) = "Hpp Bahan"
    grid.TextMatrix(0, 6) = "Hpp Kemasan"
    
    grid.ColWidth(0) = 450
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 0
    grid.ColWidth(3) = 1800
    grid.ColWidth(4) = 2800
    grid.ColWidth(5) = 1400
    grid.ColWidth(6) = 1400
    grid.ColWidth(7) = 200
    grid.ColAlignment(0) = flexAlignLeftCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignLeftCenter
    grid.ColAlignment(3) = flexAlignLeftCenter
    grid.ColAlignmentFixed(5) = flexAlignCenterCenter
    grid.ColAlignmentFixed(6) = flexAlignCenterCenter
    
    grid2.Cols = 5
    grid2.TextMatrix(0, 0) = "No"
    grid2.TextMatrix(0, 1) = "Nolot"
    grid2.TextMatrix(0, 2) = "perkilo"
    grid2.TextMatrix(0, 3) = "Actual perkilo"
    
    grid2.ColWidth(0) = 450
    grid2.ColWidth(1) = 1800
    grid2.ColWidth(2) = 2000
    grid2.ColWidth(3) = 2000
    grid2.ColWidth(4) = 200
    grid2.ColAlignment(1) = flexAlignLeftCenter
    grid2.ColAlignmentFixed(2) = flexAlignCenterCenter
    grid2.ColAlignmentFixed(3) = flexAlignCenterCenter
    
    date1 = Date
    date2 = Date
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
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
End Sub

Private Sub Tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    Select Case TabControl1.SelectedItem
'        Case 0:
'    End Select
End Sub
