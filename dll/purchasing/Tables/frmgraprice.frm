VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmgraprice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grafik Harga Supplier"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
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
      Height          =   1785
      Left            =   1560
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox List1 
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
      Height          =   1980
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtbrg 
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   7
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtnama 
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2040
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
      MICON           =   "frmgraprice.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Nama Barang"
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
      MICON           =   "frmgraprice.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtahun 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Calculator      =   "frmgraprice.frx":0634
      Caption         =   "frmgraprice.frx":0654
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmgraprice.frx":06B9
      Keys            =   "frmgraprice.frx":06D7
      Spin            =   "frmgraprice.frx":0721
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   2100
      MinValue        =   2005
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   2005
      MaxValueVT      =   1330839557
      MinValueVT      =   1431175173
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   0
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview"
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
      MICON           =   "frmgraprice.frx":0749
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
      BackColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Tahun"
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
      Left            =   840
      TabIndex        =   9
      Top             =   870
      Width           =   615
   End
   Begin VB.Label lblkode 
      Caption         =   "kode"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Supplier"
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
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmgraprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String


Private Sub cari()
    If txtnama = "" Then
        List1.Visible = False
        Exit Sub
    End If
    List1.Clear
    
    OBJ.Open dsn
    SQL = "select namasupp from am_supplier where namasupp like '" & txtnama & "%' order by namasupp"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            List1.AddItem RST!namasupp
            RST.MoveNext
        Loop
        List1.Visible = True
    Else
        List1.Visible = False
    End If
    OBJ.Close
End Sub

Private Sub caribrg()
    If txtbrg = "" Then
        List2.Visible = False
        Exit Sub
    End If
    List2.Clear
    
    OBJ.Open dsn
    SQL = "select a.kodebarang, a.namabarang from am_apitemmst a"
    SQL = SQL + " inner join am_price b on a.KodeBarang = b.kodebarang"
    SQL = SQL + " where b.kodesupp = '" & lblkode & "' and a.namabarang like '" & txtbrg & "%' order by a.namabarang"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            List2.AddItem RST!namabarang
            RST.MoveNext
        Loop
        List2.Visible = True
    Else
        List2.Visible = False
    End If
    OBJ.Close
End Sub

Private Sub cmdview_Click()
Dim thn As String
    thn = txtahun
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_pricegraph)"
    Crystal.ReportFileName = AppPath & "\reports\purchasing\tables\pricegraph.rpt"

    Crystal.ParameterFields(0) = "@kode1;" + lblkode + ";true"
    Crystal.ParameterFields(1) = "@kode2;" + txtkode1 + ";true"
    Crystal.ParameterFields(2) = "@tahun;" + thn + ";true"
    Crystal.ParameterFields(3) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtahun = Year(Date)
End Sub

Private Sub List1_DblClick()
    txtnama = List1.text
    txtnama = Trim(txtnama)
    
    OBJ.Open dsn
    SQL = "select kodesupp from am_supplier where namasupp = '" & txtnama & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblkode = RST!KodeSupp
    End If
    OBJ.Close
    List1.Visible = False

End Sub

Private Sub List2_DblClick()
    txtbrg = List2.text
    txtbrg = Trim(txtbrg)
    
    OBJ.Open dsn
    SQL = "select a.kodebarang, a.namabarang from am_apitemmst a"
    SQL = SQL + " inner join am_price b on a.KodeBarang = b.kodebarang"
    SQL = SQL + " where b.kodesupp = '" & lblkode & "' and a.namabarang = '" & txtbrg & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode1 = RST!kodebarang
    End If
    OBJ.Close
    List2.Visible = False
End Sub

Private Sub txtbrg_Change()
    If txtnama = "" Then Exit Sub
    caribrg
End Sub

Private Sub txtnama_Change()
    cari
End Sub

Private Sub cmdsearch1_Click()
    'carisql1 = "select kodebarang, namabarang from am_apitemmst"
    carisql1 = "select a.kodebarang, a.namabarang from am_apitemmst a"
    carisql1 = carisql1 + " inner join am_price b on a.KodeBarang = b.kodebarang"
    carisql1 = carisql1 + " where b.kodesupp = '" & lblkode & "'"
    namatabel = ".Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode1 = hasil
    txtbrg = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub txtnama_Click()
    OBJ.Open dsn
    SQL = "select a.kodebarang, a.namabarang from am_apitemmst a"
    SQL = SQL + " inner join am_price b on a.KodeBarang = b.kodebarang"
    SQL = SQL + " where b.kodesupp = '" & lblkode & "'"
    OBJ.Close
End Sub
