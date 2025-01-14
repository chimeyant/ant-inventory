VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frminquery1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inquiry Purchase Order"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
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
      MICON           =   "frminquery1.frx":0000
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3960
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
      MICON           =   "frminquery1.frx":031A
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
      TabIndex        =   3
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Item"
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
      MICON           =   "frminquery1.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frminquery1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
    grid.Clear
    grid.Rows = 2
    txtkode1 = ""
    Label1 = ""
    cmdsearch1.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cari()
    If txtkode1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT (a.nopo)'PO',(a.qty)'QTY',(c.namasatuan)'SATUAN' FROM am_polin a left join am_pohdr b on a.nopo=b.nopo left join am_apunit c on a.kodesatuan=c.kodesatuan WHERE a.kodebarang = '" & txtkode1 & "' and b.tglpo >= '" & batas1 & "' and b.tglpo <= '" & batas2 & "' order by a.nopo"
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    grid.ColWidth(0) = 1600
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 750
    
    If grid.Rows > 1 Then
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(grid.Rows - 1, 0) = "Total : "
        grid.TextMatrix(grid.Rows - 1, 2) = grid.TextMatrix(grid.Rows - 2, 2)
        OBJ.Open dsn
        SQL = "SELECT sum(a.qty)'QTY' FROM am_polin a left join am_pohdr b on a.nopo=b.nopo left join am_apunit c on a.kodesatuan=c.kodesatuan WHERE a.kodebarang = '" & txtkode1 & "' and b.tglpo >= '" & batas1 & "' and b.tglpo <= '" & batas2 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then grid.TextMatrix(grid.Rows - 1, 1) = RST!qty
        OBJ.Close
    End If
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode1 = hasil
    Label1 = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    
    cari
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
