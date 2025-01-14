VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmterimawiprepair 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repair data stok"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtkode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin Chameleon.chameleonButton cmdconfirm 
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   5160
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
      MICON           =   "frmterimawiprepair.frx":0000
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
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
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
   Begin Chameleon.chameleonButton cmdshow 
      Height          =   375
      Left            =   3000
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
      MICON           =   "frmterimawiprepair.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   5160
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
      MICON           =   "frmterimawiprepair.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   6975
      _Version        =   851970
      _ExtentX        =   12303
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblrow 
      Caption         =   "0 Record"
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
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "No. BPB"
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
      Top             =   200
      Width           =   1095
   End
End
Attribute VB_Name = "frmterimawiprepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String
Dim jumlah As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdconfirm_Click()
    'update ;am_bpbhdr   ;am_bpblin  :am_stok
    OBJ.Open dsn
    

    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        'Label1 = grid.Row
        'DoEvents
        SQL = "Update am_bpbhdr set nobpb='" & grid.TextMatrix(grid.Row, 4) & "'"
        SQL = SQL + " Where nobpb= '" & txtkode & "' and keterangan='" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_bpblin set nobpb='" & grid.TextMatrix(grid.Row, 4) & "'"
        SQL = SQL + " Where nobpb= '" & txtkode & "' and keterangan='" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update am_stok set noref='" & grid.TextMatrix(grid.Row, 4) & "'"
        SQL = SQL + " Where noref= '" & txtkode & "' and palet='" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    MsgBox "Data is successfully updated", vbInformation, AppName
End Sub

Private Sub cmdshow_Click()
    Dim newbpb As String
    Dim no As Integer
    'nomor new bpb nya di cari dulu yang kosong terakhir berapa !
    newbpb = "PG0-21121"    'PG0-21121445
    no = 623
    
    hapusgrid
    OBJ.Open dsn
    SQL = "Select COUNT(nobpb)'jml' from am_bpbhdr where nobpb = '" & txtkode & "' and kodegudang = 'G3'"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select distinct a.nobpb,a.keterangan,a.tglbpb From am_bpbhdr a inner join am_bpblin b"
    SQL = SQL + " on a.keterangan = b.keterangan Where a.nobpb = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        MsgBox "Data tidak ditemukan", vbCritical, AppName
        OBJ.Close
        Exit Sub
    End If
        
    grid.Row = 1
    Do While Not RST.EOF
        With grid
            no = no + 1
            .TextMatrix(.Row, 0) = grid.Row
            .TextMatrix(.Row, 1) = RST!nobpb
            .TextMatrix(.Row, 2) = RST!keterangan
            .TextMatrix(.Row, 3) = RST!tglbpb
            .TextMatrix(.Row, 4) = newbpb & no

            lblrow = .Row
            .Rows = .Rows + 1
            .Row = .Row + 1
            Pg.Value = Pg.Value + 1
            RST.MoveNext
        End With
        DoEvents
    Loop
    Pg.Value = 0
    lblrow = jumlah & " Record"
    OBJ.Close
End Sub

Private Sub Form_Load()
    grid.Cols = 5
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "No.BPB"
    grid.TextMatrix(0, 2) = "Palet"
    grid.TextMatrix(0, 3) = "Tanggal"
    grid.TextMatrix(0, 4) = "New No.BPB"
    
    grid.ColWidth(0) = 600
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1800
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1500
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignLeftCenter
    'grid.ColAlignmentFixed(2) = flexAlignCenterCenter
    'grid.ColAlignmentFixed(3) = flexAlignCenterCenter
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
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub
