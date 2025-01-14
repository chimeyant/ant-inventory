VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmmutfail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incomplete Mutation Palette (Data tidak lengkap)"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14220
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      ScaleHeight     =   375
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait....."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3375
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   6840
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
      MICON           =   "frmmutfail.frx":0000
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
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
   End
   Begin Chameleon.chameleonButton btnshow 
      Height          =   375
      Left            =   120
      TabIndex        =   2
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
      MICON           =   "frmmutfail.frx":031A
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
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   14055
      _Version        =   851970
      _ExtentX        =   24791
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
   Begin Chameleon.chameleonButton cmdproses 
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   120
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
      MICON           =   "frmmutfail.frx":0634
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
      Left            =   12240
      TabIndex        =   6
      Top             =   6840
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
      MICON           =   "frmmutfail.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "frmmutfail.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Width           =   3735
   End
End
Attribute VB_Name = "frmmutfail"
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
    hapusgrid
    cmdsave.Enabled = False
    Screen.MousePointer = vbHourglass
    Picture1.Visible = True
    OBJ.Open dsn
    SQL = "Select COUNT(palet)'jml'"
    SQL = SQL + " From am_stokgudang Where hppperkg='0.00'"
    Set RST = OBJ.Execute(SQL)
    
    jumlah = RST!jml
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    
    SQL = "Select nolot,palet,kodebarang,namabarang,kg,kgperpalet,hppperkg,qin,satuan"
    SQL = SQL + " From am_stokgudang Where hppperkg='0.00' order by nolot,palet asc"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!palet
        grid.TextMatrix(grid.Row, 3) = RST!kodebarang
        grid.TextMatrix(grid.Row, 4) = RST!NamaBarang
        grid.TextMatrix(grid.Row, 5) = Format(RST!kg, "##,##0.00")
        grid.TextMatrix(grid.Row, 6) = Format(RST!kgperpalet, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 7) = Format(RST!hppperkg, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 8) = Format(RST!qin, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 9) = RST!satuan
        grid.TextMatrix(grid.Row, 10) = "Null"
        grid.TextMatrix(grid.Row, 11) = "Null"
        grid.TextMatrix(grid.Row, 12) = "Null"
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        Pg.text = Str(Int((Pg.Value / Pg.Max) * 100)) & " % complete"
        lblrow = Pg.Value & " Palet"
        DoEvents
        RST.MoveNext
    Loop
    OBJ.Close
    Pg.Value = 0
    baris = grid.Row - 1
    lblrow = baris & " Palet"
    Screen.MousePointer = vbDefault
    Picture1.Visible = False
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    lblrow = "0 Palet."
    Pg.Value = 0
    Pg.Visible = False
    jumlah = 0
    cmdsave.Enabled = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdproses_Click()
    If grid.TextMatrix(1, 1) = "" Then Exit Sub
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    Picture1.Visible = True
    Screen.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        OBJ.Open dsn
        SQL = "Select kg,kgperpalet,hppperkg from list_hpp_produksi Where palet='" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 5) = Format(RST!kg, "##,##0.00")
            grid.TextMatrix(grid.Row, 6) = Format(RST!kgperpalet, "##,##0.00")
            grid.TextMatrix(grid.Row, 7) = Format(RST!hppperkg, "##,##0.00")
            grid.TextMatrix(grid.Row, 10) = "Data tersedia"
            grid.TextMatrix(grid.Row, 11) = "Data tersedia"
            grid.TextMatrix(grid.Row, 12) = "Data tersedia"
        Else
            grid.TextMatrix(grid.Row, 10) = "Data tidak tersedia"
            grid.TextMatrix(grid.Row, 11) = "Data tidak tersedia"
            grid.TextMatrix(grid.Row, 12) = "Data tidak tersedia"
            setAlternatingGridYelow grid.Row
        End If
        OBJ.Close
        DoEvents
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        Pg.text = "Proses melengkapi data " & Str(Int((Pg.Value / Pg.Max) * 100)) & " % complete"
    Loop
    Screen.MousePointer = vbDefault
    Picture1.Visible = False
    Pg.Value = 0
    Pg.Visible = False
    cmdsave.Enabled = True
End Sub

Private Sub cmdsave_Click()
    Pg.Max = jumlah
    Pg.Value = 0
    Pg.Visible = True
    Picture1.Visible = True
    Screen.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        OBJ.Open dsn
        SQL = "Update am_stokgudang set kg='" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + " kgperpalet='" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + " hppperkg='" & grid.TextMatrix(grid.Row, 7) & "'"
        SQL = SQL + " Where palet = '" & grid.TextMatrix(grid.Row, 2) & "' and keterangan like 'Produksi%'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        DoEvents
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
        Pg.text = "Proses melengkapi data " & Str(Int((Pg.Value / Pg.Max) * 100)) & " % complete"
    Loop
    Screen.MousePointer = vbDefault
    Picture1.Visible = False
    Pg.Value = 0
    Pg.Visible = False
    MsgBox "Data berhasil diperbaharui", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub Form_Load()
    grid.Cols = 13
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "Nolot"
    grid.TextMatrix(0, 2) = "Palet"
    grid.TextMatrix(0, 3) = "Kode"
    grid.TextMatrix(0, 4) = "Item"
    grid.TextMatrix(0, 5) = "kg"
    grid.TextMatrix(0, 6) = "kg/Palet"
    grid.TextMatrix(0, 7) = "hpp"
    grid.TextMatrix(0, 8) = "Qty"
    grid.TextMatrix(0, 9) = "Satuan"
    grid.TextMatrix(0, 10) = "Kg"
    grid.TextMatrix(0, 11) = "Kg/Palet"
    grid.TextMatrix(0, 12) = "Hpp"
    
    grid.ColWidth(0) = 600
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1700
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 2300
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 0
    grid.ColWidth(7) = 0
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 1000
    grid.ColWidth(10) = 1500
    grid.ColWidth(11) = 1500
    grid.ColWidth(12) = 1500
    grid.ColAlignment(0) = flexAlignLeftCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignLeftCenter
    grid.ColAlignment(6) = flexAlignRightCenter
    grid.ColAlignmentFixed(5) = flexAlignCenterCenter
    grid.ColAlignmentFixed(10) = flexAlignCenterCenter
    
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
        grid.TextMatrix(grid.Row, 11) = ""
        grid.TextMatrix(grid.Row, 12) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Function setAlternatingGridYelow(ByVal i As Integer)
    Dim j As Integer
    j = 0
    For j = 0 To 12
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
