VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmdatanull 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sinkron data hpp palet"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.CheckBox chklot 
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1455
      _Version        =   851970
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "nomor lot"
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
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11113
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
   Begin Chameleon.chameleonButton btnshow 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tampilkan hpp kosong"
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
      MICON           =   "frmdatanull.frx":0000
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
      Left            =   10080
      TabIndex        =   2
      Top             =   7320
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
      MICON           =   "frmdatanull.frx":031A
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
      Left            =   0
      TabIndex        =   3
      Top             =   6960
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
   Begin Chameleon.chameleonButton btnval 
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Sinc..."
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
      MICON           =   "frmdatanull.frx":0634
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
      Left            =   9120
      TabIndex        =   6
      Top             =   7320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "frmdatanull.frx":094E
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
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   3735
   End
End
Attribute VB_Name = "frmdatanull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim baris, i, jumlah As Integer

Private Sub btnshow_Click()
    If OBJ.State = 1 Then OBJ.Close
    hapusgrid
    
    OBJ.Open dsn
    If chklot.Value = xtpChecked Then
        SQL = "Select COUNT(palet)'jml' from am_stokgudang where hppperkg='0.00' and nolot <>''"
        Set RST = OBJ.Execute(SQL)
        jumlah = RST!jml
        Pg.Max = jumlah
        Pg.Value = 0
        Pg.Visible = True
    
        SQL = "Select * from am_stokgudang where hppperkg='0.00' and nolot <>''"
        Set RST = OBJ.Execute(SQL)
    ElseIf chklot.Value = xtpUnchecked Then
        SQL = "Select COUNT(palet)'jml' from am_stokgudang where hppperkg='0.00'"
        Set RST = OBJ.Execute(SQL)
        jumlah = RST!jml
        Pg.Max = jumlah
        Pg.Value = 0
        Pg.Visible = True
        
        SQL = "Select * from am_stokgudang where hppperkg='0.00'"
        Set RST = OBJ.Execute(SQL)
    End If
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!palet
        grid.TextMatrix(grid.Row, 3) = RST!kodebarang
        grid.TextMatrix(grid.Row, 4) = RST!NamaBarang
        grid.TextMatrix(grid.Row, 5) = Format(RST!kg, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 6) = Format(RST!kgperpalet, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 7) = Format(RST!hppperkg, "##,###,##0.000")
        grid.TextMatrix(grid.Row, 8) = ""   'validasi kg
        grid.TextMatrix(grid.Row, 9) = ""   'validasi kg/palet
        grid.TextMatrix(grid.Row, 10) = ""  'validasi hpp/kg
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

Private Sub btnval_Click()
    i = 0
    Pg.Max = baris
    Pg.Value = 0
    Pg.Visible = True
    Screen.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        OBJ.Open dsn
        SQL = "Select * From list_hpp_produksi Where palet='" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 8) = Format(RST!kg, "##,###,##0.00")
            grid.TextMatrix(grid.Row, 9) = Format(RST!kgperpalet, "##,###,##0.00")
            grid.TextMatrix(grid.Row, 10) = Format(RST!hppperkg, "##,###,##0.000")
        Else
            setAlternatingGridYelow grid.Row
        End If
        OBJ.Close
        DoEvents
        lblrow = "Palet sinkron: " & grid.Row
        grid.Row = grid.Row + 1
        Pg.Value = Pg.Value + 1
    Loop
    Pg.Value = 0
    cmdSave.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub cmdclose_Click()
    Unload Me
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

Private Sub cmdSave_Click()
    Me.MousePointer = vbHourglass
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        If grid.TextMatrix(grid.Row, 10) <> "" Then
        OBJ.Open dsn
        SQL = "Update am_stokgudang set kg= '" & grid.TextMatrix(grid.Row, 8) & "',"
        SQL = SQL + "kgperpalet='" & grid.TextMatrix(grid.Row, 9) & "',"
        SQL = SQL + "hppperkg='" & grid.TextMatrix(grid.Row, 10) & "'"
        SQL = SQL + " Where palet='" & grid.TextMatrix(grid.Row, 2) & "'"
        SQL = SQL + " and keterangan in ('Produksi Lem','Produksi Karpet','SJ')"
        SQL = SQL + " and hppperkg='0.00'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        End If
        grid.Row = grid.Row + 1
    Loop
    Me.MousePointer = vbDefault
    cmdSave.Enabled = False
    btnshow_Click
End Sub

Private Sub Form_Load()
    grid.Cols = 11
    grid.TextMatrix(0, 0) = "No"
    grid.TextMatrix(0, 1) = "Nolot"
    grid.TextMatrix(0, 2) = "Palet"
    grid.TextMatrix(0, 3) = "Kode"
    grid.TextMatrix(0, 4) = "Item"
    grid.TextMatrix(0, 5) = "Kg"
    grid.TextMatrix(0, 6) = "Kg/Palet"
    grid.TextMatrix(0, 7) = "Hpp/Kg"
    grid.TextMatrix(0, 8) = "Kg (sinc)"
    grid.TextMatrix(0, 9) = "Kg/Palet (sinc)"
    grid.TextMatrix(0, 10) = "Hpp/Kg (sinc)"
    
    grid.ColWidth(0) = 600
    grid.ColWidth(1) = 1800
    grid.ColWidth(2) = 1800
    grid.ColWidth(3) = 1100
    grid.ColWidth(4) = 2500
    grid.ColWidth(5) = 1100
    grid.ColWidth(6) = 1100
    grid.ColWidth(7) = 1100
    grid.ColWidth(8) = 1100
    grid.ColWidth(9) = 1100
    grid.ColWidth(10) = 1100
    grid.ColAlignment(0) = flexAlignLeftCenter
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(2) = flexAlignLeftCenter
    grid.ColAlignment(6) = flexAlignRightCenter
    grid.ColAlignmentFixed(5) = flexAlignCenterCenter
    grid.ColAlignmentFixed(10) = flexAlignCenterCenter
    
    ' Hooking the form for mouse wheel scroll
    Call WheelHook(Me.hWnd)
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
