VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmstokopnamebahanbaku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stok Peralatan Maintenance"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   9810
      TabIndex        =   17
      Top             =   690
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmstokopnamebahanbaku.frx":0000
      Caption         =   "frmstokopnamebahanbaku.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmstokopnamebahanbaku.frx":008C
      Keys            =   "frmstokopnamebahanbaku.frx":00AA
      Spin            =   "frmstokopnamebahanbaku.frx":00EC
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txttext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9000
      TabIndex        =   16
      Top             =   1335
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11220
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11700
      Picture         =   "frmstokopnamebahanbaku.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Left            =   11460
      Picture         =   "frmstokopnamebahanbaku.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3630
      Left            =   60
      TabIndex        =   11
      Top             =   1800
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   6403
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtket 
      Height          =   345
      Left            =   1395
      TabIndex        =   10
      Top             =   840
      Width           =   5625
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   1395
      TabIndex        =   8
      Top             =   465
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Format          =   143327233
      CurrentDate     =   41357
   End
   Begin VB.TextBox txtnobukti 
      Height          =   330
      Left            =   1395
      TabIndex        =   6
      Top             =   120
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   90
      TabIndex        =   4
      Top             =   5535
      Width           =   12060
      Begin XtremeSuiteControls.PushButton cmdSelesai 
         Height          =   360
         Left            =   11010
         TabIndex        =   3
         Top             =   240
         Width           =   945
         _Version        =   851970
         _ExtentX        =   1667
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Selesai"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdDelete 
         Height          =   360
         Left            =   10020
         TabIndex        =   2
         Top             =   240
         Width           =   960
         _Version        =   851970
         _ExtentX        =   1693
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Delete"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdsave 
         Height          =   360
         Left            =   9045
         TabIndex        =   1
         Top             =   240
         Width           =   960
         _Version        =   851970
         _ExtentX        =   1693
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Save"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdclear 
         Height          =   360
         Left            =   8115
         TabIndex        =   0
         Top             =   240
         Width           =   900
         _Version        =   851970
         _ExtentX        =   1587
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Clear"
         Appearance      =   6
      End
   End
   Begin VB.Line Line3 
      X1              =   8370
      X2              =   8370
      Y1              =   1440
      Y2              =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Kemasan"
      Height          =   225
      Left            =   6630
      TabIndex        =   12
      Top             =   1500
      Width           =   1470
   End
   Begin VB.Line Line2 
      X1              =   6540
      X2              =   6540
      Y1              =   1440
      Y2              =   1695
   End
   Begin VB.Line Line1 
      X1              =   6540
      X2              =   8385
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan"
      Height          =   210
      Left            =   45
      TabIndex        =   9
      Top             =   870
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
      Height          =   210
      Left            =   30
      TabIndex        =   7
      Top             =   540
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "No Bukti"
      Height          =   210
      Left            =   60
      TabIndex        =   5
      Top             =   195
      Width           =   1200
   End
End
Attribute VB_Name = "frmStokOpnameBahanBaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String
Private kodestok As String

Private posrow As Integer

Private Sub cmdclear_Click()
    txtnobukti = ""
    date1 = Date
    txtket = ""
    hapusgrid
    history
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err_msg:
    If txtnobukti = "" Then
        MsgBox "Data Not Completed...", vbCritical, AppName
        Exit Sub
    End If
    
    If grid.TextMatrix(1, 1) = "" Then
        MsgBox "Data Not Completed..", vbCritical, AppName
        Exit Sub
    End If
    
    'save to header
    SQL = "Insert Into am_apstokopnamebahanhdr ("
    SQL = SQL + "nobukti,"
    SQL = SQL + "tgl,"
    SQL = SQL + "ket,"
    SQL = SQL + "username) "
    SQL = SQL + "Values('"
    SQL = SQL + txtnobukti + "',"
    SQL = SQL + "Convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'"
    SQL = SQL + txtket + "','"
    SQL = SQL + UserOnline + "')"
    
    OBJ.Open dsn
    Set RS = OBJ.Execute(SQL)
    OBJ.Close
    
    'Save To Detail
    grid.Row = 1
    OBJ.Open dsn
    Do While True
        With grid
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "Insert Into am_apstokopnamebahanin ("
            SQL = SQL + "nobukti,"
            SQL = SQL + "kdbarang,"
            SQL = SQL + "nolot,"
            SQL = SQL + "kdpckg,"
            SQL = SQL + "qty,"
            SQL = SQL + "kdsatuan) "
            SQL = SQL + "Values('"
            SQL = SQL + txtnobukti + "','"
            SQL = SQL + .TextMatrix(.Row, 1) + "','"
            SQL = SQL + .TextMatrix(.Row, 3) + "','"
            SQL = SQL + .TextMatrix(.Row, 4) + "',"
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "'),'"
            SQL = SQL + .TextMatrix(.Row, 7) + "')"
            Set RS = OBJ.Execute(SQL)
            .Row = .Row + 1
            DoEvents
        End With
    Loop
    OBJ.Close
    
    'save to stok
    
    grid.Row = 1
    kodestok = AmbilKodeStokBaru(Format(date1, "yy.MM"))
    OBJ.Open dsn
    Do While True
        With grid
            If .TextMatrix(.Row, 1) = "" Then Exit Do
            SQL = "INSERT INTO am_stokbahan ("
            SQL = SQL + "kdstok,"
            SQL = SQL + "kdbarang,"
            SQL = SQL + "nolot,"
            SQL = SQL + "trans,"
            SQL = SQL + "ref,"
            SQL = SQL + "line,"
            SQL = SQL + "tgl,"
            SQL = SQL + "kdpckg,"
            SQL = SQL + "kdsatuan,"
            SQL = SQL + "awal,"
            SQL = SQL + "h_awal,"
            SQL = SQL + "masuk,"
            SQL = SQL + "h_masuk,"
            SQL = SQL + "keluar,"
            SQL = SQL + "h_keluar"
            SQL = SQL + ") "
            SQL = SQL + " Values('"
            SQL = SQL + kodestok + "','"  'kdstok
            SQL = SQL + .TextMatrix(.Row, 1) + "','" 'kdbarang
            SQL = SQL + .TextMatrix(.Row, 3) + "','"  'nolot
            SQL = SQL + "O" + "','"
            SQL = SQL + txtnobukti + "'," 'ref
            SQL = SQL + "convert(numeric, '" + Format(.Row, "general number") + "')," 'line
            SQL = SQL + "convert(datetime,'" + Format(date1, "MM/dd/yyyy") + "'),'" 'tgl
            SQL = SQL + .TextMatrix(.Row, 4) + "','" 'satuan package
            SQL = SQL + .TextMatrix(.Row, 7) + "'," 'satuan
            SQL = SQL + "convert(money,'" + Format(.TextMatrix(.Row, 6), "general number") + "')," 'awal
            SQL = SQL + "convert(money,'0')," 'h_awal
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0')," 'h_masuk
            SQL = SQL + "convert(money,'0')," 'keluar
            SQL = SQL + "convert(money,'0')"  'h_keluar
            SQL = SQL + ")"
            
            Set RS = OBJ.Execute(SQL)
            .Row = .Row + 1
        End With
        DoEvents
    Loop
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    Exit Sub
err_msg:
    MsgBox Err.Description
End Sub

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'grid
    With grid
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kd Bahan"
        .TextMatrix(0, 2) = "Nama Bahan"
        .TextMatrix(0, 3) = "No LOT"
        .TextMatrix(0, 4) = "K/Pck"
        .TextMatrix(0, 5) = "Packaging"
        .TextMatrix(0, 6) = "Qty"
        .TextMatrix(0, 7) = "K/Sat"
        .TextMatrix(0, 8) = "Satuan"
    End With
    setgrid
End Sub

Private Sub setgrid()
    With grid
        .RowHeightMin = 300
        .ColWidth(0) = 250
        .ColWidth(1) = 1200
        .ColWidth(2) = 2500
        .ColWidth(3) = 2500
        .ColWidth(4) = 800
        .ColWidth(5) = 1000
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 1000
    End With
End Sub

Private Sub history()
    On Error GoTo err_msg
    Dim str99 As String
    Dim strtgl As String
    Dim SQL As String
    
    strtgl = Format(date1, "YYMM")
    
    OBJ.Open dsn
    SQL = "select top 1 nobukti from am_apstokopnamebahanhdr where nobukti like 'OP-" & strtgl & "%' order by nobukti desc"
    
    Set RS = OBJ.Execute(SQL)
    If Not RS.EOF Then
        str99 = Right(RS!nobukti, 3)
    Else
        str99 = 0
    End If
    
    str99 = str99 + 1
    
    If Len(str99) = 1 Then txtnobukti = "OP-" + strtgl + "." + "00" & str99
    If Len(str99) = 2 Then txtnobukti = "OP-" + strtgl + "." & "0" & str99
    If Len(str99) = 3 Then txtnobukti = "OP-" + strtgl & "." + str99
        
    OBJ.Close
    Exit Sub
err_msg:
    OBJ.Close
    MsgBox Err.Description
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
        
    posrow = grid.Row
    
    Select Case grid.Col
        Case 0:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.CellPicture = uncheck Then
                    Set grid.CellPicture = check
                    If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                        Set grid.CellPicture = uncheck
                        hapusrow
                        Exit Sub
                    End If
                    Set grid.CellPicture = uncheck
                End If
        Case 1:
                carisql1 = "select kodebarang, namabarang from am_apitemmst"
                namatabel = "Bahan Baku"
                frmsearch.Show vbModal
          Case 3:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                With txttext
                    .Width = grid.ColWidth(grid.Col) - 40
                    .text = grid.TextMatrix(grid.Row, grid.Col)
                    .Left = grid.Left + grid.CellLeft
                    .Top = grid.Top + grid.CellTop + 20
                    .Height = grid.CellHeight - 40
                    .Visible = True
                    .SetFocus
                End With
            Case 4:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                carisql1 = "select * from am_appackaging"
                namatabel = "Packaging"
    
                frmsearch.Show vbModal
            Case 6:
                With txtnilai
                    .Width = grid.ColWidth(grid.Col) - 40
                    .Value = grid.TextMatrix(grid.Row, grid.Col)
                    .Left = grid.Left + grid.CellLeft
                    .Top = grid.Top + grid.CellTop + 20
                    .Height = grid.CellHeight - 40
                    .Visible = True
                    .SetFocus
                End With
    End Select
    
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 1:
                With grid
                    .Col = 0
                    Set .CellPicture = uncheck
                    .TextMatrix(.Row, 1) = hasil
                    .TextMatrix(.Row, 2) = hasil1
                    .Rows = .Rows + 1
                End With
                
                hasil = ""
                hasil1 = ""
                hasil2 = ""
                hasil3 = ""
                carisql1 = ""
                namatabel = ""
                'cari satuan
                SQL = "select a.kodesatuanmutasi ,b.namasatuan from am_apitemmst a inner join am_apunit b on b.kodesatuan=a.kodesatuanmutasi "
                SQL = SQL + " where a.kodebarang='" + grid.TextMatrix(grid.Row, 1) + "'"
                OBJ.Open dsn
                Set RS = OBJ.Execute(SQL)
                With grid
                    .TextMatrix(.Row, 7) = RS!kodesatuanmutasi
                    .TextMatrix(.Row, 8) = RS!namasatuan
                End With
                OBJ.Close
        Case 4:
                With grid
                    .TextMatrix(.Row, 4) = hasil
                    .TextMatrix(.Row, 5) = hasil1
                End With
                hasil = ""
                hasil1 = ""
                hasil2 = ""
                carisql1 = ""
                namatabel = ""
    End Select
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0

        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txttext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(posrow, 3) = txttext
        grid.SetFocus
    End If
End Sub

Private Sub txttext_LostFocus()
    txttext.Visible = False
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
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setgrid
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub
