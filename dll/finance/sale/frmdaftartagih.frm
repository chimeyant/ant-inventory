VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftartagih 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Penagihan"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
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
   Icon            =   "frmdaftartagih.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Crystal 
      Left            =   720
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   5280
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
      Left            =   3600
      Picture         =   "frmdaftartagih.frx":2372
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   5280
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
      Left            =   3360
      Picture         =   "frmdaftartagih.frx":2728
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   130351107
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Collector"
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
      MICON           =   "frmdaftartagih.frx":2ADE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   5520
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
      MICON           =   "frmdaftartagih.frx":2DF8
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
      Left            =   7320
      TabIndex        =   5
      Top             =   5520
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
      MICON           =   "frmdaftartagih.frx":3112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   5520
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
      MICON           =   "frmdaftartagih.frx":342C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdcetak 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View Tanda Terima"
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
      MICON           =   "frmdaftartagih.frx":3746
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "No Tagih"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Tagih"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   510
      Width           =   1455
   End
End
Attribute VB_Name = "frmdaftartagih"
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

Dim posrow, poscol As String

Private Sub cmdcetak_Click()
Dim Nofaktur As String

    grid.Row = 1
    If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        If Nofaktur = "" Then
            Nofaktur = "'" + grid.TextMatrix(grid.Row, 1) + "'"
        Else
            Nofaktur = Nofaktur + ",'" + grid.TextMatrix(grid.Row, 1) + "'"
        End If
        grid.Row = grid.Row + 1
    Loop
    
    SQL = "Select * From am_aropnfil Where noapply in(" & Nofaktur & ") ORDER BY noapply DESC"
    With rpttandaterima
        .Field6.text = grid.TextMatrix(grid.Row, 4)
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With

End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='274' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtnobukti_GotFocus()
    Call Blok(txtnobukti)
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub hapusemua()
    date1 = Date
    txtkodecust = ""
    lblnamacust = ""
    hapusgrid
End Sub

Private Sub txtkodecust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtkodecust_LostFocus()
    If txtkodecust <> "" Then caricollector
End Sub

Private Sub caricollector()
    If txtkodecust = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_collector where kode = '" & txtkodecust & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Collector " & RST!nama & " is not active !", vbExclamation, "Warning"
        txtkodecust = ""
        lblnamacust = ""
        txtkodecust.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kode, nama, idupdate from am_collector"
    namatabel = "Collector"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecust = hasil
    lblnamacust = hasil1
    caricollector
    hasil = ""
    hasil1 = ""
    grid.SetFocus
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
        
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 2000
    grid.ColWidth(5) = 2000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 0
    
    grid.Col = 0
    grid.Row = 1
    Set grid.CellPicture = check
End Sub

Private Sub grid_Click()
    Dim j As Integer
    j = 0
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtkodecust = "" Then Exit Sub
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 0
            If grid.CellPicture = check Then
                setup1 = grid.Row
                frmdaftartagihsub.Show 1
                
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                grid.Col = 0
                Set grid.CellPicture = uncheck
    
                grid.Rows = grid.Rows + 1
                grid.Col = 0
                grid.Row = grid.Rows - 1
                Set grid.CellPicture = check
                
                grid.Col = 1
            ElseIf grid.CellPicture = uncheck Then
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HE0E0E0
                Next
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, AppName) = vbNo Then
                    For j = 0 To grid.Cols - 1
                        grid.Col = j
                        grid.CellBackColor = &HFFFFFF
                    Next
                    Exit Sub
                Else
                    For j = 0 To grid.Cols - 1
                        grid.Col = j
                        grid.CellBackColor = &HFFFFFF
                    Next
                    hapusrow
                End If
            End If
    End Select
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtkodecust = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
       
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 6) = "0.00" Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop
        
    OBJ.Open dsn
    SQL = "delete from am_aropninv where identry = '" & kuser & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        in1
        
        grid.Row = grid.Row + 1
    Loop
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\daftartagih.rpt"
    Crystal.DataFiles(0) = "Proc(am_daftartagih)"
    Crystal.ParameterFields(0) = "@kode1;" & txtnobukti & ";true"
    Crystal.ParameterFields(1) = "@kode2;" & txtkodecust & ";true"
    Crystal.ParameterFields(2) = "@kode3 ;" + kuser + ";true"
    Crystal.ParameterFields(3) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub in1()
    OBJ.Open dsn
    SQL = "insert into am_aropninv ("
    SQL = SQL + "identry,"
    SQL = SQL + "kodecust,"
    SQL = SQL + "namacust,"
    SQL = SQL + "alamat,"
    SQL = SQL + "faktur,"
    SQL = SQL + "tanggal,"
    SQL = SQL + "jumlah,"
    SQL = SQL + "lineitem)"
    
    SQL = SQL + " values("
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "'" & grid.TextMatrix(grid.Row, 3) & "',"
    SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
    SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
    SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
    SQL = SQL + "convert(datetime,'" & tanggalgrid & "'),"
    SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 6), "general number")) & "'),"
    SQL = SQL + "convert(numeric,'" & grid.Row & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub cmdclear_Click()
    hapusemua
    txtnobukti = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    grid.TextMatrix(0, 1) = "Faktur"
    grid.TextMatrix(0, 2) = "Tanggal"
    grid.TextMatrix(0, 3) = "Kode Cust"
    grid.TextMatrix(0, 4) = "Nama Cust"
    grid.TextMatrix(0, 5) = "Alamat"
    grid.TextMatrix(0, 6) = "Nilai"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 2000
    grid.ColWidth(5) = 2000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 0
    
    grid.RowHeightMin = 300
    
    date1.Value = Date
    
    grid.Col = 0
    grid.Row = 1
    Set grid.CellPicture = check
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
    
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    'Set grid.CellPicture = blank
    Set grid.CellPicture = check
End Sub

Function tanggalgrid()
      tanggalgrid = Month(grid.TextMatrix(grid.Row, 2)) & "/" & Day(grid.TextMatrix(grid.Row, 2)) & "/" & Year(grid.TextMatrix(grid.Row, 2))
End Function

Function tanggalinv()
      tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
      tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
