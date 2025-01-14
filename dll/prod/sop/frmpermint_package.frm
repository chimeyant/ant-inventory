VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpermint_package 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulir Permintaan Kemasan"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtpetugas 
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
      TabIndex        =   1
      Top             =   480
      Width           =   2775
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
      Left            =   7455
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   0
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
      Left            =   7095
      Picture         =   "frmpermint_package.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   0
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
      Left            =   6840
      Picture         =   "frmpermint_package.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtnolot 
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
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   840
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
      Format          =   135462913
      CurrentDate     =   42039
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmpermint_package.frx":0630
      Caption         =   "frmpermint_package.frx":0650
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpermint_package.frx":06BC
      Keys            =   "frmpermint_package.frx":06DA
      Spin            =   "frmpermint_package.frx":071C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4128
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483642
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
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7200
      TabIndex        =   15
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
      MICON           =   "frmpermint_package.frx":0744
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
      Left            =   6240
      TabIndex        =   16
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "frmpermint_package.frx":0A5E
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
      Left            =   5280
      TabIndex        =   17
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
      MICON           =   "frmpermint_package.frx":0D78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Form ini digunakan untuk edit permintaan sebelum di konfirmasi oleh gudang"
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
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter No.Lot untuk Edit Data"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No.Lot Harus Sesuai dengan Lot SOP"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Petugas"
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
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl. Permintaan"
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
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No LOT"
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
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmpermint_package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub cmdclear_Click()
    txtnolot = ""
    txtpetugas = ""
    date1 = Date
    hapusgrid
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    If txtnolot = "" Then
        MsgBox "No.Lot belum diisi ", vbCritical, AppName
        Exit Sub
    ElseIf txtpetugas = "" Then
        MsgBox "Kolom Petugas belum diisi ", vbCritical, AppName
        Exit Sub
    End If
    
    
    OBJ.Open dsn
    SQL = "Select * From am_gudang_permintaan Where nolot='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!Status <> "0" Then
            OBJ.Close
            MsgBox "Item barang telah di confirm", vbCritical, AppName
            Exit Sub
        End If
        If MsgBox("Are you sure, you want to update this data ?", vbQuestion + vbYesNo, "Question") = vbNo Then OBJ.Close: Exit Sub
        'Update hanya berlaku jika permintaan belum dikonfirmasi oleh bag.gudang
            SQL = "Delete From am_gudang_permintaan Where nolot='" & txtnolot & "'"
            OBJ.Execute SQL
            
            SQL = "Select * From am_gudang_permintaan Where 0=1"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid.Row = 1
            Do While True
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                RST.AddNew
                RST!nolot = txtnolot
                RST!tgl = date1
                RST!petugas = txtpetugas
                RST!kodebarang = grid.TextMatrix(grid.Row, 1)
                RST!qty = grid.TextMatrix(grid.Row, 4)
                RST!Status = "0"
                RST!qty_add = "0"
                RST!keterangan = "Request"
                RST!flag = "0"
                RST!qty_confirmed = "0"
                RST!cetak_ke = "1"
                RST.Update
                grid.Row = grid.Row + 1
            Loop
            
            OBJ.Close
            MsgBox "Berhasil Diupdate", vbInformation, AppName
            cetakreport
            cmdclear_Click
        Exit Sub
    End If
    
    SQL = "Select * From am_gudang_permintaan Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!nolot = txtnolot
        RST!tgl = date1
        RST!petugas = txtpetugas
        RST!kodebarang = grid.TextMatrix(grid.Row, 1)
        RST!qty = grid.TextMatrix(grid.Row, 4)
        RST!Status = "0"
        RST!qty_add = "0"
        RST!keterangan = "Request"
        RST!flag = "0"
        RST!qty_confirmed = "0"
        RST!cetak_ke = "1"
        RST.Update
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    MsgBox "Berhasil Disimpan", vbInformation, AppName
    cetakreport
    cmdclear_Click
    
End Sub
Private Sub cetakreport()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetake_pack)"
    crystal.ReportFileName = AppPath & "\reports\produksi\take_pack.rpt"
    crystal.ParameterFields(0) = "@nolot;" & txtnolot.text & ";true"
    'crystal.ParameterFields(1) = "@username;" & nmuser & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub
Private Sub Form_Load()
    initGrid
    setGrid
    date1 = Date
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
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
                If txtnolot = "" Then Exit Sub
                carisql1 = "select kodebarang, namabarang from am_apitemmst where KodeProduk in('KTN/L','ETK/L','KLG/L','U/SP')"
                namatabel = "Kemasan"
                frmsearch.Show vbModal
        Case 4:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
    End Select
End Sub
Private Sub initGrid()
    With grid
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Kemasan"
        .TextMatrix(0, 3) = "LOT"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Kode Satuan"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "HPP"
    End With
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColWidth(5) = 500
        .ColWidth(6) = 750
        .ColWidth(7) = 0 '1500
    End With
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
    Set grid.CellPicture = blank
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
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
    'setGrid
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 1:
            grid.TextMatrix(grid.Row, 1) = hasil
            grid.TextMatrix(grid.Row, 2) = hasil1
            'cari satuan
            SQL = "select kodesatuan from am_apitemmst  "
            SQL = SQL + "where kodebarang='" & hasil & "'"
            OBJ.Open dsn
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid.TextMatrix(grid.Row, 4) = "0.00"
            grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                    
            'cari nama satuan
            SQL = "select * from am_apunit where kodesatuan ='" & grid.TextMatrix(grid.Row, 5) & "'"
            Set RST = OBJ.Execute(SQL)
            grid.TextMatrix(grid.Row, 6) = RST!namasatuan

            OBJ.Close
                    
            grid.Col = 0
            Set grid.CellPicture = uncheck

            If grid.Rows = grid.Row + 1 Then
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
            Else
            End If
            hasil = ""
            hasil1 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim hppbahan As Double
    If KeyAscii = 13 Then
        'GetStokBarang Format(date1, "yyyyMMdd"), grid.TextMatrix(grid.Row, 1), , , stokbahan
        
        'If stokbahan <= 0 Or stokbahan <= txtnilai.Value Then
            'MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
            'Exit Sub
        'End If
        grid.TextMatrix(grid.Row, 4) = txtnilai.text
        grid.TextMatrix(grid.Row, 7) = Format(getHPP(grid.TextMatrix(grid.Row, 1), stokbahan, txtnilai.Value), "##,###,###,##0.00")
        grid.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnolot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hapusgrid
        OBJ.Open dsn
        SQL = "SELECT a.*,b.NamaBarang,c.KodeSatuan,c.NamaSatuan FROM am_gudang_permintaan a"
        SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang "
        SQL = SQL + " inner join am_apunit c on b.KodeSatuan = c.KodeSatuan"
        SQL = SQL + " Where a.nolot='" & txtnolot & "' and a.status='0' and a.flag='0'"
        Set RST = OBJ.Execute(SQL)
        
        If Not RST.EOF Then
            txtpetugas = RST!petugas
        
            Do Until RST.EOF
                grid.Col = 0
                Set grid.CellPicture = uncheck
                grid.TextMatrix(grid.Row, 1) = RST!kodebarang
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "###,###,##0.00")
                grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                grid.TextMatrix(grid.Row, 6) = RST!namasatuan
                grid.TextMatrix(grid.Row, 7) = ""
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                RST.MoveNext
            Loop
        Else
            MsgBox "Data not Found.", vbExclamation, AppName
            txtnolot = ""
        End If
        OBJ.Close
    End If
End Sub
