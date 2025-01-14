VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmscanlot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan Lot"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   5760
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   10160
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   8421504
      AllowUserResizing=   1
      Appearance      =   0
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
      _Band(0).Cols   =   7
   End
   Begin VB.PictureBox Picgrid 
      BackColor       =   &H00404040&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   6345
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6405
      Begin VB.TextBox txtinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmscanlot.frx":0000
         Top             =   60
         Width           =   6210
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   5115
      Left            =   45
      TabIndex        =   8
      Top             =   1020
      Width           =   6315
      _Version        =   851970
      _ExtentX        =   11139
      _ExtentY        =   9022
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
      Appearance      =   2
      PaintManager.Position=   2
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Produk"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "grid1"
      Item(0).Control(1)=   "txtnilai"
      Item(1).Caption =   "Kemasan"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "grid2"
      Item(1).Control(1)=   "txtqty"
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   465
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmscanlot.frx":0006
         Caption         =   "frmscanlot.frx":0026
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscanlot.frx":0092
         Keys            =   "frmscanlot.frx":00B0
         Spin            =   "frmscanlot.frx":00F2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8454143
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   4305
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7594
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   8421504
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
         _Band(0).Cols   =   6
      End
      Begin TDBNumber6Ctl.TDBNumber txtqty 
         Height          =   255
         Left            =   -69790
         TabIndex        =   18
         Top             =   450
         Visible         =   0   'False
         Width           =   90
         _Version        =   65536
         _ExtentX        =   159
         _ExtentY        =   450
         Calculator      =   "frmscanlot.frx":011A
         Caption         =   "frmscanlot.frx":013A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscanlot.frx":01A6
         Keys            =   "frmscanlot.frx":01C4
         Spin            =   "frmscanlot.frx":0206
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8454143
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,##0;(##,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,##0;(##,###,##0)"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   4305
         Left            =   -69880
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7594
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   8421504
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
         _Band(0).Cols   =   8
      End
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
      Left            =   1215
      Picture         =   "frmscanlot.frx":022E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   6420
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
      Left            =   510
      Picture         =   "frmscanlot.frx":057C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   5475
      Visible         =   0   'False
      Width           =   255
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
      Left            =   150
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtpalet 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   0
      Top             =   180
      Width           =   4725
   End
   Begin VB.TextBox txtlot 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   3
      Top             =   600
      Width           =   4725
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   5250
      TabIndex        =   1
      Top             =   6255
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
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
   Begin XtremeSuiteControls.PushButton btnclear 
      Height          =   465
      Left            =   4125
      TabIndex        =   13
      Top             =   6255
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Clear"
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
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   465
      Left            =   3000
      TabIndex        =   16
      Top             =   6255
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Save"
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
   Begin XtremeSuiteControls.PushButton btnGroup 
      Height          =   465
      Left            =   75
      TabIndex        =   17
      Top             =   6240
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "+ Add Group"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Label Label2 
      Caption         =   "Nomor Palet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   210
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Lot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   2
      Top             =   630
      Width           =   1335
   End
End
Attribute VB_Name = "frmscanlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event doexit(ByVal s_exit As Boolean)

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset
Private SQL1 As String
Dim kode_stok As String
Private package, kaleng, etiket, showmode As Boolean

Dim str1, str2, str3, str4, str99 As String

Private Sub btnclear_Click()
    txtlot = ""
    txtpalet = ""
    txtnilai = ""
    txtqty = ""
    txtnilai.Visible = False
    txtqty.Visible = False
    package = False: kaleng = False: etiket = False: showmode = False:
    str1 = "": str2 = "": str3 = "": str4 = ""
    hapusgrid
    hapusgrid1
    hapusgrid2
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnGroup_Click()
    frmgroup.Show vbModal
End Sub

Private Sub btnsave_Click()
'    On Error GoTo Err_handler:
    Dim strformat   As String
    Dim strbpb      As String
    Dim strkode     As String
    Dim strelapsed, strstart, strend As String
    Dim strhpppack  As Double
    Dim StartTime   As Variant
    Dim EndTime     As Variant
    Dim ElapsedTime As Variant
    Dim strnomut As String
    Dim strkorl As String
    
    strformat = Format(Date, "yymm")
    'Time Block
    StartTime = Format(Time, "hh:mm")
    strend = "12:00"
    EndTime = Format(strend, "hh:mm")
    ElapsedTime = CDate(StartTime) + CDate(EndTime)
    
    strstart = CDate(Format(StartTime, "hh:mm"))
    strelapsed = CDate(Format(ElapsedTime, "hh:mm"))
    
    strnomut = getnomut
    If grid1.TextMatrix(grid1.Row, 1) = "" Then
        MsgBox "Produk hasil perolehan belum diisi", vbCritical, "WARNING !"
        Exit Sub
    End If
    
    
    If txtpalet = "" Or txtlot = "" Then Exit Sub
    If showmode = True Then
        MsgBox "Cannot update, View Only mode", vbExclamation, "WARNING"
        Exit Sub
    End If

    OBJ.Open dsn
    SQL = "Select * From list_produksi_kemasan Where noref = '" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)

    If Not RST.EOF Then
        MsgBox "Palet : " & txtpalet & vbCrLf & "Sudah discan dan telah tersimpan.", vbExclamation, AppName
        OBJ.Close
        btnclear_Click
        Exit Sub
    End If
    'SIMPAN KE LIST_PRODUKSI_HASIL
    SQL = "Select * From list_produksi_hasil Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = str1
        RST!nolot = txtlot
        RST!kode_bahan = grid1.TextMatrix(grid1.Row, 1)
        RST!Lot_bahan = ""
        RST!qty_bahan = grid1.TextMatrix(grid1.Row, 4)
        RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
        RST!flag_tambahan = "1"
        RST!tanggal = Format(Date, "yyyy/MM/dd")
        RST!noref = txtpalet
        RST!proses_ke = "2"
        RST.Update
        If grid1.Rows = grid1.Row + 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop
    
    'SIMPAN KE LIST_PRODUKSI_KEMASAN
    SQL = "select * from list_produksi_kemasan where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_produk = str1
        RST!nolot = txtlot
        RST!kode_bahan = grid2.TextMatrix(grid2.Row, 1)
        RST!Lot_bahan = ""
        'Cek Konversi
        OBJ1.Open dsn
        SQL1 = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            'Konversi balik ke satuan awal
            RST!qty_bahan = grid2.TextMatrix(grid2.Row, 4) / RST1!nilai
        Else
            RST!qty_bahan = grid2.TextMatrix(grid2.Row, 4)
        End If
        OBJ1.Close
        'RST!qty_bahan = grid2.TextMatrix(grid2.Row, 4)
        RST!KODE_SATUAN = grid2.TextMatrix(grid2.Row, 6)
        RST!flag_tambahan = "0"
        RST!hpp = grid2.TextMatrix(grid2.Row, 7)
        RST!tanggal = Format(Date, "yyyy/MM/dd")
        RST!noref = txtpalet
        RST!proses_ke = "0"
        RST.Update
'       grid2.Row = grid2.Row + 1
        If grid2.Rows = grid2.Row + 1 Then Exit Do
        grid2.Row = grid2.Row + 1
    Loop

    'SIMPAN KE LIST_MUTASI_PRODUKSI_HEADER
    SQL = "Select * From list_mutasi_produksi_header Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kode_palet = txtpalet
        !tanggal = Format(Date, "yyyy-mm-dd hh:mm:ss")
        !kode_produk = str1
        !nomor_lot = txtlot
        !ref1 = strstart
        !ref2 = strelapsed
        !UserName = nmuser
        !Status = "0"
        .Update
    End With
    
    'SIMPAN KE LIST_MUTASI_PRODUKSI_DETAIL
    SQL = "Select * From list_mutasi_produksi_details Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kode_palet = txtpalet
        RST!kode_barang = grid1.TextMatrix(grid1.Row, 1)
        RST!KODE_SATUAN = grid1.TextMatrix(grid1.Row, 5)
        RST!qty = grid1.TextMatrix(grid1.Row, 4)
        RST!baris = grid1.Row
        RST.Update
        If grid1.Rows = grid1.Row + 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop

    'SIMPAN KE AM_USEHDR
    SQL = "Select * From am_usehdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid2.Row = 1
    If grid2.TextMatrix(grid2.Row, 1) = "" Then GoTo uselin:
        With RST
            .AddNew
            !nobpb = txtpalet
            !tglbpb = Format(Date, "yyyy/MM/dd")
            !noorder = txtpalet
            .Update
        End With
        
uselin:
    'SIMPAN KE USELIN
    SQL = "Select * From am_uselin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!nobpb = txtpalet
        RST!kodebarang = grid2.TextMatrix(grid2.Row, 1)
        'Cek Konversi
        OBJ1.Open dsn
        SQL1 = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            'Konversi balik ke satuan awal
            RST!qty = grid2.TextMatrix(grid2.Row, 4) / RST1!nilai
        Else
            RST!qty = grid2.TextMatrix(grid2.Row, 4)
        End If
        OBJ1.Close
        'RST!qty = grid2.TextMatrix(grid2.Row, 4)
        RST!kodesatuan = grid2.TextMatrix(grid2.Row, 6)
        RST!lineitem = grid2.Row
        RST.Update
        If grid2.Rows = grid2.Row + 1 Then Exit Do
        grid2.Row = grid2.Row + 1
    Loop
    
    'ambil kodestok
    kode_stok = GetKdStok
    
    'ambil total hpp kemasan
    OBJ1.Open dsn
    SQL1 = "Select noref,SUM(hpp)'hpp_totpack' From list_produksi_kemasan Where noref = '" & txtpalet & "' Group By noref"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        strhpppack = RST1!hpp_totpack
    End If
    OBJ1.Close
    
    'SIMPAN KE TABEL STOK (tabel baru)
    SQL = "Select * From am_stok Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!kdstok = kode_stok
        RST!tanggal = Format(Date, "yyyy/MM/dd")
        RST!nolot = txtlot
        RST!palet = txtpalet
        RST!type_trans = "SC"
        RST!noref = txtlot
        RST!gudang = "G3"
        RST!kodebarang = grid1.TextMatrix(grid1.Row, 1)
        RST!namabarang = grid1.TextMatrix(grid1.Row, 2)
        RST!kodeproduk = str1
        RST!kodesatuan = grid1.TextMatrix(grid1.Row, 5)
        RST!awal = "0"
        RST!qtyin = grid1.TextMatrix(grid1.Row, 4)
        RST!qtyout = "0"
        'Cek Konversi
        OBJ1.Open dsn
        SQL1 = "Select konversi From am_itemdtl Where KodeBarang = '" & grid1.TextMatrix(grid1.Row, 1) & "' and KodeSatuan = '" & grid1.TextMatrix(grid1.Row, 5) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            'Konversi balik ke satuan awal
            RST!isi = RST1!konversi
        Else
            RST!isi = "1"
        End If
        OBJ1.Close
        RST!kg = grid1.TextMatrix(grid1.Row, 6)
        RST!hpp = grid1.TextMatrix(grid1.Row, 7)
        RST!nosj = ""
        RST!kodecust = ""
        RST!UserName = nmuser
        RST!useredit = ""
        RST!tgledit = Format(Date, "yyyy/MM/dd")
        RST!baris = grid.Row
        RST!keterangan = "Hasil Produksi"
        RST!flag = "0"
        RST!hpp_totpack = strhpppack
        RST.Update
        If grid1.Rows = grid1.Row + 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop
    
    'KodeBarang WIP Jadi tidak masuk ke gudang
    If Left(grid1.TextMatrix(grid1.Row, 1), 3) = "L98" Then GoTo basewip:
    If Left(grid1.TextMatrix(grid1.Row, 1), 3) = "K98" Then GoTo basewip:
        
    strkorl = Left(grid1.TextMatrix(grid1.Row, 1), 1)   'K or L
    If strkorl = "K" Then GoTo karpet:
    
'LOT WIP HARUS SUDAH DI CECKLIST SEBELUM SCAN KALAU TIDAK MAKA AKAN MASUK KE STOK GUDANG PUSAT
'KALAU TERLANJUR MASUK MAKA HAPUS LOT di am_bpbhdr & am_bpblin
'Cek nolot base wip kalau ada lewati am_bpbhdr & am_bpblin
    SQL = "Select * from am_sopbase Where nolot='" & txtlot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo basewip:
karpet:
    If strkorl = "L" Then
        'AMBIL KODE PRODUKSI HARIAN LEM(PHL0-...)
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 4)
            strkode = RST!nobpb
        Else
            str99 = 0
            strkode = "0"
        End If
    
        str99 = str99 + 1
        
        If Len(str99) = 1 Then
            If IsNull(strkode) Or strkode = "0" Then
                strbpb = "PHL0-" & strformat & "000" & str99
            Else
                strbpb = "PHL0-" & strformat & Mid(strkode, 10, 3) & str99
            End If
        End If
        If Len(str99) = 2 Then strbpb = "PHL0-" & strformat & Mid(RST!nobpb, 10, 2) & str99
        If Len(str99) = 3 Then strbpb = "PHL0-" & strformat & Mid(RST!nobpb, 10, 1) & str99
        If Len(str99) = 4 Then strbpb = "PHL0-" & strformat & str99
    ElseIf strkorl = "K" Then
        'AMBIL KODE PRODUKSI HARIAN LEM(PHK0-...)
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHK0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 4)
            strkode = RST!nobpb
        Else
            str99 = 0
            strkode = "0"
        End If
    
        str99 = str99 + 1
        
        If Len(str99) = 1 Then
            If IsNull(strkode) Or strkode = "0" Then
                strbpb = "PHK0-" & strformat & "000" & str99
            Else
                strbpb = "PHK0-" & strformat & Mid(strkode, 10, 3) & str99
            End If
        End If
        If Len(str99) = 2 Then strbpb = "PHK0-" & strformat & Mid(RST!nobpb, 10, 2) & str99
        If Len(str99) = 3 Then strbpb = "PHK0-" & strformat & Mid(RST!nobpb, 10, 1) & str99
        If Len(str99) = 4 Then strbpb = "PHK0-" & strformat & str99
    End If
    
    'SIMPAN KE BPBHDR
    SQL = "Select * From am_bpbhdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !Type = "01"
        !nobpb = strbpb
        !tglbpb = Format(Date, "yyyy/MM/dd")
        !kodegudang = "G3"
        !keterangan = txtpalet
        !noref = txtpalet
        !identry = nmuser
        !dateentry = Format(Date, "yyyy/MM/dd")
        !idupdate = nmuser
        !dateupdate = Format(Date, "yyyy/MM/dd")
        .Update
    End With
    
    'SIMPAN KE BPBLIN
    SQL = "Select * From am_bpblin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        RST.AddNew
        RST!Type = "01"
        RST!nobpb = strbpb
        RST!tglbpb = Format(Date, "yyyy/MM/dd")
        RST!kodebarang = grid1.TextMatrix(grid1.Row, 1)
        RST!qty = grid1.TextMatrix(grid1.Row, 4)
        RST!keterangan = txtpalet
        RST!kodesatuan = grid1.TextMatrix(grid1.Row, 5)
        RST!lineitem = grid1.Row
        RST.Update
        If grid1.Rows = grid1.Row + 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop
    
basewip:
    MsgBox "Data berhasil disimpan.", vbInformation, AppName
    btnclear_Click
    OBJ.Close
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub Form_Load()
    'Dim strbpb As String
    'Dim strformat As String
    'strformat = Format(Date, "yymm")
    'OBJ.Open dsn
    'SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
    'Set RST = OBJ.Execute(SQL)
    'If Not RST.EOF Then
    '    str99 = Right(RST!nobpb, 4)
    'Else
    '    str99 = 0
    'End If
    
    'str99 = str99 + 1
    
    'If Len(str99) = 1 Then strbpb = "PHL0-" & strformat & Mid(RST!nobpb, 10, 3) & str99 '"000" & str99
    'If Len(str99) = 2 Then strbpb = "PHL0-" & strformat & Mid(RST!nobpb, 10, 2) & str99 '"00" & str99
    'If Len(str99) = 3 Then strbpb = "PHL0-" & strformat & Mid(RST!nobpb, 10, 1) & str99 '"0" & str99
    'If Len(str99) = 4 Then strbpb = "PHL0-" & strformat & str99
    'OBJ.Close

    'Exit Sub
'=================================
    initGrid
    setGrid
    package = False: kaleng = False: etiket = False
    str1 = "": str2 = "": str3 = "": str4 = ""
    txtnilai.Visible = False
    txtqty.Visible = False
End Sub

Private Sub grid_Click()
    Dim stokbahan As Double
    Dim hppbahan As Double
    Dim konvToKg As Double
    If grid.MouseRow = 0 Then Exit Sub
    If txtlot = "" Then Exit Sub
    
    Select Case grid.Col
        Case 0, 1, 2, 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then grid.Visible = False: Picgrid.Visible = False: Exit Sub
                If package = True Then
'===================PILIH KEMASAN
                    package = False
                    grid2.Row = 1
                    Do While True
                        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
                        grid2.Rows = grid2.Rows + 1
                        grid2.Row = grid2.Row + 1
                        grid2.Row = grid2.Rows - 1
                    Loop
                    grid2.Col = 0
                    Set grid2.CellPicture = uncheck
                    grid2.TextMatrix(grid2.Row, 1) = grid.TextMatrix(grid.Row, 1)
                    grid2.TextMatrix(grid2.Row, 2) = grid.TextMatrix(grid.Row, 2)
                    grid2.TextMatrix(grid2.Row, 3) = grid.TextMatrix(grid.Row, 4)

                    If kaleng = True Then
                        'KEMASAN CHILD
                        str4 = str4 * grid.TextMatrix(grid.Row, 4)
                        grid2.TextMatrix(grid2.Row, 4) = str4
                        grid2.Col = 4
                        grid2.CellBackColor = &H80FFFF
        
                        'cek konversi to kg unit
                        OBJ.Open dsn
                        SQL = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                        Set RST = OBJ.Execute(SQL)
                        If Not RST.EOF Then
                            konvToKg = Format(str4 / RST!nilai, "##,###,##0.00")
                        Else
                            konvToKg = str4
                        End If
                        OBJ.Close
                        
                        GetStokBarang Format(Date, "yyyyMMdd"), grid2.TextMatrix(grid2.Row, 1), , , stokbahan
                        'DISINI KONVERSI KE SATUAN DARI KG/ROLL
                        'MASALAH DI KEMASAN KILOAN & ROLLAN SEDANGKAN SYSTEM CEK SATUAN PCS
                        If stokbahan <= 0 Or stokbahan <= konvToKg Then
                            MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
                            Exit Sub
                        'Else
                            'CUMA TES
                            'MsgBox konvToKg & " > " & stokbahan
                        End If
                        
                        grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, konvToKg), "##,###,###,##0.00")
                        'grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, str4), "##,###,###,##0.00")
                    Else
                        'KEMASAN HEADER
                        str3 = str3 * grid.TextMatrix(grid.Row, 4)
                        GetStokBarang Format(Date, "yyyyMMdd"), grid2.TextMatrix(grid2.Row, 1), , , stokbahan
                        
                        'cek konversi to kg unit
                        OBJ.Open dsn
                        SQL = "Select * from am_apunit_konversi Where kdbrg ='" & grid2.TextMatrix(grid2.Row, 1) & "'"
                        Set RST = OBJ.Execute(SQL)
                        If Not RST.EOF Then
                            konvToKg = Format(str3 / RST!nilai, "##,###,##0.00")
                        Else
                            konvToKg = Format(str3, "##,###,##0.00")
                        End If
                        OBJ.Close

                        grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, konvToKg), "##,###,###,##0.00")
                        
                        grid2.Col = 4
                        txtqty.Width = grid2.ColWidth(grid2.Col) - 40
                        txtqty = grid2.TextMatrix(grid2.Row, grid2.Col)
                        txtqty.Left = grid2.Left + grid2.CellLeft
                        txtqty.Top = grid2.Top + grid2.CellTop + 20
                        txtqty.Visible = True
                        txtqty.SetFocus
                    End If
                    grid2.TextMatrix(grid2.Row, 5) = grid.TextMatrix(grid.Row, 3)
                    grid2.TextMatrix(grid2.Row, 6) = grid.TextMatrix(grid.Row, 5)
                    If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
                    TabControl.SelectedItem = 1
                Else
'===================PILIH PRODUK (BARANG JADI)
                    grid1.Row = 1
                    grid1.Col = 0
                    Set grid1.CellPicture = uncheck
                    grid1.TextMatrix(grid1.Row, 1) = grid.TextMatrix(grid.Row, 1)
                    grid1.TextMatrix(grid1.Row, 2) = grid.TextMatrix(grid.Row, 2)
                    grid1.TextMatrix(grid1.Row, 3) = grid.TextMatrix(grid.Row, 4)
                    grid1.TextMatrix(grid1.Row, 5) = grid.TextMatrix(grid.Row, 3)
                    'Ambil Hpp produk
                    OBJ.Open dsn
                    SQL = "Select nolot,SUM(hpp)/SUM(qty_bahan)'perkg' From list_produksi_child "
                    SQL = SQL + "Where nolot = '" & txtlot & "' group by nolot"
                    Set RST = OBJ.Execute(SQL)
                    If RST.EOF Then
                        OBJ.Close
                        MsgBox "No Lot tidak ditemukan", vbCritical, AppName
                        Exit Sub
                    Else
                        grid1.TextMatrix(grid1.Row, 7) = Format(RST!perkg, "##,###,##0.00")
                        OBJ.Close
                    End If
                    
                    'Ambil Kg Timbangan Produk
                    
                    OBJ.Open dsn
                    SQL = "Select * From am_itemkg where tahun = '" & 20 & Left(txtlot, 2) & "'"
                    SQL = SQL + " and kodebarang = '" & grid1.TextMatrix(grid1.Row, 1) & "' and kodesatuan = '" & grid1.TextMatrix(grid1.Row, 5) & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Mid(txtlot, 3, 1) = "A" Or Mid(txtlot, 3, 2) = "01" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg1, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "B" Or Mid(txtlot, 3, 2) = "02" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg2, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "C" Or Mid(txtlot, 3, 2) = "03" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg3, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "D" Or Mid(txtlot, 3, 2) = "04" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg4, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "E" Or Mid(txtlot, 3, 2) = "05" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg5, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "F" Or Mid(txtlot, 3, 2) = "06" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg6, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "G" Or Mid(txtlot, 3, 2) = "07" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg7, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "H" Or Mid(txtlot, 3, 2) = "08" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg8, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "J" Or Mid(txtlot, 3, 2) = "09" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg9, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "K" Or Mid(txtlot, 3, 2) = "10" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg10, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "L" Or Mid(txtlot, 3, 2) = "11" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg11, "##,##0.00")
                        End If
                    ElseIf Mid(txtlot, 3, 1) = "M" Or Mid(txtlot, 3, 2) = "12" Then
                        If RST.EOF Then
                            MsgBox "Timbangan produk (Kg base unit) belum diisi", vbCritical, AppName
                            OBJ.Close
                            Exit Sub
                        Else
                            grid1.TextMatrix(grid1.Row, 6) = Format(RST!kg12, "##,##0.00")
                        End If
                    Else
                        MsgBox "Format No Lot salah..!", vbCritical, AppName
                        OBJ.Close
                        Exit Sub
                    End If
wipjadi:
                    OBJ.Close
                    grid1.Col = 4
                        
                    If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
                
                    txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
                    txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
                    txtnilai.Left = grid1.Left + grid1.CellLeft
                    txtnilai.Top = grid1.Top + grid1.CellTop + 20
                    txtnilai.Visible = True
                    txtnilai.SetFocus
                    TabControl.SelectedItem = 0
                End If
                Picgrid.Visible = False
                grid.Visible = False
    End Select
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    If showmode = True Then Exit Sub
    Select Case grid1.Col
        Case 2
            If grid1.TextMatrix(1, 1) = "" Then Exit Sub
            If grid1.TextMatrix(1, 4) = "" Then Exit Sub
            Picgrid.Visible = True
            grid.Visible = True
            str2 = grid1.TextMatrix(grid1.Row, 1)
            hapusgrid
            package = True
            
            OBJ.Open dsn
            SQL = "Select a.kode_kemasan,b.NamaBarang,b.kodesatuan,a.lev,a.konversi,a.id,a.id_root From list_konversilevel a"
            SQL = SQL + " inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang"
            SQL = SQL + " Where a.kode_barang_jadi = '" & str2 & "' and a.id_root = ''"
            Set RST = OBJ.Execute(SQL)
            initGrid
            grid.Row = 1
            Do While Not RST.EOF
                grid.TextMatrix(grid.Row, 1) = RST!kode_kemasan
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = RST!Id
                grid.TextMatrix(grid.Row, 4) = RST!konversi
                grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                grid.Col = 0
                Set grid.CellPicture = uncheck
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                RST.MoveNext
            Loop
            OBJ.Close
            txtinfo = "Pilih Kemasan" & vbCrLf & "Produk : " & grid1.TextMatrix(grid1.Row, 2)
        Case 4
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid1.ColWidth(grid1.Col) - 40
            txtnilai = grid1.TextMatrix(grid1.Row, grid1.Col)
            txtnilai.Left = grid1.Left + grid1.CellLeft
            txtnilai.Top = grid1.Top + grid1.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If showmode = True Then Exit Sub
    Select Case grid2.Col
        Case 0
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
            End If
        Case 2, 3
            If grid2.TextMatrix(1, 1) = "" Then Exit Sub
            Picgrid.Visible = True
            grid.Visible = True
            str3 = grid2.TextMatrix(grid2.Row, 5)
            str4 = grid2.TextMatrix(grid2.Row, 4)
            txtinfo = "Pilih konversi kemasan" & vbCrLf & "Packaging : " & grid2.TextMatrix(grid2.Row, 2)
            
            hapusgrid
            package = True: kaleng = True
            OBJ.Open dsn
            SQL = "Select a.kode_kemasan,b.NamaBarang,b.kodesatuan,a.lev,a.konversi,a.id,a.id_root From list_konversilevel a"
            SQL = SQL + " inner join am_apitemmst b on a.kode_kemasan=b.KodeBarang"
            SQL = SQL + " Where a.kode_barang_jadi = '" & str2 & "' and a.id_root = '" & str3 & "'"
            Set RST = OBJ.Execute(SQL)
            
            initGrid
            grid.Row = 1
            Do While Not RST.EOF
                grid.TextMatrix(grid.Row, 1) = RST!kode_kemasan
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = RST!Id
                grid.TextMatrix(grid.Row, 4) = RST!konversi
                grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                grid.Col = 0
                Set grid.CellPicture = uncheck
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
                RST.MoveNext
            Loop
            OBJ.Close
        Case 4:
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            
            txtqty.Width = grid2.ColWidth(grid2.Col) - 40
            txtqty = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtqty.Left = grid2.Left + grid2.CellLeft
            txtqty.Top = grid2.Top + grid2.CellTop + 20
            txtqty.Visible = True
            txtqty.SetFocus
    End Select
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
txtnilai.Visible = False
txtqty.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case grid1.Col
            Case 4:
                txtnilai.Visible = False
                grid1.TextMatrix(grid1.Row, 4) = txtnilai
                etiket = True
                str3 = grid1.TextMatrix(grid1.Row, 4)
        End Select
        
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtpalet_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        showmode = False
        'CEK PALET
        OBJ.Open dsn
        SQL = "Select distinct a.*,c.NamaBarang,d.NamaSatuan From list_mutasi_produksi_details a inner join list_produk_hasil b on a.kode_barang = b.kode_barang_jadi"
        SQL = SQL + " inner join am_itemmst c on a.kode_barang = c.KodeBarang"
        SQL = SQL + " inner join am_unit d on b.kode_satuan = d.KodeSatuan"
        SQL = SQL + " Where a.kode_palet ='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        hapusgrid1
        hapusgrid2
        If Not RST.EOF Then
            'BARANG SUDAH SCAN
            txtlot = Mid(txtpalet, 3, 20)
            showmode = True
            txtnilai.Visible = False
            txtqty.Visible = False
            
            'MENAMPILKAN DATA BARANG JADI (PEROLEHAN)
            grid1.Row = 1
            Do While Not RST.EOF
                grid1.Col = 0
                Set grid1.CellPicture = uncheck
                grid1.TextMatrix(grid1.Row, 1) = RST!kode_barang
                grid1.TextMatrix(grid1.Row, 2) = RST!namabarang
                grid1.TextMatrix(grid1.Row, 3) = RST!namasatuan
                grid1.TextMatrix(grid1.Row, 4) = RST!qty
                RST.MoveNext
            Loop

            'MENAMPILKAN DATA PEMAKAIAN KEMASAN
            SQL = "Select a.*,b.NamaBarang From list_produksi_kemasan a"
            SQL = SQL + " inner join am_apitemmst b on a.kode_bahan = b.KodeBarang"
            SQL = SQL + " where a.noref = '" & txtpalet & "' Order by a.qty_bahan asc"
            Set RST = OBJ.Execute(SQL)
        
            grid2.Row = 1
            Do While Not RST.EOF
                grid2.Col = 0
                Set grid2.CellPicture = uncheck
                grid2.TextMatrix(grid2.Row, 1) = RST!kode_bahan
                grid2.TextMatrix(grid2.Row, 2) = RST!namabarang
                grid2.TextMatrix(grid2.Row, 3) = "" 'RST!konversi
                grid2.TextMatrix(grid2.Row, 4) = RST!qty_bahan
                grid2.Col = 4
                grid2.CellBackColor = &H80FFFF
                grid2.TextMatrix(grid2.Row, 5) = "" 'RST!Id
                grid2.TextMatrix(grid2.Row, 6) = RST!KODE_SATUAN
                grid2.TextMatrix(grid2.Row, 7) = RST!hpp
                RST.MoveNext
                grid2.Rows = grid2.Rows + 1
                grid2.Row = grid2.Row + 1
                grid2.Row = grid2.Rows - 1
            Loop
            OBJ.Close
            Exit Sub
        End If
        

        'BELUM SCAN
        grid.Visible = True
        txtnilai.Visible = False
        txtqty.Visible = False
        
        SQL = "Select kode_produk From list_produksi_master Where nolot='" & Mid(txtpalet, 3, 20) & "'"
        Set RST = OBJ.Execute(SQL)
        
        txtlot = Mid(txtpalet, 3, 20)
        If Not RST.EOF Then str1 = RST!kode_produk
        txtinfo = "Pilih barang jadi" & vbCrLf & "Nomor Lot : " & txtlot
        
        SQL = "select distinct a.kodebarang, a.namabarang,a.KodeSatuan,c.namasatuan"
        SQL = SQL + " from am_itemdtl a inner join list_produk_hasil b on a.kodebarang=b.kode_barang_jadi"
        SQL = SQL + " inner join am_unit c on a.KodeSatuan = c.KodeSatuan"
        SQL = SQL + " and b.kode_produk='" & str1 & "' order by a.KodeBarang asc"
        Set RST = OBJ.Execute(SQL)
        
        hapusgrid
        Picgrid.Visible = True
        grid.Row = 1
        Do While Not RST.EOF
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            grid.TextMatrix(grid.Row, 3) = RST!kodesatuan
            grid.TextMatrix(grid.Row, 4) = RST!namasatuan
            grid.Col = 0
            Set grid.CellPicture = uncheck
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
        OBJ.Close
    End If
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
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub
Private Sub hapusgrid1()
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""
        grid1.TextMatrix(grid1.Row, 2) = ""
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1.TextMatrix(grid1.Row, 4) = ""
        grid1.TextMatrix(grid1.Row, 5) = ""
        
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.TextMatrix(grid1.Row, 7) = ""
        grid1.Col = 0
        Set grid1.CellPicture = blank
        If grid1.Row = 1 Then Exit Do
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
    txtnilai.Value = ""
    setGrid
End Sub
Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Col = 0
        Set grid2.CellPicture = blank
        If grid2.Row = 1 Then Exit Do
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    setGrid
End Sub
Private Sub setGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 900
        .ColWidth(2) = 3500
        .ColWidth(3) = 0
        .ColWidth(4) = 900
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
    With grid1
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 3500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 0
        .ColWidth(6) = 0 '2000
        .ColWidth(7) = 0 '2000
    End With
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 4500
        .ColWidth(3) = 0
        .ColWidth(4) = 1100
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
End Sub

Private Sub initGrid()
    If kaleng = True Or package = True Then
        With grid
            .Cols = 7
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "KODE"
            .TextMatrix(0, 2) = "NAMA BARANG"
            .TextMatrix(0, 3) = "KODE"
            .TextMatrix(0, 4) = "KONVERSI"
        End With
    Else
        With grid
            .Cols = 5
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "KODE"
            .TextMatrix(0, 2) = "NAMA BARANG"
            .TextMatrix(0, 3) = "KODE"
            .TextMatrix(0, 4) = "SATUAN"
        End With
    End If
    With grid1
        .Cols = 8 '6
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "SATUAN"
        .TextMatrix(0, 4) = "QTY"
        .Col = 4
        .CellBackColor = &H80FFFF
        '
        .TextMatrix(0, 5) = "KdSat"
        .TextMatrix(0, 6) = "kg"
        .TextMatrix(0, 7) = "hpp"
    End With
    With grid2
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "SATUAN"
        .TextMatrix(0, 4) = "QTY"
        .Col = 4
        .CellBackColor = &H80FFFF
        .TextMatrix(0, 5) = "ID"
    End With
End Sub
Private Sub hapusrow()
    If grid2.Rows = 2 Then
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.Rows = grid2.Rows - 1
        grid2.Col = 0
        Set grid2.CellPicture = blank
        If grid2.Rows = 1 Then grid2.Rows = grid2.Rows + 1: package = False: str4 = grid1.TextMatrix(grid1.Row, 4)
        Exit Sub
    Else
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
    End If
    Do While True
        If grid2.Row = grid2.Rows - 1 Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = grid2.TextMatrix(grid2.Row + 1, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row + 1, 2)
        grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row + 1, 3)
        grid2.TextMatrix(grid2.Row, 4) = grid2.TextMatrix(grid2.Row + 1, 4)
        grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row + 1, 5)
        grid2.TextMatrix(grid2.Row, 6) = grid2.TextMatrix(grid2.Row + 1, 6)
        grid2.TextMatrix(grid2.Row, 7) = grid2.TextMatrix(grid2.Row + 1, 7)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = uncheck
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    If KeyAscii = 13 Then
        Select Case grid2.Col
            Case 4:
                'CEK STOK
                GetStokBarang Format(Date, "yyyyMMdd"), grid2.TextMatrix(grid2.Row, 1), , , stokbahan
                If stokbahan <= 0 Or stokbahan < txtqty.Value Then
                    MsgBox "Stok tidak mencukupi...! stok terakhir : " & stokbahan, vbCritical, AppName
                    Exit Sub
                'Else
                    'UNTUK TES
                    'MsgBox stokbahan & " <= " & txtqty.Value
                End If

                txtqty.Visible = False
                grid2.TextMatrix(grid2.Row, 4) = txtqty
                'update 10/11/2022
                grid2.TextMatrix(grid2.Row, 7) = Format(getHPP(grid2.TextMatrix(grid2.Row, 1), stokbahan, txtqty.Value), "##,###,###,##0.00")
        End Select
    End If
End Sub

Private Sub txtqty_LostFocus()
    txtqty.Visible = False
End Sub

Function getnomut() As String    '2016060001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yyyymm")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(lotstok)as kr from am_stoklot where lotstok like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getnomut = strnumber + "0001"
    Else
        kode = CLng(Mid(RST!kr, 7, 4)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "000" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 4) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getnomut = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getnomut = strnumber + "0001"
End Function
