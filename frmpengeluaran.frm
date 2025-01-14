VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpengeluaran 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT BUKTI PENGELUARAN UANG"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdnomor 
      Height          =   300
      Left            =   6540
      TabIndex        =   16
      Top             =   480
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Nomor"
      FlatStyle       =   -1  'True
      TextAlignment   =   1
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   330
      Left            =   8175
      TabIndex        =   13
      Top             =   3825
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtuntuk 
      Appearance      =   0  'Flat
      Height          =   1260
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1500
      Width           =   7575
   End
   Begin VB.TextBox txtdibayar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   6
      Top             =   1170
      Width           =   7575
   End
   Begin VB.TextBox txtnobkt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7500
      MaxLength       =   6
      TabIndex        =   4
      Top             =   480
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   330
      Left            =   7500
      TabIndex        =   3
      Top             =   60
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
      _Version        =   393216
      Format          =   134610945
      CurrentDate     =   41743
   End
   Begin TDBNumber6Ctl.TDBNumber txtncash 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   3105
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calculator      =   "frmpengeluaran.frx":0000
      Caption         =   "frmpengeluaran.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpengeluaran.frx":008C
      Keys            =   "frmpengeluaran.frx":00AA
      Spin            =   "frmpengeluaran.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   1
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
      ValueVT         =   1638405
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin XtremeSuiteControls.PushButton cmddelete 
      Height          =   330
      Left            =   7170
      TabIndex        =   14
      Top             =   3825
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Delete"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   330
      Left            =   6180
      TabIndex        =   15
      Top             =   3825
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Simpan"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   330
      Left            =   5160
      TabIndex        =   17
      Top             =   3825
      Width           =   960
      _Version        =   851970
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Clear"
      UseVisualStyle  =   -1  'True
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   3735
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   -75
      Top             =   3675
      Width           =   10485
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Cek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label lblterbilang 
      BackStyle       =   0  'Transparent
      Caption         =   "Nol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1545
      TabIndex        =   10
      Top             =   2820
      Width           =   7590
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Terbilang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk/Ket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1545
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dibayar pada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal "
      Height          =   240
      Left            =   6735
      TabIndex        =   2
      Top             =   165
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUKTI PENGELUARAN KAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2655
      TabIndex        =   1
      Top             =   735
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PT. SPARTA PRIMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2595
   End
End
Attribute VB_Name = "frmpengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As New ADODB.Recordset
Private SQL As String
Private posedit As Boolean

Private Sub cmdclear_Click()
    txtnobkt = GetNewNumber
    date1.Value = Date
    txtdibayar = ""
    txtuntuk = ""
    txtncash = "0.00"
    lblterbilang = "Nol"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub
 
Private Function GetNewNumber() As String
    On Error GoTo err_handler
    Dim tempnumber As Long
    Dim nobkt As String
    Dim lengthnumber As Integer
    SQL = "select max(nobkt) maxnobkt from gl_pengeluaran"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    tempnumber = CLng(RST!maxnobkt) + 1
    lengthnumber = Len(Trim(Str(tempnumber)))
    txtnobkt = GetNewNumber
    Select Case lengthnumber
        Case 1: nobkt = "00000" + Trim(Str(tempnumber))
        Case 2: nobkt = "0000" + Trim(Str(tempnumber))
        Case 3: nobkt = "000" + Trim(Str(tempnumber))
        Case 4: nobkt = "00" + Trim(Str(tempnumber))
        Case 5: nobkt = "0" + Trim(Str(tempnumber))
        Case 6: nobkt = Trim(Str(tempnumber))
    End Select
    
    GetNewNumber = nobkt

    OBJ.Close
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    Exit Function
End Function

Private Sub cmddelete_Click()
    On Error GoTo err_handler:
    If MsgBox("Apakah anda yakin akan menghapus data tersebut...?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "delete form gl_pengeluaran where nobkt='" & txtnobkt & "'"
    OBJ.Execute SQL
    MsgBox "Proses hapus berhasil...!", vbInformation, AppName
    OBJ.Close
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err_handler:
    If posedit = False Then
        txtnobkt = GetNewNumber
    End If
    
    OBJ.Open dsn
    SQL = "delete from gl_pengeluaran where nobkt='" & txtnobkt & "'"
    OBJ.Execute SQL
    
    SQL = "select * from gl_pengeluaran where 0=1"
    
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nobkt = txtnobkt
        !tgl = Format(date1, "MM/dd/yyyy")
        !kepada = txtdibayar
        !keterangan = txtuntuk
        !perkiraan = ""
        !terbilang = lblterbilang
        !jumlah = Format(txtncash, "general number")
        .Update
    End With
    OBJ.Close
    MsgBox "Proses simpan berhasil...!", vbInformation, AppName
    CetakBukti
    posedit = False
    cmdclear_Click
    txtdibayar.SetFocus
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox "Proses simpan tidak berhasil...!", vbCritical, AppName
    Exit Sub
End Sub

Private Sub CetakBukti()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(gl_bukti_pengeluaran)"
    Crystal.ReportFileName = AppPath & "\reports\gl\ledger\bukti_pengeluaran.rpt"
    Crystal.ParameterFields(0) = "@nobukti;" + txtnobkt + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    txtnobkt = GetNewNumber
    date1.Value = Date
End Sub

Private Sub txtncash_Change()
    If txtncash = "" Then Exit Sub
    lblterbilang = ANGKAKEHURUF(Format(txtncash, "general number")) & " Rupiah"
End Sub

Private Sub txtnobkt_KeyPress(KeyAscii As Integer)
    If txtnobkt = "" Then Exit Sub
    If KeyAscii = 13 Then
        CariData
    End If
End Sub

Private Sub CariData()
    On Error GoTo err_handler:
    SQL = "select * from gl_pengeluaran where nobkt='" & txtnobkt & "'"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Data tidak ditemukan...!", vbInformation, AppName
        Exit Sub
    End If
    
    txtdibayar = RST!kepada
    txtuntuk = RST!keterangan
    lblterbilang = RST!terbilang
    txtncash = RST!jumlah
    posedit = True
    OBJ.Close
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description
End Sub


