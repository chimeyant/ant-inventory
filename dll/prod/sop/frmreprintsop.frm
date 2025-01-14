VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmreprintsop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Ulang SOP"
   ClientHeight    =   2115
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcari 
      Caption         =   "CARI"
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
      Left            =   3825
      TabIndex        =   6
      Top             =   60
      Width           =   960
   End
   Begin VB.TextBox txtket 
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
      Height          =   540
      Left            =   1335
      TabIndex        =   1
      Top             =   405
      Width           =   3405
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   75
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtnolot 
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
      Height          =   270
      Left            =   1350
      TabIndex        =   0
      Top             =   75
      Width           =   2430
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.PushButton cmdcetak 
      Height          =   465
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Cetak"
      BackColor       =   -2147483633
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   4680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "*  Jika di hasil cetak SPK tidak menampilkan daftar barang jadinya, Silahkan periksa Kg base unit."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "KETERANGAN :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO LOT :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   585
      TabIndex        =   2
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmreprintsop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmdcari_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    carisql1 = carisql1 + " where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdcetak_Click()
    cetaksop
End Sub

Private Sub cetaksop()
    On Error GoTo Err_handler:
    Dim cetak_ke As Integer
    Dim produk As String
    
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Akses ditolak...! ", vbCritical, AppName
        Exit Sub
    End If
proses:
    OBJ.Open dsn
    SQL = "Select nolot from list_historicetaksop where nolot ='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Nolot tidak ditemukan...! ", vbCritical, AppName
        Exit Sub
    End If

    SQL = "Select count(nolot)as jml from list_historicetaksop where nolot ='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    cetak_ke = RST!jml
    cetak_ke = cetak_ke + 1
    
    SQL = "select * from list_historicetaksop where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot
        !tanggal = Format(Date, "yyyy/MM/dd")
        !cetakan = cetak_ke
        !keterangan = txtket
        !UserName = nmuser
        .Update
    End With
    
    SQL = "Select LEFT(kode_produk,1)'produk' from list_produksi_master where nolot = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then produk = RST!produk
    OBJ.Close
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetaksop1)"
    
    If produk = "L" Then crystal.ReportFileName = AppPath & "\reports\produksi\cetak_sop.rpt"
    If produk = "K" Then crystal.ReportFileName = AppPath & "\reports\produksi\cetak_sopk.rpt"
        'Jika formula tidak sesuai dengan master u/ print pakai ini
        'crystal.DataFiles(0) = "Proc(am_cetaksop1)"
        'crystal.ReportFileName = AppPath & "\reports\produksi\cetak_sop_nomaster.rpt"
    crystal.ParameterFields(0) = "@nolot;" & txtnolot.text & ";true"
    crystal.ParameterFields(1) = "@username;" & nmuser & ";true"
    crystal.ParameterFields(2) = "@cetakan;" & "CETAKAN KE " & Str(cetak_ke) & ";true"
    crystal.ParameterFields(3) = "@kode;" & Cheap_Decrypt(txtnolot) & Cheap_Decrypt(Str(Trim(cetak_ke))) & ";true"
    crystal.ParameterFields(4) = "@nolot2;" & txtnolot.text & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    txtnolot = ""
    txtket = ""
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

