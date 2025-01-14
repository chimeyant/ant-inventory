VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmlaporanmutasi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Stock"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   615
      Top             =   1545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkexcel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export To Excel"
      Height          =   270
      Left            =   1650
      TabIndex        =   13
      Top             =   330
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Posisi Stock dan Saldo"
      Height          =   195
      Left            =   1635
      TabIndex        =   12
      Top             =   135
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mutasi Stock"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Posisi Stock"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtkode2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   144179203
      CurrentDate     =   38679
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   90
      Top             =   1575
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1680
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
      MICON           =   "frmlaporanmutasi.frx":0000
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
      Left            =   4095
      TabIndex        =   6
      Top             =   1680
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
      MICON           =   "frmlaporanmutasi.frx":031A
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
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Kode"
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
      MICON           =   "frmlaporanmutasi.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   144179203
      CurrentDate     =   38679
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Kode"
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
      MICON           =   "frmlaporanmutasi.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblExport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Harap tunggu sebentar  proses export data sedang berjalan...!"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -30
      TabIndex        =   14
      Top             =   2190
      Visible         =   0   'False
      Width           =   6180
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "From Date                                                     To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1230
      Width           =   3735
   End
End
Attribute VB_Name = "frmlaporanmutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim param As ADODB.Parameter
Dim SQL As String

Dim oPDF As ActiveReportsPDFExport.ARExportPDF
Dim oXLS As ActiveReportsExcelExport.ARExportExcel
Dim str2 As String

Private Sub cmdsearch1_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
        
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    If txtkode1 = "" Or txtkode2 = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtkode1 > txtkode2 Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date1 > date2 Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Option3.Value = True Then
        SQL = "Exec am_posisi '" & Format(date1, "yyyyMMdd") & "','" & Format(date2, "yyyyMMdd") & "','"
        SQL = SQL + txtkode1 & "','" & txtkode2 & "','" & UserOnline & "'"
    
        With rptlaporansaldobahanbaku
            .DataControl1.Source = SQL
            .DataControl1.ConnectionString = dsn
            If chkexcel.Value = 1 Then
                With cmndlg
                    .CancelError = False
                    .DialogTitle = "File DPT"
                    .Filter = "MS Execel 2003 (*.xls)|*.xls"
                .ShowSave
                    If .FileName <> "" Then
                        lblExport.Visible = True
                        DoEvents
                        Set oXLS = New ActiveReportsExcelExport.ARExportExcel
                        oXLS.FileName = cmndlg.FileName
                        rptlaporansaldobahanbaku.Run
                        oXLS.Export rptlaporansaldobahanbaku.Pages
                        MsgBox "Proses Export data berhasil...!", vbInformation, AppName
                        lblExport.Visible = False
                    Else
                        MsgBox "Nama file belum diisi"
                        Exit Sub
                    End If
                End With
            Else
                .Show vbModal
            End If
        End With
        Exit Sub
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    
    If Option1.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_mutasi)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\mut\mutasi.rpt"
    Else
        Crystal.DataFiles(0) = "Proc(am_posisi)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\mut\posisi.rpt"
    End If
    
    Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(2) = "@kode11;" + txtkode1 + ";true"
    Crystal.ParameterFields(3) = "@kode12;" + txtkode2 + ";true"
    Crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    OBJ.Open dsn
    SQL = "select top 1 * from am_invloc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1.MinDate = RST!tglupdate
        date2.MinDate = RST!tglupdate
    End If
    OBJ.Close
    
    If date1.MinDate > Date Then
        date1 = date1.MinDate
        date2 = date1.MinDate
    Else
        date1 = Date
        date2 = Date
    End If
End Sub

Private Sub txtkode1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtkode2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkode2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
