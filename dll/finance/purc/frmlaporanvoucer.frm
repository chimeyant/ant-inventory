VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmlaporanvoucer 
   Caption         =   "Laporan Voucher"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   2280
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
      MICON           =   "frmlaporanvoucer.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
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
      MICON           =   "frmlaporanvoucer.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2100
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4740
      _Version        =   851970
      _ExtentX        =   8361
      _ExtentY        =   3704
      _StockProps     =   79
      Caption         =   "Filter By"
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
      Begin VB.Frame Frame1 
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   150
         TabIndex        =   5
         Top             =   990
         Width           =   4500
         Begin MSComCtl2.DTPicker date1 
            Height          =   315
            Left            =   885
            TabIndex        =   6
            Top             =   240
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
            Format          =   122355713
            CurrentDate     =   41610
         End
         Begin MSComCtl2.DTPicker date2 
            Height          =   315
            Left            =   885
            TabIndex        =   7
            Top             =   570
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
            Format          =   122355713
            CurrentDate     =   41610
         End
         Begin VB.Label Label2 
            Caption         =   "s.d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   645
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Dari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   8
            Top             =   315
            Width           =   645
         End
      End
      Begin VB.OptionButton optstandar 
         Caption         =   "Standar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   3915
      End
      Begin VB.OptionButton optdetail 
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   3
         Top             =   285
         Value           =   -1  'True
         Width           =   3915
      End
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   2400
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   1980
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSForms.CheckBox chkexcel 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "0"
      Caption         =   "Export Excel"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblExport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Harap tunggu sebentar  proses export data sedang berjalan...!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   4980
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmlaporanvoucer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SQL As String
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
Dim oXLS As ActiveReportsExcelExport.ARExportExcel


Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    On Error GoTo Err_handler:
    If date1 > date2 Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If optdetail = True Then
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_voucher_detail)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\voucher_detail.rpt"
        Crystal.ParameterFields(0) = "@tgl1;" + Format(date1, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@tgl2;" + Format(date2, "yyyyMMdd") + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    
        Exit Sub
        'PAKAI ACTIVE REPORT (YANG LAMA)
        SQL = "select a.novoucher,a.tgl,a.kepada,a.npwp,a.alamat,a.kdkurs,a.nilai,a.ppn,SUM(b.jumlah)as total "
        SQL = SQL & "from am_voucherhdr as a inner join am_voucherin as b on a.novoucher=b.novoucher "
        SQL = SQL & "where a.tgl >='" & Format(date1, "MM/dd/yyyy") & "' and a.tgl <='" & Format(date2, "MM/dd/yyyy") & "' "
        SQL = SQL & "Group By a.novoucher,a.tgl,a.kepada,a.npwp,a.alamat,a.kdkurs,a.nilai,a.ppn "

        With rptlapvoucher
            .DataControl1.Source = SQL
            .DataControl1.ConnectionString = dsn
            .lbltgl = "Dari Tanggal : " & Format(date1, "dd/MM/yyyy") & " s.d " & Format(date2, "dd/MM/yyyy") & ""
            If chkexcel.Value = True Then
                With cmndlg
                    .CancelError = False
                    .DialogTitle = "File DPT"
                    .Filter = "MS Execel 2007 (*.xls)|*.xls"
                .ShowSave
                    If .FileName <> "" Then
                        lblExport.Visible = True
                        DoEvents
                        Set oXLS = New ActiveReportsExcelExport.ARExportExcel
                        oXLS.FileName = cmndlg.FileName
                        rptlapvoucher.Run
                        oXLS.Export rptlapvoucher.Pages
                        MsgBox "Proses Export data berhasil...!", vbInformation, AppName
                        lblExport.Visible = False
                    Else
                        MsgBox "Nama file belum diisi"
                        Exit Sub
                    End If
                End With
            Else
                .Show 'vbModal
            End If
        End With
    End If
    
    If optstandar.Value = True Then
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_voucher_header)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\voucher_header.rpt"
        Crystal.ParameterFields(0) = "@tgl1;" + Format(date1, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@tgl2;" + Format(date2, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(2) = "@user;" + nmuser + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
    
        Exit Sub
        'PAKAI ACTIVE REPORT (YANG LAMA)
        SQL = "select a.novoucher,a.tgl,a.kepada,a.npwp,a.alamat,a.kdkurs,a.ppn,SUM(b.jumlah)as total "
        SQL = SQL & "from am_voucherhdr as a inner join am_voucherin as b on a.novoucher=b.novoucher "
        SQL = SQL & "where a.tgl >='" & Format(date1, "MM/dd/yyyy") & "' and a.tgl <='" & Format(date2, "MM/dd/yyyy") & "' "
        SQL = SQL & "Group By a.novoucher,a.tgl,a.kepada,a.npwp,a.alamat,a.kdkurs,a.ppn "
        SQL = SQL & "ORDER By a.novoucher Asc "
        
        With rptlapvoucherhdr
            .DataControl1.Source = SQL
            .DataControl1.ConnectionString = dsn
            .lbltgl = "Dari Tanggal : " & Format(date1, "dd/MM/yyyy") & " s.d " & Format(date2, "dd/MM/yyyy") & ""
            If chkexcel.Value = True Then
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
                        rptlapvoucherhdr.Run
                        oXLS.Export rptlapvoucherhdr.Pages
                        MsgBox "Proses Export data berhasil...!", vbInformation, AppName
                        lblExport.Visible = False
                    Else
                        MsgBox "Nama file belum diisi"
                        Exit Sub
                    End If
                End With
            Else
                .Show 'vbModal
            End If
        End With
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub
