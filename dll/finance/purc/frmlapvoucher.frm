VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmlapvoucher 
   Caption         =   "Laporan Voucher (Processed/Unprocessed)"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   2010
      TabIndex        =   6
      Top             =   0
      Width           =   2790
      Begin VB.TextBox txtvoucher 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   8
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtvoucher2 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   7
         Top             =   615
         Width           =   1215
      End
      Begin XtremeSuiteControls.PushButton cmdvoucher 
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   195
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "From Voucher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
      End
      Begin XtremeSuiteControls.PushButton cmdvoucher2 
         Height          =   345
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "To Voucher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
      End
   End
   Begin XtremeSuiteControls.PushButton btnclose 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "View"
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
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   2880
      Top             =   885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   2415
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin MSForms.CheckBox chkexcel 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1695
      BackColor       =   8421504
      ForeColor       =   16777215
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "0"
      Caption         =   "Export Excel"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5175
   End
   Begin MSForms.OptionButton optv_hutang 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   1545
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2725;661"
      Value           =   "0"
      Caption         =   "voucher hutang"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optv_biaya 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1230
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2170;661"
      Value           =   "1"
      Caption         =   "All Voucher"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmlapvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset

Dim SQL As String
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
Dim oXLS As ActiveReportsExcelExport.ARExportExcel

Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub btnview_Click()
    If txtvoucher = "" And txtvoucher2 = "" Then Exit Sub
        If txtvoucher > txtvoucher2 Then
            MsgBox "The first voucher is greater than the second voucher", vbExclamation, "Warning"
        Exit Sub
    End If
    If optv_biaya = True Then
        viewvbiaya
    ElseIf optv_hutang = True Then
        viewvhutang
    End If
End Sub

Private Sub cmdvoucher_Click()
On Error GoTo Err_handler:
    If optv_hutang = True Then
        carisql1 = "Select distinct a.Ref1,a.NoBeli,b.tgl,c.NamaSupp From am_beliapp a inner join am_voucherhdr b "
        carisql1 = carisql1 + "on a.Ref1 = b.novoucher inner join am_supplier c on a.Kodesupp = c.KodeSupp"
        namatabel = "Voucher"
        frmsearch.Show vbModal
    ElseIf optv_biaya = True Then
        carisql1 = "Select a.novoucher,a.tgl,a.kepada,SUM(b.jumlah)as Jumlah From am_voucherhdr a "
        carisql1 = carisql1 + "inner join am_voucherin b on a.novoucher = b.novoucher"
        carisql1 = carisql1 + " group by a.novoucher,a.tgl,a.kepada order by a.novoucher"
        namatabel = "Allvoucher"
        frmsearch.Show vbModal
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdvoucher_GotFocus()
    If hasil1 = "" Then Exit Sub
    txtvoucher = hasil1
    hasil1 = ""
End Sub

Private Sub cmdvoucher2_Click()
On Error GoTo Err_handler:
    If optv_hutang = True Then
        carisql1 = "Select distinct a.Ref1,a.NoBeli,b.tgl,c.NamaSupp From am_beliapp a inner join am_voucherhdr b "
        carisql1 = carisql1 + "on a.Ref1 = b.novoucher inner join am_supplier c on a.Kodesupp = c.KodeSupp"
        namatabel = "Voucher"
        frmsearch.Show vbModal
    ElseIf optv_biaya = True Then
        carisql1 = "Select a.novoucher,a.tgl,a.kepada,SUM(b.jumlah)as Jumlah From am_voucherhdr a "
        carisql1 = carisql1 + "inner join am_voucherin b on a.novoucher = b.novoucher"
        carisql1 = carisql1 + " group by a.novoucher,a.tgl,a.kepada order by a.novoucher"
        namatabel = "Allvoucher"
        frmsearch.Show vbModal
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdvoucher2_GotFocus()
    If hasil1 = "" Then Exit Sub
    txtvoucher2 = hasil1
    hasil1 = ""
End Sub

Private Sub viewvbiaya()
    On Error GoTo Err_handler:
    Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_voucher_lpb)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\voucher_lpb.rpt"
        Crystal.ParameterFields(0) = "@vouc1;" + txtvoucher + ";true"
        Crystal.ParameterFields(1) = "@vouc2;" + txtvoucher2 + ";true"
        Crystal.ParameterFields(2) = "@user;" + nmuser + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
        
    Exit Sub
    SQL = "Select distinct a.novoucher,a.tgl,a.kepada,a.kdkurs,a.nilai,a.username,a.ppn,ISNULL(b.no_payment,'Unprocesed')'status',SUM(c.jumlah)'total' "
    SQL = SQL + "From am_voucherhdr a left outer join no_bank_payment b "
    SQL = SQL + "on a.novoucher = b.no_voucher inner join am_voucherin c on a.novoucher = c.novoucher "
    SQL = SQL + "Where a.novoucher >= '" + txtvoucher + "' and a.novoucher <= '" + txtvoucher2 + "'"
    SQL = SQL + " GROUP BY a.novoucher,a.tgl,a.kepada,a.kdkurs,a.nilai,a.username,a.ppn,b.no_payment"
    With rptlaporanvoucher
        .lblparam = "From Voucher : " & txtvoucher & " To Voucher : " & txtvoucher2
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
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
                        rptlaporanvoucher.Run
                        oXLS.Export rptlaporanvoucher.Pages
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
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub viewvhutang()
    On Error GoTo Err_handler:
    Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.DataFiles(0) = "Proc(am_voucher_bb)"
        Crystal.ReportFileName = AppPath & "\reports\finance\purc\voucher_bb.rpt"
        Crystal.ParameterFields(0) = "@vouc1;" + txtvoucher + ";true"
        Crystal.ParameterFields(1) = "@vouc2;" + txtvoucher2 + ";true"
        Crystal.ParameterFields(2) = "@user;" + nmuser + ";true"
        Crystal.RetrieveDataFiles
        Crystal.Action = 1
        
    Exit Sub
    SQL = "Select distinct a.Ref1,d.NamaSupp,a.nobeli,a.Ref2,c.tglbkt,isnull(c.NoBkt,'Unprocessed')'status' "
    SQL = SQL + "From am_beliapp a left outer join am_apopnfil b "
    SQL = SQL + "on a.NoBeli = b.NoBeli left outer join am_apcashlin c "
    SQL = SQL + "on a.Ref2 = c.NoApply inner join am_supplier d "
    SQL = SQL + "on a.kodesupp = d.kodesupp "
    SQL = SQL + "Where a.Ref1 >= '" + txtvoucher + "' and a.Ref1 <= '" + txtvoucher2 + "'"
    
        With rptvoucher_hutang
            .lblparam = "From Voucher : " & txtvoucher & " To Voucher : " & txtvoucher2
            .DataControl1.Source = SQL
            .DataControl1.ConnectionString = dsn
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
                        rptvoucher_hutang.Run
                        oXLS.Export rptvoucher_hutang.Pages
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
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

