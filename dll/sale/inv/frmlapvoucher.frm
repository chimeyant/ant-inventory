VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmlapvoucher 
   Caption         =   "Laporan Voucher"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnclose 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "View"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker dtptgl2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
      _Version        =   851970
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtptgl1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1935
      _Version        =   851970
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   4575
   End
   Begin MSForms.OptionButton optv_hutang 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3201;873"
      Value           =   "0"
      Caption         =   "voucher hutang"
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.OptionButton optv_biaya 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2990;873"
      Value           =   "0"
      Caption         =   "voucher biaya"
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmlapvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub btnview_Click()
    If optv_biaya = True Then
        viewvbiaya
    Else
        viewvhutang
    End If
End Sub

Private Sub Form_Load()
    dtptgl1 = Date
    dtptgl2 = Date
End Sub

Private Sub optv_biaya_Click()
    dtptgl1.Move 240, 840
    dtptgl2.Move 240, 1560
End Sub

Private Sub optv_hutang_Click()
    dtptgl1.Move 2760, 840
    dtptgl2.Move 2760, 1560
End Sub

Private Sub viewvbiaya()
    Dim tgl1 As String
    Dim tgl2 As String

    tgl1 = Format(dtptgl1, "yyyy/MM/dd")
    tgl2 = Format(dtptgl2, "yyyy/MM/dd")
    SQL = "Select distinct a.novoucher,a.tgl,a.kepada,c.notrx "
    SQL = SQL + "from am_voucherhdr a left outer join no_bank_payment b "
    SQL = SQL + "on a.novoucher = b.no_voucher "
    SQL = SQL + "right outer join gl_transaksi c on c.notrx = b.notrx "
    SQL = SQL + "Where a.tgl >= '" + tgl1 + "' and a.tgl <= '" + tgl2 + "' "
    SQL = SQL + "Order By a.novoucher asc"
    
    With rptlaporanvoucher
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .Show
    End With
    
End Sub

Private Sub viewvhutang()
'
End Sub

Function tanggal1()
    tanggal1 = Month(dtptgl1) & "/" & Day(dtptgl1) & "/" & Year(dtptgl1)
End Function

Function tanggal2()
    tanggal2 = Month(dtptgl2) & "/" & Day(dtptgl2) & "/" & Year(dtptgl2)
End Function
