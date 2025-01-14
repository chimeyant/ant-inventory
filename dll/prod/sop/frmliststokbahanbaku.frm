VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmliststokbahanbaku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Stok Bahan Baku"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   390
      Left            =   2805
      TabIndex        =   1
      Top             =   2010
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bahan Baku"
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   3675
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1545
         TabIndex        =   10
         Top             =   570
         Width           =   2070
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Kelompok Produk"
         Height          =   270
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Range Bahan Baku"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   930
         Width           =   3270
      End
      Begin VB.TextBox txtsd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2355
         TabIndex        =   6
         Top             =   1260
         Width           =   1215
      End
      Begin XtremeSuiteControls.PushButton btndari 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   1230
         Width           =   405
         _Version        =   851970
         _ExtentX        =   714
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Dari"
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
      End
      Begin VB.TextBox txtdari 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   1260
         Width           =   1245
      End
      Begin XtremeSuiteControls.PushButton btnsd 
         Height          =   330
         Left            =   1860
         TabIndex        =   5
         Top             =   1230
         Width           =   465
         _Version        =   851970
         _ExtentX        =   820
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "s.d"
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   330
         Left            =   405
         TabIndex        =   9
         Top             =   555
         Width           =   1050
         _Version        =   851970
         _ExtentX        =   1852
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Produk"
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
      End
   End
   Begin XtremeSuiteControls.PushButton cmdpreview 
      Height          =   390
      Left            =   1845
      TabIndex        =   2
      Top             =   2010
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Preview"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   -90
      Top             =   1935
      Width           =   5400
   End
End
Attribute VB_Name = "frmliststokbahanbaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String


Private Sub btndari_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    frmsearch.Show vbModal
End Sub

Private Sub btndari_GotFocus()
    If hasil = "" Then Exit Sub
    txtdari = hasil
    carisql1 = ""
    namatabel = ""
    hasil = ""
End Sub

Private Sub btnsd_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    frmsearch.Show vbModal
End Sub

Private Sub btnsd_GotFocus()
    If hasil = "" Then Exit Sub
    txtsd = hasil
    carisql1 = ""
    namatabel = ""
    hasil = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub


Private Sub cmdpreview_Click()
    On Error GoTo err_msg
    If txtdari = "" And txtsd = "" Then
        
    End If
    
    OBJ.Open dsn
    SQL = "exec am_stokbahanbaku"
    Set RS = OBJ.Execute(SQL)
    OBJ.Close
    
    With rptstokbahanbaku
        SQL = "select * from am_tempstokbahanbaku order by kdbarang"
        .DataControl1.ConnectionString = dsn
        .DataControl1.Source = SQL
        .Show vbModal
    End With
            
    Exit Sub
err_msg:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub PushButton1_Click()
    carisql1 = "select * from am_itemcode where lev=2 or lev=3"
    namatabel = "Produk"
    frmsearch.Show vbModal
End Sub
