VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmotoritas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Formula"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnmprod 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1350
      TabIndex        =   2
      Top             =   495
      Width           =   3420
   End
   Begin VB.TextBox txtKdProduk 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1350
      TabIndex        =   1
      Top             =   135
      Width           =   2010
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
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
   Begin XtremeSuiteControls.PushButton btnKodeProduk 
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   1140
      _Version        =   851970
      _ExtentX        =   2011
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Kode Produk : "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   0
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   465
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   1170
      _Version        =   851970
      _ExtentX        =   2064
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Open Formula"
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
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Produk  :"
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
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   1200
   End
End
Attribute VB_Name = "frmotoritas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private SQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnKodeProduk_Click()
    carisql1 = "select * from am_itemcode where (lev =3 or lev =4)"
    namatabel = "Produk"
    frmsearch.Show vbModal
End Sub

Private Sub btnKodeProduk_GotFocus()
    If hasil = "" Then Exit Sub
    txtKdProduk = hasil1
    txtnmprod = hasil2
    carisql1 = ""
    namatabel = ""
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsave_Click()
    If txtKdProduk = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from list_produk_masterkey where 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RS
        .AddNew
        !KODE_PRODUK = txtKdProduk
        !tgl = Date
        !UserName = nmuser
        !otoritas = "1"
        !keterangan = ""
        .Update
    End With
    MsgBox "Formula SOP " & txtnmprod & " has been opened", vbInformation, AppName
    txtKdProduk = ""
    txtnmprod = ""
    OBJ.Close
End Sub

