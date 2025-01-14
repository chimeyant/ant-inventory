VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmaddjenis 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAMBAH JENIS BARANG"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtjenis 
      Appearance      =   0  'Flat
      Height          =   570
      Left            =   1365
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1020
      Width           =   3840
   End
   Begin VB.TextBox txtkodejenis 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1365
      MaxLength       =   3
      TabIndex        =   0
      Top             =   630
      Width           =   870
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   480
      Left            =   4260
      TabIndex        =   4
      Top             =   1785
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Close"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   465
      Left            =   3315
      TabIndex        =   5
      Top             =   1800
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Save"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   450
      Left            =   2370
      TabIndex        =   6
      Top             =   1815
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Clear"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " JENIS BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   30
      TabIndex        =   7
      Top             =   120
      Width           =   2985
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   -600
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   2805
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   2220
      Shape           =   4  'Rounded Rectangle
      Top             =   1695
      Width           =   3600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE JENIS"
      Height          =   225
      Left            =   45
      TabIndex        =   2
      Top             =   705
      Width           =   1275
   End
End
Attribute VB_Name = "frmaddjenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private obj As New ADODB.Connection
Private rs As ADODB.Recordset

Private Sub cmdclear_Click()
    txtkodejenis = ""
    txtjenis = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err_handler:
    sql = "select * from list_jenis where 0=1"
    obj.Open dsn
    Set rs = New ADODB.Recordset
    rs.Open sql, obj, adOpenDynamic, adLockOptimistic
    With rs
        .AddNew
        !kd_jenis = txtkodejenis
        !jenis = txtjenis
        .Update
    End With
    obj.Close
    MsgBox suksessimpan, vbInformation, AppName
    cmdclear_Click
    Exit Sub
err_handler:
    If obj.State = 1 Then obj.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

