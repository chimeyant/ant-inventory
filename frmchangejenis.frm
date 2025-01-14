VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmchangejenis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubah Jenis"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtkodejenis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   1
      Top             =   75
      Width           =   870
   End
   Begin VB.TextBox txtjenis 
      Appearance      =   0  'Flat
      Height          =   570
      Left            =   1530
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   3675
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   360
      Left            =   4290
      TabIndex        =   2
      Top             =   1110
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Close"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   360
      Left            =   3345
      TabIndex        =   3
      Top             =   1110
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Save"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   1110
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Clear"
      Appearance      =   6
   End
   Begin VB.Label Label1 
      Caption         =   "KODE JENIS"
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "JENIS"
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   435
      Width           =   1275
   End
End
Attribute VB_Name = "frmchangejenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private obj As New ADODB.Connection
Private rs As ADODB.Recordset
Private sql As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub txtkodejenis_Change()

End Sub

Private Sub txtkodejenis_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        sql = "select * from list_jenis where kd_jenis='" & txtkodejenis & "'"
        obj.Open dsn
        Set rs = obj.Execute(sql)
        If rs.EOF Then
            MsgBox "Data tidak ditemukan", vbCritical, AppName
            obj.Close
        End If
        txtjenis = rs!jenis
        obj.Close
    End If
End Sub
