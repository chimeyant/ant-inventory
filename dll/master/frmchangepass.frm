VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmchangepass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   390
      Left            =   3750
      TabIndex        =   6
      Top             =   1260
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtrepass 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   780
      Width           =   3225
   End
   Begin VB.TextBox txtnewpass 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   450
      Width           =   3225
   End
   Begin VB.TextBox txtoldpassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   105
      Width           =   3225
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   390
      Left            =   2820
      TabIndex        =   7
      Top             =   1260
      Width           =   915
      _Version        =   851970
      _ExtentX        =   1614
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Save"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Re-Type"
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   825
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1110
   End
End
Attribute VB_Name = "frmchangepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RST As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private SQL As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    Dim valpass As String
    On Error GoTo err_msg
    SQL = "select pass from list_users where username='" & nmuser & "'"
    OpenDB
    Set RST = ConSQL.Execute(SQL)
    valpass = Cheap_Decrypt(RST!pass)
    
    If valpass <> txtoldpassword Then
        MsgBox "Old your password not valid..!", vbCritical, AppName
        CloseSQLDB
        Exit Sub
    End If
    
    If txtnewpass <> txtrepass Then
        MsgBox "Password Not Match....!", vbCritical, AppName
        CloseSQLDB
        Exit Sub
    End If
    
    SQL = "Update list_users set pass='" & Cheap_Encrypt(txtnewpass) & "' where username ='" & nmuser & "'"
    ConSQL.Execute (SQL)
    MsgBox "Change Password Completed...!", vbInformation, AppName
    CloseSQLDB
    
    Exit Sub
err_msg:
    MsgBox Err.Description, vbCritical, AppName
End Sub
