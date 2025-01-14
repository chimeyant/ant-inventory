VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private v_token As String

Public Property Let Token(ByVal s_token As String)
    v_token = s_token
End Property

Public Property Let Path(ByVal s_path As String)
    AppPath = s_path
End Property

Public Sub Show()
    If v_token = "kusumah" Then
        LoadDatabaseProperty
        frmModul.Show
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    hWnd = frmModul.hWnd
End Function

