VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private v_token As String
Private v_name As String


Public Property Let Token(ByVal s_token As String)
    v_token = s_token
End Property

Public Property Let Path(ByVal s_path As String)
    AppPath = s_path
End Property

Public Property Let FormName(ByVal s_name As String)
   v_name = s_name
End Property

Public Property Let fastsearch(ByVal s_fastsearching As Boolean)
    v_fastsearching = s_fastsearching
End Property

Public Property Let fstgl1(ByVal s_fstgl1 As Date)
    v_fstgl1 = s_fstgl1
End Property

Public Property Let fstgl2(ByVal s_fstgl2 As Date)
    v_fstgl2 = s_fstgl2
End Property

Public Property Let UserOnline(ByVal s_useronline As String)
    nmuser = s_useronline
End Property

Public Property Let AsRemote(ByVal s_remote As Boolean)
    remoteserver = s_remote
End Property

Public Property Let IpServer(ByVal s_IpServer As String)
    dbRemoteServer = s_IpServer
End Property

Public Sub Show()
    If v_token = "kusumah" Then
        LoadDatabaseProperty
        Select Case v_name
            Case "rpttrialbalance": frmbalancetrial.Show
            Case "rptbukubesar": frmbukubesar.Show
            Case "rptworksheet": frmworksheet.Show
            Case "rptbukukas": frmbukukas.Show
            Case "rptbalance": frmbalance.Show
            Case "rptincome": frmincome.Show
            Case "rptlistaktiva": frmaktivadaftar.Show
            Case "rptnilaiaktiva": frmaktivanilai.Show
            Case "rptaktivajual": frmaktivajual.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
       Case "rpttrialbalance": hWnd = frmbalancetrial.hWnd
       Case "rptbukubesar": hWnd = frmbukubesar.hWnd
       Case "rptworksheet": hWnd = frmworksheet.hWnd
       Case "rptbukukas": hWnd = frmbukukas.hWnd
       Case "rptbalance": hWnd = frmbalance.hWnd
       Case "rptincome": hWnd = frmincome.hWnd
       Case "rptlistaktiva": hWnd = frmaktivadaftar.hWnd
       Case "rptnilaiaktiva": hWnd = frmaktivanilai.hWnd
       Case "rptaktivajual": hWnd = frmaktivajual.hWnd
    End Select
End Function





