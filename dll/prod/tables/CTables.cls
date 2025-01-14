VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTables"
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

Public Property Let UserLevel(ByVal s_userlevel As String)
    UserOnLineLevel = s_userlevel
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
            Case "addreaktor": frmaddreaktor.Show
            Case "addsop": frmaddsop.Show
            Case "addresep": frmResep.Show
            Case "addresepk": frmresep_karet.Show
            Case "konversikemasan": frmkonvlevel.Show
            Case "konversich": frmkonvlevel_edit.Show
            Case "konversilist": frmkonvlist.Show
            Case "listlem": frmlistlem.Show
            Case "konvunit": frmkonvunit.Show
            Case "definekg": frmitemkg.Show
            Case "prodplan": frmrcnprod.Show
            Case "prodpled": frmrcnprodedit.Show
            Case "prodplist": frmrcnview.Show
            Case "ktgori": frmkategori.Show
            Case "reaktor": frmreaktor.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addreaktor": hWnd = frmaddreaktor.hWnd
        Case "addsop": hWnd = frmaddsop.hWnd
        Case "addresep": hWnd = frmResep.hWnd
        Case "addresepk": hWnd = frmresep_karet.hWnd
        Case "konversikemasan": hWnd = frmkonvlevel.hWnd
        Case "konversich": hWnd = frmkonvlevel_edit.hWnd
        Case "konversilist": hWnd = frmkonvlist.hWnd
        Case "listlem": hWnd = frmlistlem.hWnd
        Case "konvunit": hWnd = frmkonvunit.hWnd
        Case "definekg": hWnd = frmitemkg.hWnd
        Case "prodplan": hWnd = frmrcnprod.hWnd
        Case "prodpled": hWnd = frmrcnprodedit.hWnd
        Case "prodplist": hWnd = frmrcnview.hWnd
        Case "ktgori": hWnd = frmkategori.hWnd
        Case "reaktor": hWnd = frmreaktor.hWnd
    End Select
End Function



