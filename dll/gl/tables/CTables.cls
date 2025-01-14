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
            Case "addtypecomp": frmtypecomp.Show
            Case "changetypecomp": frmtypecompedit.Show
            Case "typecomplist": frmtypecomplist.Show
            Case "addmasteracc": frmaccount.Show
            Case "changemasteracc": frmaccountedit.Show
            Case "listmasteracc": frmaccountlist.Show
            Case "addcompany": frmcompany.Show
            Case "changecompany": frmcompanyedit.Show
            Case "listcompany": frmcompanylist.Show
            Case "browseacc": frmbrowseacc.Show
            Case "listaccount": frmlistaccount.Show
            Case "listbudget": frmlistbudget.Show
            Case "addkurs": frmkurs.Show
            Case "changekurs": frmkursedit.Show
            Case "listkurs": frmkurslist.Show
            Case "addjenisfa": frmjenisfa.Show
            Case "changejenisfa": frmjenisfaedit.Show
            Case "listjenisfa": frmjenisfalist.Show
            Case "addbank": frmbank.Show
            Case "changebank": frmbankedit.Show
            Case "listbank": frmbanklist.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addtypecomp": hWnd = frmtypecomp.hWnd
        Case "changetypecomp": hWnd = frmtypecompedit.hWnd
        Case "typecomplist": hWnd = frmtypecomplist.hWnd
        Case "addmasteracc": hWnd = frmaccount.hWnd
        Case "changemasteracc": hWnd = frmaccountedit.hWnd
        Case "listmasteracc": hWnd = frmaccountlist.hWnd
        Case "addcompany": hWnd = frmcompany.hWnd
        Case "changecompany": hWnd = frmcompanyedit.hWnd
        Case "listcompany": hWnd = frmcompanylist.hWnd
        Case "browseacc": hWnd = frmbrowseacc.hWnd
        Case "listaccount": hWnd = frmlistaccount.hWnd
        Case "listbudget": hWnd = frmlistbudget.hWnd
        Case "addkurs": hWnd = frmkurs.hWnd
        Case "changekurs": hWnd = frmkursedit.hWnd
        Case "listkurs": hWnd = frmkurslist.hWnd
        Case "addjenisfa": hWnd = frmjenisfa.hWnd
        Case "changejenisfa": hWnd = frmjenisfaedit.hWnd
        Case "listjenisfa": hWnd = frmjenisfalist.hWnd
        Case "addbank": hWnd = frmbank.hWnd
        Case "changebank": hWnd = frmbankedit.hWnd
        Case "listbank": hWnd = frmbanklist.hWnd
    End Select
End Function


