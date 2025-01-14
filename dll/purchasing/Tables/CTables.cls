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
            Case "addsupplier": frmsupplier.Show
            Case "changesupplier": frmsupplieredit.Show
            Case "daftarsupplier": frmsupplierlist.Show
            Case "pricelist": frmpricelist.Show
            Case "addunit": frmunit.Show
            Case "changeunit": frmunitedit.Show
            Case "listunit": frmlaporan2.Show
            Case "addbahanbaku": frmitem.Show
            Case "listitembb": frmlaporan5.Show
            Case "addpackaging": frmpackaging.Show
            Case "viewstokbahanbaku": frmsoh.Show
            Case "addmin": frmminadd.Show
            Case "cngmin": frmminedit.Show
            Case "listmin": frmlaporanminstok.Show
            Case "graphicprice": frmgraprice.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addsupplier": hWnd = frmsupplier.hWnd
        Case "changesupplier": hWnd = frmsupplieredit.hWnd
        Case "daftarsupplier": hWnd = frmsupplierlist.hWnd
        Case "pricelist": hWnd = frmpricelist.hWnd
        Case "addunit": hWnd = frmunit.hWnd
        Case "changeunit": hWnd = frmunitedit.hWnd
        Case "listunit": hWnd = frmlaporan2.hWnd
        Case "addbahanbaku": hWnd = frmitem.hWnd
        Case "listitembb": hWnd = frmlaporan5.hWnd
        Case "addpackaging": hWnd = frmpackaging.hWnd
        Case "viewstokbahanbaku": hWnd = frmsoh.hWnd
        Case "addmin": hWnd = frmminadd.hWnd
        Case "cngmin": hWnd = frmminedit.hWnd
        Case "listmin": hWnd = frmlaporanminstok.hWnd
        Case "graphicprice": hWnd = frmgraprice.hWnd
    End Select
End Function

