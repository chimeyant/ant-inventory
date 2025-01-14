VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLedger"
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
            Case "addjurnal": frmtran.Show
            Case "changejurnal": frmtranedit.Show
            Case "tranlist": frmtranlist.Show
            Case "tranpostingjurnal": frmtranposting.Show
            Case "untranpostingjurnal": frmtranunposting.Show
            Case "cashbankin": frmintran.Show
            Case "changecashbankin": frmintranedit.Show
            Case "listcashbankin": frmintranlist.Show
            Case "postingcashbankin": frmintranposting.Show
            Case "unpostingcashbankin": frmintranunposting.Show
            Case "addcashbankout": frmoutran.Show
            Case "addcashbankout2": frmoutran2.Show
            Case "lapcbo": frmLaporanCBO.Show
            Case "changecashbankout": frmoutranedit.Show
            Case "listcashbankout": frmoutranlist.Show
            Case "postingcashbankout": frmoutranposting.Show
            Case "unpostingcashbankout": frmoutranunposting.Show
            Case "buktikeluar": frmpengeluaran.Show
            Case "buktinew": frmoutrankas.Show
            Case "editbuktinew": frmoutrankasedit.Show
            Case "listbuktikeluar": frmrekapKK.Show
            Case "reprintbuktinew": frmreprint.Show
            Case "trans_etol": frme_tol.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addjurnal": hWnd = frmtran.hWnd
        Case "changejurnal": hWnd = frmtranedit.hWnd
        Case "tranlist": hWnd = frmtranlist.hWnd
        Case "tranpostingjurnal": hWnd = frmtranposting.hWnd
        Case "untranpostingjurnal": hWnd = frmtranunposting.hWnd
        Case "cashbankin": hWnd = frmintran.hWnd
        Case "changecashbankin": hWnd = frmintranedit.hWnd
        Case "listcashbankin": hWnd = frmintranlist.hWnd
        Case "postingcashbankin": hWnd = frmintranposting.hWnd
        Case "unpostingcashbankin": hWnd = frmintranunposting.hWnd
        Case "addcashbankout": hWnd = frmoutran.hWnd
        Case "addcashbankout2": hWnd = frmoutran2.hWnd
        Case "lapcbo": hWnd = frmLaporanCBO.hWnd
        Case "changecashbankout": hWnd = frmoutranedit.hWnd
        Case "listcashbankout": hWnd = frmoutranlist.hWnd
        Case "postingcashbankout": hWnd = frmoutranposting.hWnd
        Case "unpostingcashbankout": hWnd = frmoutranunposting.hWnd
        Case "buktikeluar": hWnd = frmpengeluaran.hWnd
        Case "buktinew": hWnd = frmoutrankas.hWnd
        Case "editbuktinew": hWnd = frmoutrankasedit.hWnd
        Case "listbuktikeluar": hWnd = frmrekapKK.hWnd
        Case "reprintbuktinew": hWnd = frmreprint.hWnd
        Case "trans_etol": hWnd = frme_tol.hWnd
    End Select
End Function




