VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSOP"
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

Public Property Let setUserOnlineLevel(ByVal s_user_level As String)
    UserOnLineLevel = s_user_level
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
            Case "addsop": frmaddsop.Show
            Case "editsop": frmeditsop.Show
            Case "changesop": frmreprintsop.Show
            Case "permintaan": frmpermint_package.Show
            Case "listconfirm": frmpermint_list.Show
            Case "reportpack": frmpermint_report.Show
            Case "konversikemasan": frmkonversi.Show
            Case "liststokbahanbaku": frmlapsop.Show
            Case "listlotpalet": frmlaplot.Show
            Case "monthlyreport": frmmonthlyprod.Show
            Case "topkilo": frmtopkg.Show
            Case "group_report": frmgroup_report.Show
            Case "scanlot": frmscanlot.Show
            Case "ubahlot": frmeditpalet.Show
            Case "kuncipalet": frmunblock_palet.Show
            Case "rfid": frmvalpalet.Show 'frmaddhpppack.Show  'frmRfID.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addsop": hWnd = frmaddsop.hWnd
        Case "editsop": hWnd = frmeditsop.hWnd
        Case "changesop": hWnd = frmreprintsop.hWnd
        Case "permintaan": hWnd = frmpermint_package.hWnd
        Case "listconfirm": hWnd = frmpermint_list.hWnd
        Case "reportpack": hWnd = frmpermint_report.hWnd
        Case "konversikemasan": hWnd = frmkonversi.hWnd
        Case "liststokbahanbaku": hWnd = frmlapsop.hWnd
        Case "listlotpalet": hWnd = frmlaplot.hWnd
        Case "monthlyreport": hWnd = frmmonthlyprod.hWnd
        Case "topkilo": hWnd = frmtopkg.hWnd
        Case "group_report": hWnd = frmgroup_report.hWnd
        Case "scanlot": hWnd = frmscanlot.hWnd
        Case "ubahlot": hWnd = frmeditpalet.hWnd
        Case "kuncipalet": hWnd = frmunblock_palet.hWnd
        Case "rfid": hWnd = frmvalpalet.hWnd 'frmaddhpppack.hWnd  'frmRfID.hWnd
    End Select
End Function




