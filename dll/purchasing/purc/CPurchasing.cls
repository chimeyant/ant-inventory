VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPurchasing"
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
            Case "addpermintaan": frmpermintaan.Show
            Case "cngpermintaan": frmpermintaanedit.Show
            Case "clspermintaan": frmpermintaanclose.Show
            Case "listpermintaan": frmpermintaanlist.Show
            Case "daylistpermintaan": frmpermintcount.Show
            Case "addpo": frmpurchaseorder.Show
            Case "changepo": frmpurchaseorderedit.Show
            Case "closepo": frmpurchaseorderclose.Show
            Case "addpenerimaan": frmpenerimaan.Show
            Case "changepenerimaan": frmpenerimaanedit.Show
            Case "returpenerimaan": frmpenerimaanretur.Show
            Case "printbpb": frmprintbpb.Show
            Case "addpemakaian": frmpemakaian.Show
            Case "changepemakaian": frmpemakaianedit.Show
            Case "sisapemakaian": frmpemakaiansisa.Show
            Case "inquerypo": frminquery1.Show
            Case "daftarpo": frmpurchaseorderlist.Show
            Case "lappembelian": frmlaporanpembelian.Show
            Case "lappo": frmlaporanpo.Show
            Case "lappemakaian": frmlaporanpemakaian.Show
            Case "printpermintaan": frmReprint.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addpermintaan": hWnd = frmpermintaan.hWnd
        Case "cngpermintaan": hWnd = frmpermintaanedit.hWnd
        Case "clspermintaan": hWnd = frmpermintaanclose.hWnd
        Case "listpermintaan": hWnd = frmpermintaanlist.hWnd
        Case "daylistpermintaan": hWnd = frmpermintcount.hWnd
        Case "addpo": hWnd = frmpurchaseorder.hWnd
        Case "changepo": hWnd = frmpurchaseorderedit.hWnd
        Case "closepo": hWnd = frmpurchaseorderclose.hWnd
        Case "addpenerimaan": hWnd = frmpenerimaan.hWnd
        Case "changepenerimaan": hWnd = frmpenerimaanedit.hWnd
        Case "returpenerimaan": hWnd = frmpenerimaanretur.hWnd
        Case "printbpb": hWnd = frmprintbpb.hWnd
        Case "addpemakaian": hWnd = frmpemakaian.hWnd
        Case "changepemakaian": hWnd = frmpemakaianedit.hWnd
        Case "sisapemakaian": hWnd = frmpemakaiansisa.hWnd
        Case "inquerypo": hWnd = frminquery1.hWnd
        Case "daftarpo": hWnd = frmpurchaseorderlist.hWnd
        Case "lappembelian": hWnd = frmlaporanpembelian.hWnd
        Case "lappo": hWnd = frmlaporanpo.hWnd
        Case "lappemakaian": hWnd = frmlaporanpemakaian.hWnd
        Case "printpermintaan": hWnd = frmReprint.hWnd
    End Select
End Function


