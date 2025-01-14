VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSale"
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
            Case "addkoreksi": frmcorar.Show
            Case "changekoreksi": frmcoraredit.Show
            Case "writeoff": frmwriteoff.Show
            Case "addpembayaran": frmpayar.Show
            Case "changepembayaran": frmpayaredit.Show
            Case "daftarkoreksi": frmdaftarkoreksi.Show
            Case "daftarpenagihan": frmdaftartagih.Show
            Case "daftarpiutang": frmdaftarpiutang.Show
            Case "daftarpiutangkartu": frmdaftarpiutang_kartu.Show
            Case "daftarpiutangsisa": frmdaftarpiutang_sisa.Show
            Case "daftarmutasipiutang": frmmutasipiutang.Show
            Case "daftarpembayaran": frmdaftarbayarbydate.Show
            Case "daftarpembayarandetail": frmdaftarbayarbydatedetail.Show
            Case "daftarttpembayaran": frmdaftarTT.Show
            Case "maintenancegiro": frmcekgiro.Show
            Case "addgirotolak": frmcekar.Show
            Case "changegirotolak": frmcekaredit.Show
            Case "daftargiro": frmdaftargiro.Show
            Case "daftargiroturun": frmdaftargiroturun.Show
            Case "defineaccbank": frmdefineaccbank.Show
            Case "definejurnal": frmdefinejurnal.Show
            Case "defineaging": frmaging.Show
            Case "definekgbase": frmitemkg.Show
            Case "defineaccountcust": frmdefineacccustomer.Show
            Case "lap_posting_penjualan": frmlaporanpostingpenjualan.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addkoreksi": hWnd = frmcorar.hWnd
        Case "changekoreksi": hWnd = frmcoraredit.hWnd
        Case "writeoff": hWnd = frmwriteoff.hWnd
        Case "addpembayaran": hWnd = frmpayar.hWnd
        Case "changepembayaran": hWnd = frmpayaredit.hWnd
        Case "daftarkoreksi": hWnd = frmdaftarkoreksi.hWnd
        Case "daftarpenagihan": hWnd = frmdaftartagih.hWnd
        Case "daftarpiutang": hWnd = frmdaftarpiutang.hWnd
        Case "daftarpiutangkartu": hWnd = frmdaftarpiutang_kartu.hWnd
        Case "daftarpiutangsisa": hWnd = frmdaftarpiutang_sisa.hWnd
        Case "daftarmutasipiutang": hWnd = frmmutasipiutang.hWnd
        Case "daftarpembayaran": hWnd = frmdaftarbayarbydate.hWnd
        Case "daftarpemabyarandetail": hWnd = frmdaftarbayarbydatedetail.hWnd
        Case "daftarttpembayaran": hWnd = frmdaftarTT.hWnd
        Case "maintenancegiro": hWnd = frmcekgiro.hWnd
        Case "addgirotolak": hWnd = frmcekar.hWnd
        Case "changegirotolak": hWnd = frmcekaredit.hWnd
        Case "daftargiro": hWnd = frmdaftargiro.hWnd
        Case "daftargiroturun": hWnd = frmdaftargiroturun.hWnd
        Case "defineaccbank": hWnd = frmdefineaccbank.hWnd
        Case "definejurnal": hWnd = frmdefinejurnal.hWnd
        Case "defineaging": hWnd = frmaging.hWnd
        Case "definekgbase": hWnd = frmitemkg.hWnd
        Case "defineaccountcust": hWnd = frmdefineacccustomer.hWnd
        Case "lap_posting_penjualan": hWnd = frmlaporanpostingpenjualan.hWnd
    End Select
End Function

