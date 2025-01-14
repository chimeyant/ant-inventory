VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInvoicing"
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
        subme
        Select Case v_name
            Case "AddSO": frmsorder.Show
            Case "ChangeSO": frmsorderedit.Show
            Case "socancel": frmsocancel.Show
            Case "AddSJ": frmsjalan.Show
            Case "ChangeSJ": frmsjalanedit.Show
            Case "PrintSJ": frmsjalanprint.Show
            Case "AddLot": frmaddlotsj.Show
            Case "addfjual": frminvoice.Show
            Case "changefjual": frminvoicedit.Show
            Case "printfjual": frminvoiceprint.Show
            Case "prevfjual": frminquery.Show
            Case "definefpajak": frmseri.Show
            Case "browsefpajak": frmseribrowse.Show
            Case "addsjsby": frmsjalansby.Show
            Case "listsjsby": frmsjalansbylist.Show
            Case "iqso": frminqueryso.Show
            Case "solist": frmsorderlist.Show
            Case "lapsjdaftar": frmdaftarsurat.Show
            Case "lapsjdaftargd": frmdaftarsuratgudang.Show
            Case "lapsjbyfaktur": frmsjbyfaktur.Show
            Case "lapsjbylot": frmsjbylot.Show
            Case "lapsjlap": frmdaftarsuratjalan.Show
            Case "dafjual": frmdaftarjual.Show
            Case "dafjualdet": frmdaftarjualbayar.Show
            Case "jualmonthly": frmdaftarjual2.Show
            Case "bykategori": frmdaftarjualkategori.Show
            Case "dafkomisi": frmdaftarkomisi.Show
            Case "anjual": frmanalisajual.Show
            Case "lapvouc": frmlapvoucher.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "AddSO": hWnd = frmsorder.hWnd
        Case "ChangeSO": hWnd = frmsorderedit.hWnd
        Case "socancel": hWnd = frmsocancel.hWnd
        Case "AddSJ": hWnd = frmsjalan.hWnd
        Case "ChangeSJ": hWnd = frmsjalanedit.hWnd
        Case "PrintSJ": hWnd = frmsjalanprint.hWnd
        Case "AddLot": hWnd = frmaddlotsj.hWnd
        Case "addfjual": hWnd = frminvoice.hWnd
        Case "changefjual": hWnd = frminvoicedit.hWnd
        Case "printfjual": hWnd = frminvoiceprint.hWnd
        Case "prevfjual": hWnd = frminquery.hWnd
        Case "definefpajak": hWnd = frmseri.hWnd
        Case "browsefpajak": hWnd = frmseribrowse.hWnd
        Case "addsjsby": hWnd = frmsjalansby.hWnd
        Case "listsjsby": hWnd = frmsjalansbylist.hWnd
        Case "iqso": hWnd = frminqueryso.hWnd
        Case "solist": hWnd = frmsorderlist.hWnd
        Case "lapsjdaftar": hWnd = frmdaftarsurat.hWnd
        Case "lapsjdaftargd": hWnd = frmdaftarsuratgudang.hWnd
        Case "lapsjbyfaktur": hWnd = frmsjbyfaktur.hWnd
        Case "lapsjbylot": hWnd = frmsjbyfaktur.hWnd
        Case "lapsjlap": hWnd = frmdaftarsuratjalan.hWnd
        Case "dafjual": hWnd = frmdaftarjual.hWnd
        Case "dafjualdet": hWnd = frmdaftarjualbayar.hWnd
        Case "jualmonthly": hWnd = frmdaftarjual2.hWnd
        Case "bykatgori": hWnd = frmdaftarjualkategori.hWnd
        Case "dafkomisi": hWnd = frmdaftarkomisi.hWnd
        Case "anjual": hWnd = frmanalisajual.hWnd
        Case "lapvouc": hWnd = frmlapvoucher.hWnd
    End Select
End Function
