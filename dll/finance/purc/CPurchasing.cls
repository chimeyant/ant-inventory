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
            Case "confirmpenerimaan": frmpenerimaan_app.Show
            Case "reprintpenerimaan": frmreprint.Show
            Case "confirmpenerimaanretur": frmpenerimaan_appretur.Show
            Case "unconfirm": frmunconfirm.Show
            Case "unconfirmretur": frmunconfirmretur.Show
            Case "createvoucher": frmvoucher.Show
            Case "addpembayaran": frmpayap.Show
            Case "changepembayaran": frmpayapedit.Show
            Case "unposthutang": frmunposthutang.Show
            Case "addkoreksi": frmcorar.Show
            Case "changekoreksi": frmcoraredit.Show
            Case "daftarkoreksi": frmcorarlist.Show
            Case "daftarbayarhutang": frmlaporan3.Show
            Case "daftarhutangkartu": frmlaporanhutangkartu.Show
            Case "daftarsisahutang": frmlaporanhutangsisa.Show
            Case "lapbyjt": frmlaporanhutangjt.Show
            Case "maintenancegiro": frmcekgiro.Show
            Case "addgirotolak": frmcekap.Show
            Case "changegirotolak": frmcekapedit.Show
            Case "daftargiro": frmgirolist.Show
            Case "daftarconfirmunconfirm": frmlaporan1.Show
            Case "laporanvoucer": frmlaporanvoucer.Show
            Case "laporanprocesvoucer": frmlapvoucher.Show
            Case "definesupp": frmdefineaccsupp.Show
            Case "defineaccbank": frmdefineaccbank.Show
            Case "definejurnal": frmdefinejurnal.Show
            Case "printjournalhutang": frmprint_journal_hutang.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "confirmpenerimaan": hWnd = frmpenerimaan_app.hWnd
        Case "reprintpenerimaan": hWnd = frmreprint.hWnd
        Case "confirmpenerimaanretur": hWnd = frmpenerimaan_appretur.hWnd
        Case "unconfirm": hWnd = frmunconfirm.hWnd
        Case "unconfirmretur": hWnd = frmunconfirmretur.hWnd
        Case "createvoucher": hWnd = frmvoucher.hWnd
        Case "addpembayaran": hWnd = frmpayap.hWnd
        Case "changepembayaran": hWnd = frmpayapedit.hWnd
        Case "unposthutang": hWnd = frmunposthutang.hWnd
        Case "addkoreksi": hWnd = frmcorar.hWnd
        Case "changekoreksi": hWnd = frmcoraredit.hWnd
        Case "daftarkoreksi": hWnd = frmcorarlist.hWnd
        Case "daftarbayarhutang": hWnd = frmlaporan3.hWnd
        Case "daftarhutangkartu": hWnd = frmlaporanhutangkartu.hWnd
        Case "daftarsisahutang": hWnd = frmlaporanhutangsisa.hWnd
        Case "lapbyjt": hWnd = frmlaporanhutangjt.hWnd
        Case "maintenancegiro": hWnd = frmcekgiro.hWnd
        Case "addgirotolak": hWnd = frmcekap.hWnd
        Case "changegirotolak": hWnd = frmcekapedit.hWnd
        Case "daftargiro": hWnd = frmgirolist.hWnd
        Case "daftarconfirmunconfirm": hWnd = frmlaporan1.hWnd
        Case "laporanvoucer": hWnd = frmlaporanvoucer.hWnd
        Case "laporanprocesvoucer": hWnd = frmlapvoucher.hWnd
        Case "definesupp": hWnd = frmdefineaccsupp.hWnd
        Case "defineaccbank": hWnd = frmdefineaccbank.hWnd
        Case "definejurnal": hWnd = frmdefinejurnal.hWnd
        Case "printjournalhutang": hWnd = frmprint_journal_hutang.hWnd
    End Select
End Function

