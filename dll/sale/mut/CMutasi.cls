VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private v_token As String
Private v_name As String
Private v_dsn As String


Public Property Let token(ByVal s_token As String)
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

Public Property Let dsnServer(ByVal s_dsn As String)
    v_dsn = s_dsn
End Property

Public Sub Show()
    If v_token = "kusumah" Then
        LoadDatabaseProperty
        Select Case v_name
            Case "addmut": frmterima.Show
            Case "changemut": frmterimaedit.Show
            Case "overzak": frmoverzak.Show
            Case "mutbylot": frmmutstokbylot.Show
            Case "mutpabrik": frmmutpabrik.Show
            Case "adjstok": frmAdjStok.Show
            Case "mutlist": frmterimalist.Show
            Case "mutwip": frmstokbylot.Show  'frmmutwip.Show
            Case "kartu": frmstokkartu.Show
            Case "mutgudang": frmmutgdg.Show
            Case "mutfail": frmmutfail.Show
            Case "pindahgudang": frmmutasigdg.Show 'frmpindah.Show
            Case "pindahgudangch": frmpindahedit.Show
            Case "pindahgudangprt": frmpindah_reprint.Show
            Case "pindahgudangpalet": frmpalet.Show
            Case "printmut": frmterimaprint.Show
            Case "printstok": frmdaftarposisi.Show
            Case "scanpalet": frmterimawip.Show 'frmterimawiprepair.Show  '
            Case "comparepalet": frmvalidasipalet.Show
            Case "packadd": frmpackadd.Show
            Case "packlist": frmpacklist.Show
            Case "packclose": frmpackclose.Show
            Case "towip": frmpindahwip.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addmut": hWnd = frmterima.hWnd
        Case "changemut": hWnd = frmterimaedit.hWnd
        Case "overzak": hWnd = frmoverzak.hWnd
        Case "mutbylot": hWnd = frmmutstokbylot.hWnd
        Case "mutpabrik": hWnd = frmmutpabrik.hWnd
        Case "adjstok": hWnd = frmAdjStok.hWnd
        Case "mutlist": hWnd = frmterimalist.hWnd
        Case "mutwip": hWnd = frmstokbylot.hWnd  'frmmutwip.hWnd
        Case "kartu": hWnd = frmstokkartu.hWnd
        Case "mutgudang": hWnd = frmmutgdg.hWnd
        Case "mutfail": hWnd = frmmutfail.hWnd
        Case "pindahgudang": hWnd = frmmutasigdg.hWnd 'frmpindah.hWnd
        Case "pindahgudangch": hWnd = frmpindahedit.hWnd
        Case "pindahgudangprt": hWnd = frmpindah_reprint.hWnd
        Case "pindahgudangpalet": hWnd = frmpalet.hWnd
        Case "printmut": hWnd = frmterimaprint.hWnd
        Case "printstok": hWnd = frmdaftarposisi.hWnd
        Case "scanpalet": hWnd = frmterimawip.hWnd 'frmterimawiprepair.hWnd  '
        Case "comparepalet": hWnd = frmvalidasipalet.hWnd
        Case "packadd": hWnd = frmpackadd.hWnd
        Case "packlist": hWnd = frmpacklist.hWnd
        Case "packclose": hWnd = frmpackclose.hWnd
        Case "towip": hWnd = frmpindahwip.hWnd
    End Select
End Function

