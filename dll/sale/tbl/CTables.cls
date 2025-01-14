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
            Case "addunit": frmunit.Show
            Case "changeunit": frmunitedit.Show
            Case "listunit": frmunitlist.Show
            Case "addcat": frmcategory.Show
            Case "changecat": frmcategoryedit.Show
            Case "catlist": frmcategorylist.Show
            Case "additem": frmitem.Show
            Case "itemlist": frmitemlist.Show
            Case "addgudang": frmgudang.Show
            Case "changegudang": frmgudangedit.Show
            Case "gudanglist": frmgudanglist.Show
            Case "addarea": frmarea.Show
            Case "changearea": frmareaedit.Show
            Case "arealist": frmarealist.Show
            Case "addcust": frmcustomer.Show
            Case "changecust": frmcustomeredit.Show
            Case "custlist": frmcustomerlist.Show
            Case "salesmanage": frmsalesmanage.Show
        End Select
    Else
        MsgBox "Illegal DLL File .....!", vbCritical, AppName
        Exit Sub
    End If
End Sub

Public Function hWnd() As Long
    Select Case v_name
        Case "addunit": hWnd = frmunit.hWnd
        Case "changeunit": hWnd = frmunitedit.hWnd
        Case "listunit": hWnd = frmunitlist.hWnd
        Case "addcat": hWnd = frmcategory.hWnd
        Case "changecat": hWnd = frmcategoryedit.hWnd
        Case "catlist": hWnd = frmcategorylist.hWnd
        Case "additem": hWnd = frmitem.hWnd
        Case "itemlist": hWnd = frmitemlist.hWnd
        Case "addgudang": hWnd = frmgudang.hWnd
        Case "changegudang": hWnd = frmgudangedit.hWnd
        Case "gudanglist": hWnd = frmgudanglist.hWnd
        Case "addarea": hWnd = frmarea.hWnd
        Case "changearea": hWnd = frmareaedit.hWnd
        Case "arealist": hWnd = frmarealist.hWnd
        Case "addcust": hWnd = frmcustomer.hWnd
        Case "changecust": hWnd = frmcustomeredit.hWnd
        Case "custlist": hWnd = frmcustomerlist.hWnd
        Case "salesmanage": hWnd = frmsalesmanage.hWnd
    End Select
End Function
