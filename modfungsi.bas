Attribute VB_Name = "modfungsi"
Private OBJ As ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Sub LoadDatabaseProperty()
    Dim StrKoneksi As String
    Dim ArrKoneksi() As String
    Dim i As Integer
    
    Open App.Path + "\sqlserver.dll" For Input As #1
        Line Input #1, StrKoneksi
    Close
    
    ArrKoneksi = Split(StrKoneksi, "|")
    
    For i = 1 To 4 Step 1
        If i = 1 Then
            dbServer = Cheap_Decrypt(ArrKoneksi(0))
        End If
        If i = 2 Then
            DBName = Cheap_Decrypt(ArrKoneksi(1))
        End If
        If i = 3 Then
            dbUser = Cheap_Decrypt(ArrKoneksi(2))
        End If
        If i = 4 Then
            dbPass = Cheap_Decrypt(ArrKoneksi(3))
        End If
    Next i
    
    If remoteserver = True Then
        dbServer = dbRemoteServer
    End If
    'initial dsn
    dsn = "Provider=SQLOLEDB.1;Password=" + dbPass + ";User ID=" + dbUser + ";Initial Catalog=" + DBName + ";Data Source=" + dbServer
End Sub


Function getUnlock() As Boolean
    'On Error GoTo err_handler:
    Dim status As Boolean
    Dim strCon As String
    
   strCon = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & "antsoftmedia.com" & ";DATABASE=" & "antsoft_lock" & ";" & _
         "UID=" & "kusumah" & ";PWD=" & "s3d3rh4n4" & ";OPTION=3"
    
    Set OBJ = New ADODB.Connection
    OBJ.Open strCon
    
    SQL = "select * from data_sparta "
    
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenForwardOnly, adLockOptimistic
    If RST.EOF Then
        OBJ.Close
        MsgBox "Server tidak ditemukan...!, Silahkan ulang kembali", vbCritical, AppName
        End
    End If
    
   
    If RST!status <> "BLOCK" Then
        getUnlock = True
    Else
        getUnlock = False
    End If
    OBJ.Close
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    getUnlock = False
End Function

Public Function dsnreport() As Variant
    dsnreport = "DSN=" & dbServer & ";UID=" & dbUser & ";PWD=" & dbPass & ";DSQ=" & DBName & ""
End Function

