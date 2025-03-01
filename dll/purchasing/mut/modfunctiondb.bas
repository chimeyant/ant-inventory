Attribute VB_Name = "modfunctiondb"
Sub LoadDatabaseProperty()
    Dim StrKoneksi As String
    Dim ArrKoneksi() As String
    Dim i As Integer
    
    Open AppPath + "\sqlserver.dll" For Input As #1
        Line Input #1, StrKoneksi
    Close
    
    ArrKoneksi = Split(StrKoneksi, "|")
    
    For i = 1 To 4 Step 1
        If i = 1 Then
            dbServer = Cheap_Decrypt(ArrKoneksi(0))
        End If
        If i = 2 Then
            dbName = Cheap_Decrypt(ArrKoneksi(1))
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
    dsn = "Provider=SQLOLEDB.1;Password=" + dbPass + ";User ID=" + dbUser + ";Initial Catalog=" + dbName + ";Data Source=" + dbServer
End Sub

Sub OpenDB()
    OpenSQLDB dbServer, dbName, dbUser, dbPass
End Sub

Public Function dsnreport() As Variant
    dsnreport = "DSN=" & dbServer & ";UID=" & dbUser & ";PWD=" & dbPass & ";DSQ=" & dbName & ""
End Function

Function GetNoStok() As String
    On Error GoTo err_handler
    SQL = "GET_NOSTOK"
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    Set CMD = New ADODB.Command
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = SQL
    CMD.ActiveConnection = OBJ
    Set RST = CMD.Execute
    GetNoStok = RST!NOSTOK
    OBJ.Close
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    
End Function

