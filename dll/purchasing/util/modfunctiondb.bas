Attribute VB_Name = "modfunctiondb"
Private OBJ As New ADODB.Connection
Private RS As ADODB.Recordset
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

Function AmbilKodeStokBaru(ByVal strtgl As String) As String
   On Error GoTo err_msg
    Dim str99 As String
    Dim SQL As String
    
    OBJ.Open dsn
    SQL = "select top 1 kdstok from am_stokbahan where kdstok like 'ST-" & strtgl & "%' order by kdstok desc"
    
    Set RS = OBJ.Execute(SQL)
    If Not RS.EOF Then
        str99 = Right(RS!kdstok, 3)
    Else
        str99 = 0
    End If
    
    str99 = str99 + 1
    
    If Len(str99) = 1 Then AmbilKodeStokBaru = "ST-" + strtgl + "." + "00" & str99
    If Len(str99) = 2 Then AmbilKodeStokBaru = "ST-" + strtgl + "." & "0" & str99
    If Len(str99) = 3 Then AmbilKodeStokBaru = "ST-" + strtgl & "." + str99
        
    OBJ.Close
    Exit Function
err_msg:
   
    MsgBox Err.Description
    
End Function


