Attribute VB_Name = "modfunctiondb"
Private Obj As ADODB.Connection
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
    dsn1 = "Provider=SQLOLEDB.1;Password=" + dbPass + ";User ID=" + dbUser + ";Initial Catalog=" + dbName + ";Data Source=" + dbServer
End Sub

Sub OpenDB()
    OpenSQLDB dbServer, dbName, dbUser, dbPass
End Sub

Public Function dsnreport() As Variant
    dsnreport = "DSN=" & dbServer & ";UID=" & dbUser & ";PWD=" & dbPass & ";DSQ=" & dbName & ""
End Function

Public Function dsnreport1() As Variant
    dsnreport1 = "DSN=" & dbServer & ";UID=" & dbUser & ";PWD=" & dbPass & ";DSQ=" & dbName & ""
End Function

Function Getnilai(ByVal Nofaktur As String) As Double
    Dim SQL As String
    Set Obj = New ADODB.Connection
    Obj.Open dsn
    SQL = "select noapply,isnull(sum(amount+potongan+ppn+selisih),0)'nilai' "
    SQL = SQL + "from am_aropnfil where noapply like '" & Nofaktur & "%' group by noapply"
    Set RS = Obj.Execute(SQL)
    Getnilai = RS!nilai
    Obj.Close
End Function
