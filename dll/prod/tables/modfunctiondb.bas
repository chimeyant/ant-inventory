Attribute VB_Name = "modfunctiondb"
Option Explicit

Private OBJ As ADODB.Connection
Private RS As ADODB.Recordset
Private CMD As ADODB.Command
Private PARAM As ADODB.Parameter
Private SQL As String


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

Public Function GetStokBarang(ByVal s_tanggal As String, ByVal s_barang As String, Optional ByRef namabarang As String, _
Optional ByRef namasatuan As String, Optional ByRef qty_stok As Double)
    On Error GoTo Err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    Set CMD = New ADODB.Command
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "am_posisi"
    CMD.ActiveConnection = OBJ
    Set PARAM = CMD.CreateParameter("@KODE1", adVarChar, adParamInput, 20, s_tanggal)
    CMD.Parameters.Append PARAM
    Set PARAM = CMD.CreateParameter("@KODE2", adVarChar, adParamInput, 20, s_tanggal)
    CMD.Parameters.Append PARAM
    Set PARAM = CMD.CreateParameter("@KODE11", adVarChar, adParamInput, 20, s_barang)
    CMD.Parameters.Append PARAM
    Set PARAM = CMD.CreateParameter("@KODE12", adVarChar, adParamInput, 20, s_barang)
    CMD.Parameters.Append PARAM
    Set PARAM = CMD.CreateParameter("@username", adVarChar, adParamInput, 20, nmuser)
    CMD.Parameters.Append PARAM
    Set RS = CMD.Execute
    If Not RS.EOF Then
        namabarang = RS.Fields(1)
        namasatuan = RS.Fields(2)
        qty_stok = RS.Fields(9)
    End If
    OBJ.Close
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    namabarang = ""
    namasatuan = ""
    qty_stok = 0
End Function
