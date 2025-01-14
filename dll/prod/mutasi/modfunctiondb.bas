Attribute VB_Name = "modfunctiondb"

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


Function getHPP(ByVal kode_barang As String, ByVal qty_stok As Double, ByVal qty_use As Double) As Double
    'On Error GoTo err_handler:
    Dim temp_qty As Double
    Dim temp_sisa As Double
    Dim temp_use As Double

    Dim nobpb As String
    Dim hpp As Double
    
    'cari nobpb awal hitung
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    
    SQL = "select * from view_histori_barang where kodebarang='" & kode_barang & "' order by tanggal desc"
    Set RS = OBJ.Execute(SQL)
    Do While Not RS.EOF
        temp_qty = temp_qty + RS!qty
        nobpb = RS!nobeli
        If temp_qty >= qty_stok Then Exit Do
        RS.MoveNext
    Loop
        
    SQL = "select * from view_histori_barang where nobeli >= '" & nobpb & "' and kodebarang='" & kode_barang & "' order by nobeli asc"
    Set RS = OBJ.Execute(SQL)
    Do While Not RS.EOF
        If qty_use < RS!qty Then
            hpp = hpp + (RS!harga * qty_use)
            Exit Do
        End If
        If qty_use = RS!qty Then
            hpp = hpp + (RS!qty * RS!harga)
            Exit Do
        End If
        If qty_use > RS!qty Then
            hpp = hpp + (RS!qty * RS!harga)
            qty_use = qty_use - RS!qty
        End If
        RS.MoveNext
    Loop
    OBJ.Close
    getHPP = hpp
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
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
    Set PARAM = CMD.CreateParameter("@KODE1", adVarChar, adParamInput, 20, "20150101")
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

Public Function GetTotalProduksi(ByVal s_nolot As String) As Double
    On Error GoTo Err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select sum(qty_bahan) as jml from list_produksi_child where nolot='" & s_nolot & "'"
    Set RS = OBJ.Execute(SQL)
    GetTotalProduksi = RS!jml
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    GetTotalProduksi = 0
End Function

Public Function GetToKilogram(ByVal kodebarang As String, ByVal kodesatuan As String, ByVal tanggal As String) _
As Double
    On Error GoTo Err_handler:
    Dim bulan As Integer
    Dim tahun As Integer
    Dim kilo As Double
    
    bulan = Month(tanggal)
    tahun = Year(tanggal)
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select * from am_itemkg where kodebarang='" & kodebarang & "' and kodesatuan='" & kodesatuan & "' "
    SQL = SQL + "and tahun='" & tahun & "'"
    
    Set RS = OBJ.Execute(SQL)
    If RS.EOF Then
        OBJ.Close
        GetToKilogram = 1
        Exit Function
    End If
    
    If bulan = 1 Then
        kilo = RS!kg1
    End If
    If bulan = 2 Then
        kilo = RS!kg2
    End If
    If bulan = 3 Then
        kilo = RS!kg3
    End If
    If bulan = 4 Then
        kilo = RS!kg4
    End If
    If bulan = 5 Then
        kilo = RS!kg5
    End If
    If bulan = 6 Then
        kilo = RS!kg6
    End If
    If bulan = 7 Then
        kilo = RS!kg7
    End If
    If bulan = 8 Then
        kilo = RS!kg8
    End If
    If bulan = 9 Then
        kilo = RS!kg9
    End If
    If bulan = 10 Then
        kilo = RS!kg10
    End If
    If bulan = 11 Then
        kilo = RS!kg11
    End If
    If bulan = 12 Then
        kilo = RS!kg12
    End If
    
    GetToKilogram = kilo
    OBJ.Close
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function

Function GetNoStok() As String
On Error GoTo Err_handler
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
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function

Function GetJmlRow() As String
On Error GoTo Err_handler
    SQL = "GET_JUMLAHROW"
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    Set CMD = New ADODB.Command
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = SQL
    CMD.ActiveConnection = OBJ
    Set RST = CMD.Execute
    GetJmlRow = RST!jml
    OBJ.Close
    Exit Function
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function
