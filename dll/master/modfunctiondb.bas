Attribute VB_Name = "modfunctiondb"
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
End Sub

Sub OpenDB()
    OpenSQLDB dbServer, dbName, dbUser, dbPass
End Sub

Function Display_Data(ByVal A_RS As ADODB.Recordset) As ADODB.Recordset
    'declare variable
    Dim i As Integer
    Dim j As Integer
    
    Set RS = New ADODB.Recordset
    
    For i = 0 To A_RS.Fields.Count - 1
        RS.Fields.Append A_RS.Fields(i).Name, adChar, 255, adFldIsNullable
    Next
    RS.Open
    
    Do While Not A_RS.EOF
        RS.AddNew
            For i = 0 To A_RS.Fields.Count - 1
                RS.Fields(i).Value = A_RS.Fields(i).Value
            Next
            RS.Update
        A_RS.MoveNext
        DoEvents
    Loop
       
    Set Display_Data = RS
End Function

Function KodeDept(ByVal Departemen As String) As String
    On Error GoTo err_handler
    Dim SQL As String
    SQL = "select kode_dept from am_apdepartemen where dept='" + Departemen + "'"
    OpenDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    KodeDept = RS!kode_dept
    CloseSQLDB
    Exit Function
err_handler:
    MsgBox "Gagal melakukan pengambilan kode departemen...!. " + Err.Description, vbCritical, "Warning"
End Function

Function NamaDepartemen(ByVal KodeDepartemen As String) As String
    On Error GoTo err_handler
    Dim SQL As String
    
    SQL = "Select dept from am_apdepartemen where kode_dept='" + KodeDepartemen + "'"
    OpenDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    NamaDepartemen = RS!dept
    CloseSQLDB
    Exit Function
err_handler:
    MsgBox "Gagal melakukan pengambilan nama departemen...!. " + Err.Description, vbCritical, "Warning"
End Function

Function NamaLevel(ByVal KodeDepartemen, KodeLevel As String) As String
    On Error Resume Next
    Dim SQL As String
    
    SQL = "select nmlevel from am_aplevel where kode_dept='" + KodeDepartemen + "' and kode_level='" + KodeLevel + "'"
    OpenDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    NamaLevel = RS!nmlevel
    CloseSQLDB
    Exit Function
err_handler:
    MsgBox "Gagal melakukan pengambilan nama level....!. " + Err.Description, vbCritical, "Warning"
End Function

Function KodeLevel(ByVal KodeDepartemen, Level As String) As String
    On Error GoTo err_handler
    Dim SQL As String
    
    SQL = "select distinct kode_level from am_aplevel where kode_dept='" + KodeDepartemen + "' and nmlevel='" + Level + "'"
    OpenDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    KodeLevel = RS!kode_level
    Exit Function
err_handler:
    MsgBox "Gagal melakukan pengambilan kode level...!." + Err.Description, vbCritical, "Warning"
End Function
