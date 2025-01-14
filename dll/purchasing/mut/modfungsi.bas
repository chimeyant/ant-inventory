Attribute VB_Name = "modfungsi"
Option Explicit

Private OBJ As ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String

Function GetSaldoBahanBaku(ByVal s_kodebarang As String, ByVal s_qty As Integer) As Double
        On Error GoTo err_handler
        Dim saldo As Double
        saldo = 0
        
        Set OBJ = New ADODB.Connection
        
        OBJ.Open dsn
        SQL = "select a.nopo,a.tglpo ,b.kodebarang,b.qtyuse,b.price "
        SQL = SQL + "am_pohdr a inner join am_pohdr b  on a.nopo = b.nopo "
        SQL = SQL + "where b.kodebarang= '" & s_kodebarang & "'"
        SQL = SQL + "order by a.tglpo asc"
        
        Set RS = OBJ.Execute(SQL)
        Do While Not RS.EOF
            If s_qty > RS!qty Then
                saldo = saldo + (RS!qty * RS!price)
                s_qty = s_qty - RS!qty
            Else
                saldo = saldo + (s_qty * RS!price)
                Exit Do
            End If
            RS.MoveNext
        Loop
        OBJ.Close
        GetSaldoBahanBaku = saldo
        Exit Function
err_handler:
        OBJ.Close
        MsgBox "Gagal mengambil data saldo stok bahan baku...! " & Err.Description, vbCritical, AppName
End Function
