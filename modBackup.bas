Attribute VB_Name = "modBackup"
Option Explicit

Private OBJ As ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String

'proses hapus data sales order
Function hapus_salesorder() As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_sohdr where noso like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_solin where noso like 'L%'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_salesorder = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
    hapus_salesorder = False
End Function

'proses hapus data surat jalan
Function hapus_suratjalan() As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_sjhdr where nosj like 'LL%'"
    OBJ.Execute SQL
    SQL = "delete from am_sjlin where nosj like 'LL%'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_suratjalan = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_suratjalan = False
End Function

'proses hapus invoice
Function hapus_invoice() As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_invhdr where nobkt like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_invlin where nobkt like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_invlinlin where nobkt like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_invdesc where noinv like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_invdelete where nobkt like 'L%'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_invoice = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_invoice = False
End Function

'proses hapus data pembayaran
Function hapus_pembayaran() As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_cashhdr where nobkt like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_cashlin where nobkt like 'L%'"
    OBJ.Execute SQL
    SQL = "delete from am_cashsub where nobkt like 'L%'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_pembayaran = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_pembayaran = False
End Function

'proses hapus data ledger penjualan
Function hapus_ledger_penjualan() As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    SQL = "delete from gl_transaksi where desctrx like 'L0%'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_ledger_penjualan = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_ledger_penjualan = False
End Function

Function hapus_po(ByVal nopo As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    SQL = "delete from am_pohdr where nopo = '" & nopo & "'"
    OBJ.Execute SQL
    SQL = "delete from am_polin where nopo = '" & nopo & "'"
    OBJ.Execute SQL
    SQL = "delete from am_podrop where  nopo ='" & nopo & "'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_po = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_po = False
End Function

Function hapus_penerimaan(ByVal noBPB As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_belihdr where nobeli ='" & noBPB & "'"
    OBJ.Execute SQL
    SQL = "delete from am_belilin where nobeli ='" & noBPB & "'"
    OBJ.Execute SQL
    SQL = "delete from am_beliapp where nobeli ='" & noBPB & "'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_penerimaan = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_penerimaan = False
End Function

Function hapus_voucher(ByVal novoucher As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_voucherhdr where novoucher='" & novoucher & "'"
    OBJ.Execute SQL
    SQL = "delete from am_voucherin where novoucher ='" & novoucher & "'"
    OBJ.Close
    hapus_voucher = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_voucher = False
End Function

Function hapus_pembayaran_pembelian_apopnfil(ByVal noapply As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_apopnfil where noapply ='" & noapply & "'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_pembayaran_pembelian_apopnfil = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    hapus_pembayaran_pembelian_apopnfil = False
End Function

Function hapus_pembayaran_pembelian_apcash(ByVal noapply As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "delete from am_apcashhdr where noapply ='" & noapply & "'"
    OBJ.Execute SQL
    SQL = "delete from am_apcashlin where noapply ='" & noapply & "'"
    OBJ.Execute SQL
    SQL = "delete from am_apcashlinppn where noapply ='" & noapply & "'"
    OBJ.Execute SQL
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function

Function hapus_pembayaran_pembelian_apcashsub(ByVal nobukti As String) As Boolean
    On Error GoTo err_handler:
    Set OBJ = New ADODB.Connection
    SQL = "delete from am_apcashsub where nobukti = '" & nobukti & "'"
    OBJ.Execute SQL
    OBJ.Close
    hapus_pembayaran_pembelian_apcashsub = True
    Exit Function
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function

Function hapus_ledger_pembelian() As Boolean
    On Error GoTo err_handler:
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Function


