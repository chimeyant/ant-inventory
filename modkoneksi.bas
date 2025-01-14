Attribute VB_Name = "modkoneksi"
'Program Name   : Exclusive Inventory Technology System
'Alias          : EI-Tech System
'Copyright      : 2012
'Company        : SPARTA PRIMA
'Programmer     : U. Selamat Raharja
Option Explicit

'Declare Database SQL Server Property
Public ConSQL As ADODB.Connection
Private RSSQL As ADODB.Recordset


Function OpenSQLDB(ByVal SQL_DbServer, SQL_DbName, SQL_DbUser, SQL_DbPass As String) As Boolean
    'Declare Variable
    On Error GoTo err_msg
    Dim StrKoneksi As String
    
    StrKoneksi = "Provider=SQLOLEDB.1;Password=" + SQL_DbPass + ";User ID=" + SQL_DbUser + ";Initial Catalog=" + SQL_DbName + ";Data Source=" + SQL_DbServer
    dsn = StrKoneksi
    Set ConSQL = New ADODB.Connection
    ConSQL.Open StrKoneksi
    
    OpenSQLDB = False
    
    Exit Function
err_msg:
    OpenSQLDB = False
End Function

Sub CloseSQLDB()
    If ConSQL.State <> 0 Then
        ConSQL.Close
    End If
End Sub






