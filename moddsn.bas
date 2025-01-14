Attribute VB_Name = "moddsn"
Option Explicit
Private Const REG_SZ = 1    'Constant for a string variable type.
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
       "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, _
       phkResult As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, _
       ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
       cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal Hkey As Long) As Long
       
Public Sub AddFirebirdDSN(Optional ByVal S_DataSourceName As String, _
Optional ByVal S_AutoQoutedIdentifier As String = "N", _
Optional ByVal S_CharacterSet As String = "NONE", _
Optional ByVal S_Client As String, _
Optional ByVal S_DbName As String, _
Optional ByVal S_Description As String, _
Optional ByVal S_Dialec As String = "3", _
Optional ByVal S_Driver As String, _
Optional ByVal S_JdbcDriver As String = "iscDbc", _
Optional ByVal S_LockTimerWaitTrans As String, _
Optional ByVal S_NoWait As String = "N", _
Optional ByVal S_Password As String, _
Optional ByVal S_QuotedIdentifier As String = "Y", _
Optional ByVal S_ReadOnly As String = "N", _
Optional ByVal S_Role As String, _
Optional ByVal S_SafeThread As String = "Y", _
Optional ByVal S_SensitiveIdentifier As String = "N", _
Optional ByVal S_User As String = "sysdba", _
Optional ByVal S_UseSchemaIdentifier As String = "0")

    Dim DataSourceName As String
    Dim AutoQoutedIdentifier As String
    Dim CharacterSet As String
    Dim Client As String
    Dim DBName As String
    Dim Description As String
    Dim Dialect As String
    Dim Driver As String
    Dim JdbcDriver As String
    Dim LockTimeoutWaitTransactions As String
    Dim NoWait As String
    Dim Password As String
    Dim QuotedIdentifier As String
    Dim ReadOnly As String
    Dim Role As String
    Dim SafeThread As String
    Dim SensitiveIdentifier As String
    Dim User As String
    Dim UseSchemaIdentifier As String
    Dim lResult As Long
    Dim hKeyHandle As Long

   'Specify the DSN parameters.
    DataSourceName = S_DataSourceName
    AutoQoutedIdentifier = S_AutoQoutedIdentifier
    CharacterSet = S_CharacterSet
    Client = S_Client
    DBName = S_DbName
    Description = S_Description
    Dialect = S_Dialec
    Driver = S_Driver
    JdbcDriver = S_JdbcDriver
    LockTimeoutWaitTransactions = S_LockTimerWaitTrans
    NoWait = S_NoWait
    Password = S_Password
    QuotedIdentifier = S_QuotedIdentifier
    ReadOnly = S_ReadOnly
    Role = S_Role
    SafeThread = S_SafeThread
    SensitiveIdentifier = S_SensitiveIdentifier
    User = S_User
    UseSchemaIdentifier = S_UseSchemaIdentifier
    
    'Create the new DSN key.
    lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

    'Set the values of the new DSN key.
    lResult = RegSetValueEx(hKeyHandle, "AutoQuotedIdentifier", 0&, REG_SZ, _
      ByVal AutoQoutedIdentifier, Len(AutoQoutedIdentifier))
      
    lResult = RegSetValueEx(hKeyHandle, "CharacterSet", 0&, REG_SZ, _
      ByVal CharacterSet, Len(CharacterSet))
    
    lResult = RegSetValueEx(hKeyHandle, "Client", 0&, REG_SZ, _
      ByVal Client, Len(Client))
    
    lResult = RegSetValueEx(hKeyHandle, "Dbname", 0&, REG_SZ, _
      ByVal DBName, Len(DBName))
    
    lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
    
    lResult = RegSetValueEx(hKeyHandle, "Dialect", 0&, REG_SZ, _
      ByVal Dialect, Len(Dialect))
    
    lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal Driver, Len(Driver))
    
    lResult = RegSetValueEx(hKeyHandle, "JdbcDriver", 0&, REG_SZ, _
      ByVal JdbcDriver, Len(JdbcDriver))
    
    lResult = RegSetValueEx(hKeyHandle, "LockTimeoutWaitTransactions", 0&, REG_SZ, _
      ByVal LockTimeoutWaitTransactions, Len(LockTimeoutWaitTransactions))
      
    lResult = RegSetValueEx(hKeyHandle, "NoWait", 0&, REG_SZ, _
      ByVal NoWait, Len(NoWait))
      
    lResult = RegSetValueEx(hKeyHandle, "Password", 0&, REG_SZ, _
      ByVal Password, Len(Password))
      
    lResult = RegSetValueEx(hKeyHandle, "QuotedIdentifier", 0&, REG_SZ, _
      ByVal QuotedIdentifier, Len(QuotedIdentifier))
        
    lResult = RegSetValueEx(hKeyHandle, "ReadOnly", 0&, REG_SZ, _
      ByVal ReadOnly, Len(ReadOnly))
        
    lResult = RegSetValueEx(hKeyHandle, "Role", 0&, REG_SZ, _
      ByVal Role, Len(Role))
        
    lResult = RegSetValueEx(hKeyHandle, "SafeThread", 0&, REG_SZ, _
      ByVal SafeThread, Len(SafeThread))
      
    lResult = RegSetValueEx(hKeyHandle, "SensitiveIdentifier", 0&, REG_SZ, _
      ByVal SensitiveIdentifier, Len(SensitiveIdentifier))
      
    lResult = RegSetValueEx(hKeyHandle, "User", 0&, REG_SZ, _
      ByVal User, Len(User))
    
    lResult = RegSetValueEx(hKeyHandle, "UseSchemaIdentifier", 0&, REG_SZ, _
      ByVal UseSchemaIdentifier, Len(UseSchemaIdentifier))
    
   'Close the new DSN key.
   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal "Firebird/InterBase(r) driver", Len("Firebird/InterBase(r) driver"))
   lResult = RegCloseKey(hKeyHandle)

End Sub

Public Sub AddMySQLDSN(ByVal DSNName As String, ByVal DBName As String, Optional DSNDescription As String, _
Optional ByVal PathDriverName As String = "C:\Program Files\MySQL\Connector ODBC 5.1\myodbc5.dll", _
Optional ByVal DriverName As String = "MySQL ODBC 5.1 Driver", _
Optional ByVal Server As String = "localhost", Optional ByVal Port As String = "3306", _
Optional ByVal UID As String = "")


    Dim lResult As Long
    Dim hKeyHandle As Long

    'Create the new DSN key.
    lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DSNName, hKeyHandle)
    
    'Create the new DSN Property
    lResult = RegSetValueEx(hKeyHandle, "DATABASE", 0&, REG_SZ, _
      ByVal DBName, Len(DBName))
    
    lResult = RegSetValueEx(hKeyHandle, "DESCRIPTION", 0&, REG_SZ, _
      ByVal DSNDescription, Len(DSNDescription))
    
    lResult = RegSetValueEx(hKeyHandle, "DRIVER", 0&, REG_SZ, _
      ByVal PathDriverName, Len(PathDriverName))
    
    lResult = RegSetValueEx(hKeyHandle, "PORT", 0&, REG_SZ, _
      ByVal Port, Len(Port))
    
    lResult = RegSetValueEx(hKeyHandle, "SERVER", 0&, REG_SZ, _
      ByVal Server, Len(Server))
      
    lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, _
      ByVal UID, Len(UID))
      
    
    'Close the new DSN key.
   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DSNName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)
    
End Sub

Public Sub AddSQLSERVERDSN(ByVal DSNName As String, ByVal DBName As String, Optional DSNDescription As String, _
Optional ByVal PathDriverName As String = "C:\Windows\system32\Sqlsrv32.dll", _
Optional ByVal DriverName As String = "SQL Server", _
Optional ByVal Server As String = "localhost", _
Optional ByVal UID As String = "")


    Dim lResult As Long
    Dim hKeyHandle As Long

    'Create the new DSN key.
    lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DSNName, hKeyHandle)
    
    'Create the new DSN Property
    lResult = RegSetValueEx(hKeyHandle, "DATABASE", 0&, REG_SZ, _
      ByVal DBName, Len(DBName))
    
    lResult = RegSetValueEx(hKeyHandle, "DESCRIPTION", 0&, REG_SZ, _
      ByVal DSNDescription, Len(DSNDescription))
    
    lResult = RegSetValueEx(hKeyHandle, "DRIVER", 0&, REG_SZ, _
      ByVal PathDriverName, Len(PathDriverName))
    
    
    lResult = RegSetValueEx(hKeyHandle, "SERVER", 0&, REG_SZ, _
      ByVal Server, Len(Server))
      
    lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal UID, Len(UID))
      
    
    'Close the new DSN key.
   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DSNName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)
    
End Sub





