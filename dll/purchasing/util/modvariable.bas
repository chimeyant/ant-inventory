Attribute VB_Name = "modvariable"
'Declared Information Aplication

Global Const FullAppName = "Ant Inventory System"
Global Const AppName = "Ant Inventory System"
Global Const AppVer = "Ver. 1.02"
Global Const AppProgrammer1 = "U. Selamat Raharja"
Global Const Bahasa = "IND"
Global AppPath As String

Global UserOnline As String
Global UserOnLineLevel As String
Global UserOnlineDept As String
Global TglServer As String

'Declare Database Property
Global dbServer As String
Global dbName As String
Global dbPort As String
Global dbUser As String
Global dbPass As String
Global remoteserver As Boolean
Global dbRemoteServer As String

'Other Declare Variable
Public namatabel As String
Public carisql1 As String
Public hasil, hasil1, hasil2, hasil3, hasil4 As String
Public report1, report2, report3, report4 As String
Public dsn, kuser, kcomp, nmuser As String
Public setup1 As Boolean
Public myArray(10, 1) As String
Public x1, y1 As Integer

Public v_fastsearching As Boolean
Public v_fstgl1 As Date
Public v_fstgl2 As Date


Public Const LOCALE_SSHORTDATE As Long = &H1F
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long

Public Function Blok(Txt As Object) As String
    Txt.SelStart = 0
    Txt.SelLength = Len(Txt)
End Function

Public Function GetTheComputerName() As String
On Error GoTo ErrorHandlermodule

    Dim strComputerName As String ' Variable to return the path of computer name
    
    strComputerName = Space(250) ' Initilize the buffer to receive the string
    GetComputerName strComputerName, Len(strComputerName)
    strComputerName = Mid(Trim$(strComputerName), 1, Len(Trim$(strComputerName)) - 1)
    GetTheComputerName = strComputerName

    Exit Function
 
ErrorHandlermodule:
    Err.Raise Err.Number, Err.Source & "/Utils.GetTheComputerName", Err.Description
End Function



