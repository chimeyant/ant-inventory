Attribute VB_Name = "modvariable"
'Declared Information Aplication

Global Const FullAppName = "Ant Inventory System"
Global Const AppName = "Ant Inventory System"
Global Const AppVer = "Ver. 1.02"
Global Const AppProgrammer1 = "U. Selamat Raharja"
Global Const AppProgrammer2 = "Chandra Kirana"
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
Public hasil, hasil1, hasil2, hasil3 As String
Public report1, report2, report3, report4, report5 As String
Public dsn, dsn1, kuser, kcomp, nmuser As String
Public setup1, setup2, setup3, setup4 As String
Public par1, par2, par3, par4, par5 As String
Public kar_1, kar_2, kar_3, kar_4 As String
Public w, z As Integer
Public x1, y1 As Integer
Public ops_tf, ops_tf1, ops_ct As Boolean
Public myArray(10, 1) As String
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


Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function



