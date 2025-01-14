Attribute VB_Name = "modmain"
Sub Main()
    'cek regional and date setting
    Dim tglSystem, tglApp As String
    Dim knfMsg As Integer
    
    tglSystem = Date
    tglApp = Format(Date, "dd/MM/yyyy")
    
    If tglSystem <> tglApp Then
        knfMsg = MsgBox("Silahkan Setting Terlebih Dahulu " & Chr(13) & "Region Setting anda ke English US dan System Tanggal ke dd/MM/yyyy ", vbOKCancel, AppName + " " + AppVer)
        If knfMsg = 1 Then
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,4")
        End
        End If
    End If
    
    'cek legalitas software
'    If getUnlock = False Then
'        MsgBox "Fatal Error, Please Contact Your Administrator...!!!", vbCritical, AppName
'        End
'    End If
   
    'load file setting
    LoadDatabaseProperty
    
    With frmmain
        .Caption = AppName & " " & AppVer
        .WindowState = 2
        .Show
    End With
End Sub

