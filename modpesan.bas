Attribute VB_Name = "modpesan"
Global Const msg_sukses_simpan = "Proses simpan data berhasil... "
Global Const msg_sukses_update = "Proses rubah data berhasil... "
Global Const msg_sukses_hapus = "Proses hapus data berhasil... "

Global Const msg_konfirmasi_update = "Apakah anda yakin akan merubah data tersebut.. "
Global Const msg_konfirmasi_hapus = "Apakah anda yakin akan menghapus data tersebut... "


Global Const msg_err_dataisian = "Pengisian data tidak lengkap.. "
Global Const msg_err_emptypass = "Kata sandi tidak boleh.."
Global Const msg_err_password_tidak_sama = "Kata sandi tidak sama...."
Global Const msg_err_penggunaada = "Pengguna telah ada... "
Global Const msg_err_simpan = "Proses simpan tidak berhasil... "
Global Const msg_err_update = "Proses rubah tidak berhasil... "
Global Const msg_err_hapus = "Proses hapus tidak berhasil... "
Global Const msg_err_user_denied = "User tidak dizinkan...! "

Function ANGKAKEHURUF(ByVal n As Currency) As String
  Dim SAT As Variant
  
  SAT = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
  Select Case n
    Case 0 To 11
        ANGKAKEHURUF = " " + SAT(Fix(n))
    Case 12 To 19
        ANGKAKEHURUF = ANGKAKEHURUF(n Mod 10) + " Belas"
    Case 20 To 99
        ANGKAKEHURUF = ANGKAKEHURUF(Fix(n / 10)) + " Puluh" + ANGKAKEHURUF(n Mod 10)
    Case 100 To 199
        ANGKAKEHURUF = " Seratus" + ANGKAKEHURUF(n - 100)
    Case 200 To 999
        ANGKAKEHURUF = ANGKAKEHURUF(Fix(n / 100)) + " Ratus" + ANGKAKEHURUF(n Mod 100)
    Case 1000 To 1999
        ANGKAKEHURUF = " Seribu" + ANGKAKEHURUF(n - 1000)
    Case 2000 To 999999
        ANGKAKEHURUF = ANGKAKEHURUF(Fix(n / 1000)) + " Ribu" + ANGKAKEHURUF(n Mod 1000)
    Case 1000000 To 999999999
        ANGKAKEHURUF = ANGKAKEHURUF(Fix(n / 1000000)) + " Juta" + ANGKAKEHURUF(n Mod 1000000)
    Case Else
        ANGKAKEHURUF = ANGKAKEHURUF(Fix(n / 1000000000)) + " Milyar" + ANGKAKEHURUF(n Mod 1000000000)
  End Select
End Function

Public Function original(ByVal sumber As String) As String
    org3 = 0
    org4 = format_coa
    org6 = 0
    Do Until org3 > Len(Trim(sumber)) - 1
        org5 = Mid(org4, org3 + 1, 1)
    
        If (org5 = "." Or org5 = "-") Then org6 = org6 + 1
        
        org3 = org3 + 1
    Loop
    org4 = Mid(org4, 1, Len(Trim(sumber)) + org6)
    
    original = Format(sumber, Replace(org4, "X", "&"))
End Function

Public Function x_original(ByVal x_sumber As String) As String
    org1 = 0
    x_original = ""
    
    Do Until org1 > Len(Trim(x_sumber)) - 1
        org2 = Mid(x_sumber, org1 + 1, 1)
    
        If Not (org2 = "." Or org2 = "-") Then x_original = x_original & org2
        
        org1 = org1 + 1
    Loop
End Function

