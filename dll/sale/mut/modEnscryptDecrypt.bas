Attribute VB_Name = "modEnscryptDecrypt"
'Program Name   : Exclusive Inventory Technology System
'Alias          : EI-Tech System
'Copyright      : 2012
'Company        : SPARTA PRIMA
'Programmer     : U. Selamat Raharja & Adnan Nur Kholik

Option Explicit

Function Cheap_Encrypt(text As String) As String
'Declaring variables
Dim i As Integer
Dim X As String
Dim fin As String
'This is an example of a cheap ass way to encrypt something
'i did use this code in this example
'i just have it here to show you the cheezy way =P

'I know its alot easier and alot shorter
'and alot easier a hell of a lot easier to program
'but this is the crap way to encrypt text
'
'most encryptors i see these days do just this too
'including PAT or JK's..except this is alot shorter code
'all this bascially does it swap ASCII positions of
'the different characters within the string..kinda crappy

For i% = 1 To Len(text)
    X = Mid(text, i, 1) 'grapping individual character from string
    'this is simple.
    'it swaps the ASCII of the characeter with another
    'which is what all encrypters do..pretty gay
    fin$ = fin$ & Chr$(255 - Asc(X))
Next i%
'now wasn't that easy
Cheap_Encrypt$ = fin$

'I'm not ripping on anybody here but try better ways to encrypt
'text. I didn't take much thinking to create this code
'the more advanced the code is the harder something is
'going to be to decrypt
'well i dunno
'i guess this could be usefull in some cases..like
'encrypting passwords in an INI like Pat or JK said
'but even so the other encrpytion code would be alot
'harder to decrypt

End Function

Function Cheap_Decrypt(text As String) As String
'Declaring variables
Dim i As Integer
Dim X As String
Dim fin As String
'This is an example of a cheap ass way to encrypt something
'i did use this code in this example
'i just have it here to show you the cheezy way =P

'I know its alot easier and alot shorter
'and alot easier a hell of a lot easier to program
'but this is the crap way to encrypt text
'
'most encryptors i see these days do just this too
'including PAT or JK's..except this is alot shorter code
'all this bascially does it swap ASCII positions of
'the different characters within the string..kinda crappy

For i% = 1 To Len(text)
    X = Mid(text, i, 1) 'grapping individual character from string
    'this is simple.
    'it swaps the ASCII of the characeter with another
    'which is what all encrypters do..pretty gay
    fin$ = fin$ & Chr$(255 - Asc(X))
Next i%
'now wasn't that easy
Cheap_Decrypt$ = fin$


'I'm not ripping on anybody here but try better ways to encrypt
'text. I didn't take much thinking to create this code
'the more advanced the code is the harder something is
'going to be to decrypt
'well i dunno
'i guess this could be usefull in some cases..like
'encrypting passwords in an INI like Pat or JK said
'but even so the other encrpytion code would be alot
'harder to decrypt
End Function






