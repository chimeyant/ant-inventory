VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpttandaterima 
   Caption         =   "TANDA TERIMA"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20370
   Icon            =   "rpttandaterima.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "rpttandaterima.dsx":000C
End
Attribute VB_Name = "rpttandaterima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim jumlah As Double
Dim total As Double
Dim huruf As String
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If i = 0 Then i = 1
        Field1.DataValue = Str(i) + "."
    i = i + 1
    If EOF = False Then
        jumlah = Getnilai(DataControl1.Recordset!noapply)
        total = total + jumlah
        Field4 = Format(jumlah, "#,##0")
        Field5 = Format(total, "#,##0")
        Field6 = frmdaftartagih.grid.TextMatrix(1, 4)
    End If
    huruf = ANGKAKEHURUF(Format(total, "general number")) & " Rupiah"
    lblterbilang = huruf
End Sub

