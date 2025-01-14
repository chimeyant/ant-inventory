VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptlaporanvoucher 
   Caption         =   "Laporan Voucher"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "rptlaporanvoucher.dsx":0000
End
Attribute VB_Name = "rptlaporanvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset

Private Sub ActiveReport_FetchData(EOF As Boolean)
    Label19 = Format(Date, "dd/MM/yyyy")
    Label19 = "Printed On " & Label19 & ", By " & nmuser
    
    Dim strSQL As String
    With DataControl1.Recordset
        If Not .EOF Then
            If IsNull(!Status) Or !Status = "" Or !Status = "Unprocesed" Then
                Field7.ForeColor = vbRed
            Else
                Field7.ForeColor = &H0&
            End If
        End If
    End With
End Sub

Private Sub ActiveReport_PageStart()
    Field10.DataValue = "Page " & pageNumber
End Sub
