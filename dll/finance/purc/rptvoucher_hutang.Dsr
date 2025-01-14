VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptvoucher_hutang 
   Caption         =   "Processed / Unprocessed  Voucher"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "rptvoucher_hutang.dsx":0000
End
Attribute VB_Name = "rptvoucher_hutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset

Private Sub ActiveReport_FetchData(EOF As Boolean)
    Label9 = Format(Date, "dd/MM/yyyy")
    Label9 = "Printed On " & Label9 & ", By " & nmuser

    Dim strSQL As String
    With DataControl1.Recordset
        If Not .EOF Then
            If IsNull(!Status) Or !Status = "" Or !Status = "Unprocessed" Then
                Field4.ForeColor = vbRed
            Else
                Field4.ForeColor = &H0&
            End If
        End If
    End With
End Sub

Private Sub ActiveReport_PageStart()
    Field7.DataValue = "Page " & pageNumber
End Sub
