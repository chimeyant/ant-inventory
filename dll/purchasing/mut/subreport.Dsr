VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} subreport 
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "subreport.dsx":0000
End
Attribute VB_Name = "subreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If EOF = False Then
        If DataControl1.Recordset!tsaldo <> 0 Then
            Field2.text = Format(DataControl1.Recordset!tsaldo / DataControl1.Recordset!tqty, "##,###,###,##0.00")
        Else
            Field2.text = "0"
        End If
    End If
End Sub

