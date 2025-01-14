VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptlaporansaldobahanbaku 
   Caption         =   "Laporan Saldo Bahan Baku"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "rptlaporansaldobahanbaku.dsx":0000
End
Attribute VB_Name = "rptlaporansaldobahanbaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_Format()
    Dim strSQL As String

      
    With DataControl1.Recordset
            If Not .EOF Then
                
                strSQL = "exec am_saldo_bahanbaku '" & .Fields("kodebarang") & "'," & .Fields("onhand") & ""
                Set SubReport1.object = New subreport
                With SubReport1.object
                .DataControl1.Source = strSQL
                    .DataControl1.ConnectionString = DataControl1.ConnectionString
                End With
           
                End If
            End With
End Sub
