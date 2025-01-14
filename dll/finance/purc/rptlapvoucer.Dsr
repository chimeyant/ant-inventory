VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptlapvoucher 
   Caption         =   "Laporan Voucher Detail"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13755
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   24262
   _ExtentY        =   13494
   SectionData     =   "rptlapvoucer.dsx":0000
End
Attribute VB_Name = "rptlapvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_Format()
    Dim strSQL As String
    With DataControl1.Recordset
        If Not .EOF Then
            strSQL = "SELECT * From am_voucherin "
            strSQL = strSQL + "Where novoucher='" + .Fields("novoucher") + "'"
            Set Subrpt.object = New sublapvoucher
            With Subrpt.object
                .DataControl1.Source = strSQL
                .DataControl1.ConnectionString = dsn
            End With
        End If
    End With
End Sub

