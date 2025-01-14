VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} sublapvoucher 
   Caption         =   "ActiveReport1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14670
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   25876
   _ExtentY        =   9393
   SectionData     =   "sublapvoucher.dsx":0000
End
Attribute VB_Name = "sublapvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset

Private Sub ActiveReport_FetchData(EOF As Boolean)
On Error Resume Next
    Dim strSQL As String
    With DataControl1.Recordset
        If Not .EOF Then
            If !perkiraan = "" Then
                OBJ.Open dsn
                strSQL = "Select namabarang From am_apitemmst Where kodebarang = '" & .Fields("keterangan") & "'"
                Set RST = OBJ.Execute(strSQL)
                Field2 = RST!namabarang
                OBJ.Close
            Else
                Field2 = !keterangan
            End If
        End If
    End With
End Sub



