VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt_trans_in 
   Caption         =   "Adding Cash/Bank In"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   22251
   _ExtentY        =   14049
   SectionData     =   "rpt_trans_in.dsx":0000
End
Attribute VB_Name = "rpt_trans_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Dim StrSQL As String
    
    With DataControl1.Recordset
        If Not .EOF Then
                
            StrSQL = "Select a.noactrx,a.desctrx,a.amounttrx,a.cekbg,b.nmac "
            StrSQL = StrSQL + "From gl_transaksi as a inner join gl_masterac as b "
            StrSQL = StrSQL + "on b.noac=a.noactrx "
            StrSQL = StrSQL + "Where a.notrx = '" + .Fields("notrx") + "' and a.dbkrtrx = 'K' and (a.kdtrx = 'KM' or a.kdtrx = 'BM')"

            Set subrpt.object = New subrpt_trans_in
            With subrpt.object
                .DataControl1.Source = StrSQL
                .DataControl1.ConnectionString = dsn
            End With
        End If
    End With
End Sub

