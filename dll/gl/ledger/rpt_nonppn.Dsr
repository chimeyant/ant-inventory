VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt_nonppn 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   23442
   _ExtentY        =   13811
   SectionData     =   "rpt_nonppn.dsx":0000
End
Attribute VB_Name = "rpt_nonppn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Dim StrSQL As String
    
    With DataControl1.Recordset
        If Not .EOF Then
                
            StrSQL = "Select a.noactrx,a.desctrx,a.amounttrx,a.nilaitrx,b.no_voucher "
            StrSQL = StrSQL + "From gl_transaksi as a inner join no_bank_payment as b "
            StrSQL = StrSQL + "On a.notrx = b.notrx "
            StrSQL = StrSQL + "Where b.no_voucher = '" + .Fields("no_voucher") + "' and a.dbkrtrx = 'D' and a.idupdate = '2' and b.flag = '2'"
            Set subrpt.object = New subrpt_nonppn
            With subrpt.object
                .DataControl1.Source = StrSQL
                .DataControl1.ConnectionString = dsn
            End With
        End If
    End With
End Sub

