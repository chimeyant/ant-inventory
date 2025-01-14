VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptpermintaanlist 
   Caption         =   "List Permintaan Barang"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   25135
   _ExtentY        =   13996
   SectionData     =   "rptpermintaanlist.dsx":0000
End
Attribute VB_Name = "rptpermintaanlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Dim StrSQL As String
    
    With DataControl1.Recordset
        If Not .EOF Then
                
            StrSQL = "Select * From am_perminin "
            StrSQL = StrSQL + "Where nobkt = '" + .Fields("nobkt") + "'"
            Set subrpt.object = New rptsublist
            With subrpt.object
                .DataControl1.Source = StrSQL
                .DataControl1.ConnectionString = dsn
            End With
        End If
    End With
End Sub
