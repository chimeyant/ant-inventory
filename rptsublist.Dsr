VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptsublist 
   Caption         =   "ActiveReport1"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13260
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   23389
   _ExtentY        =   9260
   SectionData     =   "rptsublist.dsx":0000
End
Attribute VB_Name = "rptsublist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If i = 0 Then i = 1
        Field1.DataValue = Str(i) + "."
    i = i + 1
End Sub
