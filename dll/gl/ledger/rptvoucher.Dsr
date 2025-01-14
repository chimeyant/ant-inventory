VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBB 
   Caption         =   "Print Bank Payment"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   Icon            =   "rptvoucher.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19288
   SectionData     =   "rptvoucher.dsx":08CA
End
Attribute VB_Name = "rptBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_FetchData(EOF As Boolean)
    Field32 = Format(Field32, "###,##0.00")
    Field27 = Format(Field23, "###,##0.00")
    Field23 = Format(Field27, "###,##0.00")
    Filed28 = Format(Field28, "###,##0.00")
End Sub

Private Sub GroupFooter1_Format()
    Field27 = Format(Field23, "###,##0.00")
    Field23 = Format(Field27, "###,##0.00")
    Filed28 = Format(Field28, "###,##0.00")
End Sub

Private Sub GroupHeader1_Format()
    Field27 = Format(Field23, "###,##0.00")
    Field23 = Format(Field27, "###,##0.00")
    Filed28 = Format(Field28, "###,##0.00")
End Sub

