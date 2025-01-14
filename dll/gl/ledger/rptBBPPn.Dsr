VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBBPPn 
   Caption         =   "Print Bank Payment"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13155
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   23204
   _ExtentY        =   13917
   SectionData     =   "rptBBPPn.dsx":0000
End
Attribute VB_Name = "rptBBPPn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_FetchData(EOF As Boolean)
    Field27 = Format(Field27, "###,##0.00")
    lblppn = Format(SpyRound(lblppn), "###,###,##0.00")
    lblhutang = Format(SpyRound(lblhutang), "###,###,##0.00")
End Sub

Private Sub Detail_Format()
    Field4 = Format(SpyRound(Field4), "###,##0.00")
End Sub

Private Sub GroupFooter1_Format()
    Field27 = Format(Field27, "###,##0.00")
End Sub

Private Sub GroupHeader1_Format()
    Field27 = Format(Field27, "###,##0.00")
End Sub


Private Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.6) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRound = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRound = Val(arVal(0)) Else: SpyRound = Val(arVal(0)) + 1
End Function
Private Function SpyRoundUp(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.1) As Double
    Dim sNumber As String: Dim arVal() As String: sNumber = dNumber: If InStr(1, sNumber, ".") = 0 Then SpyRoundUp = dNumber Else: arVal = Split(sNumber, "."): sNumber = "0." & arVal(1): dNumber = Val(sNumber): If dNumber < doNotRoundUpIfLessThan Then SpyRoundUp = Val(arVal(0)) Else: SpyRoundUp = Val(arVal(0)) + 1
End Function
