Attribute VB_Name = "ModScrollGrid"

Option Explicit

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157


Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
''''''''''''' Scroll stuff
'no 1
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(Lwnd)
      If fFrm Is Nothing Then
        ' it's not a form
        If Not IsOver(Lwnd, Xpos, Ypos) And IsOver(GetParent(Lwnd), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetParent(Lwnd)).MouseWheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.hWnd, Xpos, Ypos) Then fFrm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
      End If
  End Select
  
  WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub WheelHook(ByVal hWnd As Long)
  On Error Resume Next
  SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'Public Sub WheelUnHook(ByVal hWnd As Long)
  'On Error Resume Next
  'SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
  'RemoveProp hWnd, "PrevWndProc"
'End Sub

Public Sub HorFlexGridScroll(ByRef HFG As MSHFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim NewValue As Long
  Dim Lstep As Single

  On Error Resume Next
  With HFG
  
    Lstep = .Height / .RowHeight(0)
    Lstep = Int(Lstep)
    
    If .Rows < Lstep Then Exit Sub
    Do While Not (.RowIsVisible(.TopRow + Lstep))
      Lstep = Lstep - 1
    Loop
    If Rotation > 0 Then
        NewValue = .TopRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .TopRow + Lstep
        If NewValue > .Rows - 1 Then
            NewValue = .Rows - 1
        End If
    End If
    .TopRow = NewValue
  End With
End Sub

Public Function IsOver(ByVal hWnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hWnd, rectCtl
  With rectCtl
    If lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom Then IsOver = True
  End With
End Function

Private Function GetForm(ByVal hWnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hWnd = hWnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

'''''' End of scroll

''''''''''''''''''''''''''''''''''''''''''
