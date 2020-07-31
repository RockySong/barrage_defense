Attribute VB_Name = "mdlMouse"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As String) As Long

Public Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
        As Long, ByVal dwNewLong As Long) As Long
        
Public Declare Function CallWindowProc Lib "user32" Alias _
        "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal _
        hwnd As Long, ByVal Msg As Long, ByVal wParam As _
        Long, ByVal lParam As Long) As Long
  
Const GWL_WNDPROC = (-4&)

Dim PrevWndProc&

Private Const WM_DESTROY = &H2


Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long

Public Const TME_CANCEL = &H80000000
Public Const TME_HOVER = &H1&
Public Const TME_LEAVE = &H2&
Public Const TME_NONCLIENT = &H10&
Public Const TME_QUERY = &H40000000

Private Const WM_MOUSELEAVE = &H2A3&

Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public bTracking As Boolean

Public Sub InitWndProc(hwnd As Long)
  PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub
Public Function SubWndProc(ByVal hwnd As Long, ByVal Msg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) _
                            As Long

   If Msg = WM_DESTROY Then Terminate (hwnd)

   '处理鼠标移出消息
   If Msg = WM_MOUSELEAVE Then
      bTracking = False
      Stop
   End If
   SubWndProc = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
End Function


Public Sub Terminate(hwnd As Long)
  Call SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
End Sub
