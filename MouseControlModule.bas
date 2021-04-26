Attribute VB_Name = "MouseControlModule"

Option Explicit

Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const TME_HOVER = &H1
Private Const TME_LEAVE = &H2
Private Const TME_CANCEL = &H80000000
Private Const HOVER_DEFAULT = &HFFFFFFFF
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEHOVER = &H2A1
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSEWHEEL = &H20A
Private Const GWL_WNDPROC = (-4)

Dim trackCol As Collection

Public Function StartTrack(trak As MouseControlClass)

Dim prevProc As Long

If trackCol Is Nothing Then
    Set trackCol = New Collection
End If

trak.prevProc = SetWindowLong(trak.hWnd, GWL_WNDPROC, AddressOf WindowProc)

trackCol.Add trak, CStr(trak.hWnd)

RequestTracking trak

End Function

Public Function EndTrack(trak As MouseControlClass)

Dim trk As tagTRACKMOUSEEVENT

If trackCol Is Nothing Then Exit Function

Call SetWindowLong(trak.hWnd, GWL_WNDPROC, trak.prevProc)

trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER Or TME_CANCEL
trk.hwndTrack = trak.hWnd

TrackMouseEvent trk

On Error Resume Next
trackCol.Remove CStr(trak.hWnd)
On Error GoTo 0

If trackCol.Count = 0 Then
    Set trackCol = Nothing
End If

End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim MouseKeys As Long
Dim MouseRotation As Long
Dim MousePosX As Long
Dim MousePosY As Long

Dim OverFlag As Boolean

Dim trak As MouseControlClass

    On Error GoTo resumewindowproc

    Set trak = trackCol.Item(CStr(hWnd))

    If uMsg = WM_MOUSEWHEEL Then

        MouseKeys = wParam And 65535
        MouseRotation = wParam / 65536
        MousePosX = lParam And 65535
        MousePosY = lParam / 65536

        OverFlag = Not trak.overOnly

        If OverFlag = False Then OverFlag = ReturnIsMouseOver(hWnd, MousePosX, MousePosY)

        If OverFlag = True Then
            trak.RaiseWheelScroll MouseKeys, MouseRotation, MousePosX, MousePosY
        End If

    ElseIf uMsg = WM_MOUSELEAVE Then
        
        trak.RaiseHoverChange False
        trak.IsHovered = False
    
    ElseIf uMsg = WM_MOUSEHOVER Then
        
        CheckHovering trak
    
    ElseIf uMsg = WM_MOUSEMOVE Then
        
        CheckHovering trak
        
        RequestTracking trak
        
        MousePosX = lParam And 65535
        MousePosY = lParam / 65536

        WindowProc = CallWindowProc(trak.prevProc, hWnd, uMsg, wParam, lParam)
    
        trak.RaiseMouseMove wParam, MousePosX, MousePosY
    
    Else
        
        WindowProc = CallWindowProc(trak.prevProc, hWnd, uMsg, wParam, lParam)
    
    End If

    Exit Function

resumewindowproc:
    Debug.Print Err.Description

End Function

Private Function RequestTracking(trak As MouseControlClass)

Dim trk As tagTRACKMOUSEEVENT

trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER
trk.dwHoverTime = trak.HoverTime
trk.hwndTrack = trak.hWnd

TrackMouseEvent trk

End Function

Private Function CheckHovering(trak As MouseControlClass)
        
If trak.IsHovered = False Then
    trak.IsHovered = True
    trak.RaiseHoverChange True
End If

End Function

Public Function ReturnIsMouseOver(ByVal hWnd As Long, ByVal MousePosX As Long, ByVal MousePosY As Long) As Boolean

Dim ControlRect As RECT

GetWindowRect hWnd, ControlRect

With ControlRect
    ReturnIsMouseOver = (MousePosX >= .Left And MousePosX <= .Right And MousePosY >= .Top And MousePosY <= .Bottom)
End With

End Function

Public Sub GetUserControlOffset(ByVal hWnd As Long, MouseOffsetX As Long, MouseOffsetY As Long)

Dim ControlRect As RECT

GetWindowRect hWnd, ControlRect

MouseOffsetX = ControlRect.Left
MouseOffsetY = ControlRect.Top

End Sub
