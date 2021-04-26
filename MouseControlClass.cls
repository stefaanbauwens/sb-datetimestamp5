VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseControlClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hWnd As Long
Public IsHovered As Boolean
Public overOnly As Boolean
Public HoverTime As Long
Public prevProc As Long

Public Event MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)
Public Event WheelScroll(MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)
Public Event HoverChange(IsHovered As Boolean)

Public Sub RaiseHoverChange(IsHovered As Boolean)

RaiseEvent HoverChange(IsHovered)

End Sub

Public Sub RaiseMouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

RaiseEvent MouseMove(MouseState, MouseX, MouseY)

End Sub

Public Sub RaiseWheelScroll(MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)

RaiseEvent WheelScroll(MouseKeys, MouseRotation, MousePosX, MousePosY)

End Sub

Private Sub Class_Initialize()

HoverTime = 400
IsHovered = False
overOnly = True

End Sub
