VERSION 5.00
Begin VB.UserControl DateTimeStamp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ScaleHeight     =   1950
   ScaleWidth      =   2550
   ToolboxBitmap   =   "DateTimeStamp.ctx":0000
End
Attribute VB_Name = "DateTimeStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Enum DTSCellTypeList
   TypeEmpty
   TypeNormal
   TypeSelected
End Enum
Enum DTSWeekdayList
    DayMonday = 0
    DayTuesday = 1
    DayWednesday = 2
    DayThursday = 3
    DayFriday = 4
    DaySaturday = 5
    DaySunday = 6
End Enum
Enum DTSHeaderModeList
    HeaderNone = 0
    HeaderSmall = 1
    HeaderNormal = 2
End Enum

Dim DTSControls As DTSControls
Dim DTSFont As StdFont

Dim WithEvents ControlTracking As MouseControlClass
Attribute ControlTracking.VB_VarHelpID = -1

Public Event DateTimeChange()
Public Event DateTimeKeyPress(KeyAscii As Integer)
Public Event DateTimeMouseUp(MouseButton As Integer)
Public Event MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)
Public Event HoverChange(IsHovered As Boolean)

Private Sub UserControl_InitProperties()

Set ControlTracking = New MouseControlClass

ControlTracking.HoverTime = 400

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim PropValuetext As String

PropValuetext = PropBag.ReadProperty("DateTimeValuetext", "")

DTSControls.ControlHeaderMode = PropBag.ReadProperty("HeaderMode", 2)
DTSControls.ControlWeekdayStart = PropBag.ReadProperty("StartingWeekday", 0)
DTSControls.ControlDateTimeFormat = PropBag.ReadProperty("DateTimeFormat", "D")
DTSControls.ControlEmptyBackground = PropBag.ReadProperty("EmptyBackColor", QBColor(7))
DTSControls.ControlDisabledBackground = PropBag.ReadProperty("DisabledBackColor", &HE0E0E0)
DTSControls.ControlNormalBackground = PropBag.ReadProperty("NormalBackColor", QBColor(15))
DTSControls.ControlNormalForeground = PropBag.ReadProperty("NormalForeColor", QBColor(0))
DTSControls.ControlHeaderBackground = PropBag.ReadProperty("HeaderBackColor", QBColor(8))
DTSControls.ControlHeaderForeground = PropBag.ReadProperty("HeaderForeColor", QBColor(15))
DTSControls.ControlSelectedBackground = PropBag.ReadProperty("SelectedBackColor", QBColor(9))
DTSControls.ControlSelectedForeground = PropBag.ReadProperty("SelectedForeColor", QBColor(15))
DTSControls.ControlNowColor = PropBag.ReadProperty("TodayColor", QBColor(12))
DTSControls.ControlReverseScrollFlag = PropBag.ReadProperty("ReverseScroll", False)
DTSControls.ControlFlatMode = PropBag.ReadProperty("FlatMode", False)
DTSControls.ControlNowFlag = PropBag.ReadProperty("TodayFlag", True)

Set DTSFont = PropBag.ReadProperty("DateTimeFont", Ambient.Font)
Set UserControl.Font = DTSFont

Call ResizeCalendar(Me, DTSControls)
Call RecalculateCalendar(DTSControls)

Call LetNewValuetext(Me, DTSControls, PropValuetext)

Set ControlTracking = New MouseControlClass

ControlTracking.hWnd = UserControl.hWnd
ControlTracking.HoverTime = PropBag.ReadProperty("HoverTime", 400)
ControlTracking.overOnly = True

If Ambient.UserMode Then
    StartTrack ControlTracking
End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim PropValuetext As String

PropValuetext = Me.DateTimeValuetext

Call PropBag.WriteProperty("DateTimeValuetext", PropValuetext, "")

Call PropBag.WriteProperty("HeaderMode", DTSControls.ControlHeaderMode, 2)
Call PropBag.WriteProperty("StartingWeekday", DTSControls.ControlWeekdayStart, 0)
Call PropBag.WriteProperty("DateTimeFormat", DTSControls.ControlDateTimeFormat, "D")
Call PropBag.WriteProperty("EmptyBackColor", DTSControls.ControlEmptyBackground, QBColor(7))
Call PropBag.WriteProperty("DisabledBackColor", DTSControls.ControlDisabledBackground, &HE0E0E0)
Call PropBag.WriteProperty("NormalBackColor", DTSControls.ControlNormalBackground, QBColor(15))
Call PropBag.WriteProperty("NormalForeColor", DTSControls.ControlNormalForeground, QBColor(0))
Call PropBag.WriteProperty("HeaderBackColor", DTSControls.ControlHeaderBackground, QBColor(8))
Call PropBag.WriteProperty("HeaderForeColor", DTSControls.ControlHeaderForeground, QBColor(15))
Call PropBag.WriteProperty("SelectedBackColor", DTSControls.ControlSelectedBackground, QBColor(9))
Call PropBag.WriteProperty("SelectedForeColor", DTSControls.ControlSelectedForeground, QBColor(15))
Call PropBag.WriteProperty("TodayColor", DTSControls.ControlNowColor, QBColor(12))
Call PropBag.WriteProperty("ReverseScroll", DTSControls.ControlReverseScrollFlag, False)
Call PropBag.WriteProperty("FlatMode", DTSControls.ControlFlatMode, False)
Call PropBag.WriteProperty("TodayFlag", DTSControls.ControlNowFlag, True)

Call PropBag.WriteProperty("DateTimeFont", DTSFont, Ambient.Font)

Call PropBag.WriteProperty("HoverTime", ControlTracking.HoverTime, 400)

End Sub

Private Sub UserControl_Terminate()

On Error Resume Next
EndTrack ControlTracking
On Error GoTo 0

Set ControlTracking = Nothing

End Sub

Private Sub UserControl_Show()

Call UserControl_Resize

End Sub

Private Sub UserControl_Resize()

Call ResizeDateTimeStampControl(Me, DTSControls)

End Sub

Private Sub UserControl_Initialize()

Call InitializeDTSControls(DTSControls)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

RaiseEvent DateTimeKeyPress(KeyAscii)

End Sub

Private Sub UserControl_DblClick()

Call VerifyMouseAction(Me, DTSControls)

End Sub

Private Sub UserControl_MouseDown(MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

If (MouseButton And 1) = 0 Then Exit Sub

DTSControls.ControlMouseX = MouseX
DTSControls.ControlMouseY = MouseY

Call VerifyMouseAction(Me, DTSControls)

End Sub

Private Sub UserControl_MouseUp(MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

RaiseEvent DateTimeMouseUp(MouseButton)

End Sub

Public Property Get DateTimeFormat() As String

DateTimeFormat = DTSControls.ControlDateTimeFormat

End Property

Public Property Let DateTimeFormat(NewDateTimeFormat As String)

If DTSControls.ControlDateTimeFormat <> NewDateTimeFormat Then

    Call SetLimitMode(DTSControls, 0)
    
    DTSControls.ControlDateTimeFormat = NewDateTimeFormat
    DTSControls.ControlClickMode = "="
    
    Call ResizeCalendar(Me, DTSControls)
    Call RecalculateCalendar(DTSControls)
    
    Call RedrawAllHeaders(Me, DTSControls)
    Call RedrawAllCells(Me, DTSControls)
    Call RedrawBorders(Me, DTSControls)
    
    Me.RaiseDateTimeChange

End If

End Property

Public Property Get StartingWeekday() As DTSWeekdayList

StartingWeekday = DTSControls.ControlWeekdayStart

End Property

Public Property Let StartingWeekday(NewStartingWeekday As DTSWeekdayList)

DTSControls.ControlWeekdayStart = NewStartingWeekday

Call RecalculateCalendar(DTSControls)

Call RedrawDateHeaders(Me, DTSControls)
Call RedrawDateCells(Me, DTSControls)
Call RedrawBorders(Me, DTSControls)

End Property

Public Property Get HeaderMode() As DTSHeaderModeList

HeaderMode = DTSControls.ControlHeaderMode

End Property

Public Property Let HeaderMode(NewMode As DTSHeaderModeList)

If DTSControls.ControlHeaderMode <> NewMode Then

    DTSControls.ControlHeaderMode = NewMode

    Call ResizeCalendar(Me, DTSControls)

    Call RedrawAllHeaders(Me, DTSControls)
    Call RedrawAllCells(Me, DTSControls)
    Call RedrawBorders(Me, DTSControls)

End If

End Property

Public Property Get HeaderBackColor() As OLE_COLOR

HeaderBackColor = DTSControls.ControlHeaderBackground

End Property

Public Property Let HeaderBackColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlHeaderBackground, NewColor, 1)

End Property

Public Property Get HeaderForeColor() As OLE_COLOR

HeaderForeColor = DTSControls.ControlHeaderForeground

End Property

Public Property Let HeaderForeColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlHeaderForeground, NewColor, 1)

End Property

Public Property Get EmptyBackColor() As OLE_COLOR

EmptyBackColor = DTSControls.ControlEmptyBackground

End Property

Public Property Let EmptyBackColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlEmptyBackground, NewColor, 2)

End Property

Public Property Get DisabledBackColor() As OLE_COLOR

DisabledBackColor = DTSControls.ControlDisabledBackground

End Property

Public Property Let DisabledBackColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlDisabledBackground, NewColor, 2)

End Property

Public Property Get NormalBackColor() As OLE_COLOR

NormalBackColor = DTSControls.ControlNormalBackground

End Property

Public Property Let NormalBackColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlNormalBackground, NewColor, 30)

End Property

Public Property Get NormalForeColor() As OLE_COLOR

NormalForeColor = DTSControls.ControlNormalForeground

End Property

Public Property Let NormalForeColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlNormalForeground, NewColor, 30)

End Property

Public Property Get SelectedBackColor() As OLE_COLOR

SelectedBackColor = DTSControls.ControlSelectedBackground

End Property

Public Property Let SelectedBackColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlSelectedBackground, NewColor, 30)

End Property

Public Property Get SelectedForeColor() As OLE_COLOR

SelectedForeColor = DTSControls.ControlSelectedForeground

End Property

Public Property Let SelectedForeColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlSelectedForeground, NewColor, 30)

End Property

Public Property Get TodayColor() As OLE_COLOR

TodayColor = DTSControls.ControlNowColor

End Property

Public Property Let TodayColor(NewColor As OLE_COLOR)

Call SwapControlColor(Me, DTSControls, DTSControls.ControlNowColor, NewColor, 2)

End Property

Public Property Get FlatMode() As Boolean

FlatMode = DTSControls.ControlFlatMode

End Property

Public Property Let FlatMode(NewFlag As Boolean)

If DTSControls.ControlFlatMode <> NewFlag Then

    DTSControls.ControlFlatMode = NewFlag

    Call ResizeDateTimeStampControl(Me, DTSControls)

End If

End Property

Public Property Get TodayFlag() As Boolean

TodayFlag = DTSControls.ControlNowFlag

End Property

Public Property Let TodayFlag(NewFlag As Boolean)

DTSControls.ControlNowFlag = NewFlag

Call RedrawDateCells(Me, DTSControls)
Call RedrawBorders(Me, DTSControls)

End Property

Public Property Get ReverseScroll() As Boolean

ReverseScroll = DTSControls.ControlReverseScrollFlag

End Property

Public Property Let ReverseScroll(NewFlag As Boolean)

DTSControls.ControlReverseScrollFlag = NewFlag

End Property

Public Property Get DateTimeFont() As Font

Set DateTimeFont = UserControl.Font

End Property

Public Property Set DateTimeFont(ByVal NewFont As Font)

Set UserControl.Font = NewFont
Set DTSFont = NewFont

Call RedrawAllHeaders(Me, DTSControls)
Call RedrawAllCells(Me, DTSControls)
Call RedrawBorders(Me, DTSControls)

End Property

Public Property Get DateTimeValuetext() As String

DateTimeValuetext = GetDateTimeValuetext(DTSControls)

End Property

Public Property Let DateTimeValuetext(NewValuetext As String)

Call LetNewValuetext(Me, DTSControls, NewValuetext)

End Property

Public Property Get DateTimeMinimum() As String

DateTimeMinimum = GetDateTimeMinimum(DTSControls)

End Property

Public Property Let DateTimeMinimum(NewMinimum As String)

Call LetNewMinimum(Me, DTSControls, NewMinimum)

End Property

Public Property Get DateTimeMaximum() As String

DateTimeMaximum = GetDateTimeMaximum(DTSControls)

End Property

Public Property Let DateTimeMaximum(NewMaximum As String)

Call LetNewMaximum(Me, DTSControls, NewMaximum)

End Property

Public Property Get HoverTime() As Long

HoverTime = ControlTracking.HoverTime

End Property

Public Property Let HoverTime(newHoverTime As Long)

ControlTracking.HoverTime = newHoverTime

End Property

Public Sub SetToday()

Call MouseSwapToday(Me, DTSControls)

End Sub

Public Sub SetStandardLabels()

Call SetSpecialLabels("Mon,Tue,Wed,Thu,Fri,Sat,Sun", "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", "Month", "Year", "Hou", "Min", "Sec")

End Sub

Public Sub SetSpecialLabels(DayLabels As String, MonthLabels As String, MonthHeader As String, YearHeader As String, HourHeader As String, MinuteHeader As String, SecondHeader As String)

Call SplitLabelFields(DayLabels, DTSControls.ControlWeekdayNames, 6)
Call SplitLabelFields(MonthLabels, DTSControls.ControlMonthNames, 11)

DTSControls.ControlHeaderMonth = MonthHeader
DTSControls.ControlHeaderYear = YearHeader
DTSControls.ControlHeaderTime(0) = HourHeader
DTSControls.ControlHeaderTime(1) = MinuteHeader
DTSControls.ControlHeaderTime(2) = SecondHeader

Call RedrawAllHeaders(Me, DTSControls)
Call RedrawBorders(Me, DTSControls)

End Sub

Friend Sub DrawLine(DrawX1 As Long, DrawY1 As Long, DrawX2 As Long, DrawY2 As Long, DrawColor As OLE_COLOR, Optional DrawMode As String = "")

Select Case DrawMode
    Case "B":  UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor, B
    Case "BF": UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor, BF
    Case Else: UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor
End Select

End Sub

Friend Sub DrawPrint(DrawX As Long, DrawY As Long, DrawColor As OLE_COLOR, DrawText As String)

UserControl.CurrentX = DrawX
UserControl.CurrentY = DrawY
UserControl.ForeColor = DrawColor
UserControl.Print DrawText

End Sub

Friend Function TextWidth(TextBuffer As String) As Long

TextWidth = UserControl.TextWidth(TextBuffer)

End Function

Friend Function TextHeight(TextBuffer As String) As Long

TextHeight = UserControl.TextHeight(TextBuffer)

End Function

Friend Function ScaleWidth() As Long

ScaleWidth = UserControl.ScaleWidth

End Function

Friend Function ScaleHeight() As Long

ScaleHeight = UserControl.ScaleHeight

End Function

Friend Sub RaiseDateTimeChange()

RaiseEvent DateTimeChange

End Sub

Public Sub ScrollDateTime(DayChange As Integer, MonthChange As Integer, YearChange As Integer, HourChange As Integer, MinuteChange As Integer, SecondChange As Integer, CronoChange As Integer)

Call DoScrollDateTime(Me, DTSControls, DayChange, MonthChange, YearChange, HourChange, MinuteChange, SecondChange, CronoChange)

End Sub

Private Sub ControlTracking_HoverChange(IsHovered As Boolean)

RaiseEvent HoverChange(IsHovered)

End Sub

Private Sub ControlTracking_MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

RaiseEvent MouseMove(MouseState, MouseX, MouseY)

End Sub

Private Sub ControlTracking_WheelScroll(MouseKeys As Long, MouseRotation As Long, MousePosX As Long, MousePosY As Long)

Dim MouseOffsetX As Long
Dim MouseOffsetY As Long

Call GetUserControlOffset(UserControl.hWnd, MouseOffsetX, MouseOffsetY)

MouseOffsetX = (MousePosX - MouseOffsetX) * 15
MouseOffsetY = (MousePosY - MouseOffsetY) * 15

Call DoScrollMouseWheel(Me, DTSControls, MouseKeys, MouseRotation, MouseOffsetX, MouseOffsetY)

End Sub
