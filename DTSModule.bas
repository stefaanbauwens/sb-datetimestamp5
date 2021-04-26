Attribute VB_Name = "DTSModule"

Public Const DTSCellBordercolor = &H646464
Public Const DTSTimingSeparator = "."
Public Const DTSSecondSeparator = ":"
Public Const DTSHourSeparator = ":"
Public Const DTSDateSeparator = "/"

Public Type DTSValuetext
    ValuetextDay As Integer
    ValuetextYear As Integer
    ValuetextMonth As Integer
    ValuetextTime(0 To 2) As Integer
    ValuetextCrono As Long
End Type
Public Type DTSControls
    ControlMouseX As Single
    ControlMouseY As Single
    ControlLimitMode As Integer
    ControlMinimumValuetext As DTSValuetext
    ControlMaximumValuetext As DTSValuetext
    ControlCurrentValuetext As DTSValuetext
    ControlEmptyBackground As OLE_COLOR
    ControlDisabledBackground As OLE_COLOR
    ControlNormalBackground As OLE_COLOR
    ControlNormalForeground As OLE_COLOR
    ControlSelectedBackground As OLE_COLOR
    ControlSelectedForeground As OLE_COLOR
    ControlNowColor As OLE_COLOR
    ControlNowFlag As Boolean
    ControlNowDay As Integer
    ControlNowYear As Integer
    ControlNowMonth As Integer
    ControlDayOffset As Integer
    ControlDayCellwidth As Long
    ControlDayCellheight As Long
    ControlDayCellnumber As Integer
    ControlDayCelltext(0 To 41) As String * 3
    ControlDayCelltype(0 To 41) As Integer
    ControlWeekdayNames(0 To 6) As String
    ControlWeekdayStart As Integer
    ControlCronoDecimals As Integer
    ControlCronoCellheight As Long
    ControlCronoCellwidth As Long
    ControlCronoCellleft As Long
    ControlHeaderMonth As String
    ControlHeaderYear As String
    ControlHeaderCellheight As Long
    ControlHeaderMode As Integer
    ControlHeaderTime(0 To 2) As String
    ControlHeaderBackground As OLE_COLOR
    ControlHeaderForeground As OLE_COLOR
    ControlMonthNames(0 To 11) As String
    ControlMonthCellnumber As Integer
    ControlMonthCellheight As Long
    ControlMonthCellwidth As Long
    ControlMonthCellleft As Long
    ControlYearCellwidth As Long
    ControlYearCellleft As Long
    ControlYearCellheight As Long
    ControlYearOffset As Long
    ControlYearStart As Integer
    ControlYearEnd As Integer
    ControlDateTimeFormat As String
    ControlInitialized As Boolean
    ControlTimeCellheight As Long
    ControlTimeCellwidth As Long
    ControlTimeCellleft As Long
    ControlClickMode As String
    ControlReverseScrollFlag As Boolean
    ControlFlatMode As Boolean
    ControlLastDay As Long
End Type

Public Sub InitializeDTSControls(InitializeControls As DTSControls)

InitializeControls.ControlInitialized = False

InitializeControls.ControlWeekdayStart = 0

InitializeControls.ControlEmptyBackground = QBColor(7)
InitializeControls.ControlDisabledBackground = &HE0E0E0
InitializeControls.ControlNormalBackground = QBColor(15)
InitializeControls.ControlNormalForeground = QBColor(0)
InitializeControls.ControlHeaderBackground = QBColor(8)
InitializeControls.ControlHeaderForeground = QBColor(15)
InitializeControls.ControlSelectedBackground = QBColor(9)
InitializeControls.ControlSelectedForeground = QBColor(15)
InitializeControls.ControlNowColor = QBColor(12)

InitializeControls.ControlNowFlag = True
InitializeControls.ControlNowDay = Day(Now)
InitializeControls.ControlNowMonth = Month(Now)
InitializeControls.ControlNowYear = Year(Now)

InitializeControls.ControlCurrentValuetext.ValuetextDay = InitializeControls.ControlNowDay
InitializeControls.ControlCurrentValuetext.ValuetextMonth = InitializeControls.ControlNowMonth
InitializeControls.ControlCurrentValuetext.ValuetextYear = InitializeControls.ControlNowYear
InitializeControls.ControlCurrentValuetext.ValuetextTime(0) = Hour(Now)
InitializeControls.ControlCurrentValuetext.ValuetextTime(1) = Minute(Now)
InitializeControls.ControlCurrentValuetext.ValuetextTime(2) = Second(Now)
InitializeControls.ControlCurrentValuetext.ValuetextCrono = 0

Call SetLimitMode(InitializeControls, 0)

InitializeControls.ControlDayCellnumber = -1
InitializeControls.ControlMonthCellnumber = -1
InitializeControls.ControlYearOffset = 195

InitializeControls.ControlClickMode = "="

InitializeControls.ControlHeaderMode = 2

InitializeControls.ControlFlatMode = False

End Sub

Public Sub ResizeDateTimeStampControl(ResizeMe As DateTimeStamp, ResizeControls As DTSControls)

If ResizeControls.ControlInitialized = False Then

    ResizeControls.ControlInitialized = True

    ResizeMe.SetStandardLabels

Else

    Call ResizeCalendar(ResizeMe, ResizeControls)
    Call RecalculateCalendar(ResizeControls)
    
    Call RedrawAllHeaders(ResizeMe, ResizeControls)

End If

Call RedrawAllCells(ResizeMe, ResizeControls)
Call RedrawBorders(ResizeMe, ResizeControls)

End Sub

Public Function GetDateTimeValuetext(GetControls As DTSControls) As String

GetDateTimeValuetext = DoGetDateTimeControldata(GetControls, GetControls.ControlCurrentValuetext)

End Function

Public Sub LetNewValuetext(NewMe As DateTimeStamp, NewControls As DTSControls, NewValuetext As String)

Call DoLetNewControldata(NewControls, NewControls.ControlCurrentValuetext, NewValuetext, 0)
Call VerifyValuedateLimits(NewControls, NewControls.ControlCurrentValuetext)

NewControls.ControlClickMode = "="

If NewControls.ControlDayCellwidth > 0 Or NewControls.ControlMonthCellwidth > 0 Then

    Call VerifyMonthChange(NewMe, NewControls)

Else

    Call RedrawYearCells(NewMe, NewControls)
    Call RedrawTimeSecondsCells(NewMe, NewControls)
    Call RedrawCronoCells(NewMe, NewControls)
    Call RedrawBorders(NewMe, NewControls)

    NewMe.RaiseDateTimeChange

End If

End Sub

Public Function GetDateTimeMinimum(GetControls As DTSControls) As String

If (GetControls.ControlLimitMode And 1) = 1 Then
    GetDateTimeMinimum = DoGetDateTimeControldata(GetControls, GetControls.ControlMinimumValuetext)
Else
    GetDateTimeMinimum = ""
End If

End Function

Public Sub LetNewMinimum(NewMe As DateTimeStamp, NewControls As DTSControls, NewMinimum As String)

If NewMinimum = "" Then
    
    Call SetLimitMode(NewControls, (NewControls.ControlLimitMode And 2))

Else
    Call DoLetNewControldata(NewControls, NewControls.ControlMinimumValuetext, NewMinimum, 0)

    Call SetLimitMode(NewControls, (NewControls.ControlLimitMode Or 1))

    If (NewControls.ControlLimitMode And 2) = 2 Then
        If ValidValuedateLimit(NewControls, 88, 8888, NewControls.ControlMaximumValuetext, NewControls.ControlMinimumValuetext, 1) = False Then Call SetLimitMode(NewControls, 1)
    End If

End If

Call ResizeAndRedraw(NewMe, NewControls)

End Sub

Public Function GetDateTimeMaximum(GetControls As DTSControls) As String

If (GetControls.ControlLimitMode And 2) = 2 Then
    GetDateTimeMaximum = DoGetDateTimeControldata(GetControls, GetControls.ControlMaximumValuetext)
Else
    GetDateTimeMaximum = ""
End If

End Function

Public Sub LetNewMaximum(NewMe As DateTimeStamp, NewControls As DTSControls, NewMaximum As String)

If NewMaximum = "" Then
    
    Call SetLimitMode(NewControls, (NewControls.ControlLimitMode And 1))

Else

    Call DoLetNewControldata(NewControls, NewControls.ControlMaximumValuetext, NewMaximum, 9999)

    Call SetLimitMode(NewControls, (NewControls.ControlLimitMode Or 2))

    If (NewControls.ControlLimitMode And 1) = 1 Then
        If ValidValuedateLimit(NewControls, 88, 8888, NewControls.ControlMinimumValuetext, NewControls.ControlMaximumValuetext, 2) = False Then Call SetLimitMode(NewControls, 2)
    End If

End If

Call ResizeAndRedraw(NewMe, NewControls)

End Sub

Public Sub SetLimitMode(SetControls As DTSControls, SetNewMode As Integer)

Dim SetRange As Long

SetControls.ControlLimitMode = SetNewMode

If (SetNewMode And 1) = 0 Then
    SetControls.ControlMinimumValuetext.ValuetextDay = 1
    SetControls.ControlMinimumValuetext.ValuetextMonth = 1
    SetControls.ControlMinimumValuetext.ValuetextYear = 1000
    SetControls.ControlMinimumValuetext.ValuetextTime(0) = 0
    SetControls.ControlMinimumValuetext.ValuetextTime(1) = 0
    SetControls.ControlMinimumValuetext.ValuetextTime(2) = 0
    SetControls.ControlMinimumValuetext.ValuetextCrono = 0
End If

If (SetNewMode And 2) = 0 Then
    SetControls.ControlMaximumValuetext.ValuetextDay = 31
    SetControls.ControlMaximumValuetext.ValuetextMonth = 12
    SetControls.ControlMaximumValuetext.ValuetextYear = 2999
    SetControls.ControlMaximumValuetext.ValuetextTime(0) = 23
    SetControls.ControlMaximumValuetext.ValuetextTime(1) = 59
    SetControls.ControlMaximumValuetext.ValuetextTime(2) = 59
    SetControls.ControlMaximumValuetext.ValuetextCrono = 999999
End If

SetRange = SetControls.ControlMaximumValuetext.ValuetextYear - SetControls.ControlMinimumValuetext.ValuetextYear + 1

If SetNewMode <> 3 Or SetRange < 1 Or SetRange > 12 Then
    SetControls.ControlYearStart = 0
    SetControls.ControlYearEnd = 0
Else
    SetControls.ControlYearStart = SetControls.ControlMinimumValuetext.ValuetextYear
    SetControls.ControlYearEnd = SetControls.ControlMaximumValuetext.ValuetextYear
End If

End Sub

Public Sub RecalculateCalendar(PrepareControls As DTSControls)

Dim PrepareDay As Variant
Dim PrepareLast As Variant
Dim PrepareFirst As Variant
Dim PrepareWeekday As Integer
Dim PrepareStatus As Integer
Dim PrepareIndex As Integer
Dim PrepareCell As Integer

If PrepareControls.ControlYearCellleft <= 0 Then Exit Sub
If PrepareControls.ControlDayCellwidth <= 0 Then Exit Sub

PrepareFirst = DateSerial(PrepareControls.ControlCurrentValuetext.ValuetextYear, PrepareControls.ControlCurrentValuetext.ValuetextMonth, 1)
PrepareLast = DateSerial(PrepareControls.ControlCurrentValuetext.ValuetextYear, PrepareControls.ControlCurrentValuetext.ValuetextMonth + 1, 1)

PrepareWeekday = Weekday(PrepareFirst, vbMonday) - PrepareControls.ControlWeekdayStart

If PrepareWeekday < 1 Then PrepareWeekday = PrepareWeekday + 7

PrepareCell = 0
PrepareIndex = 1
PrepareStatus = 0

While PrepareIndex < 43

    If PrepareStatus = 0 Then
        If PrepareWeekday = PrepareIndex Then
            PrepareControls.ControlDayOffset = PrepareIndex + 6 - 8
            PrepareDay = PrepareFirst
            PrepareStatus = 1
        Else
            PrepareControls.ControlDayCelltext(PrepareCell) = ""
            PrepareControls.ControlDayCelltype(PrepareCell) = TypeEmpty
        End If
    End If

    If PrepareStatus = 1 Then
        If PrepareDay >= PrepareLast Then
            PrepareStatus = 2
        Else
            PrepareControls.ControlDayCelltext(PrepareCell) = DatePartFormat(Day(PrepareDay))
            If Val(PrepareControls.ControlDayCelltext(PrepareCell)) = PrepareControls.ControlCurrentValuetext.ValuetextDay Then
                PrepareControls.ControlDayCelltype(PrepareCell) = TypeSelected
            Else
                PrepareControls.ControlDayCelltype(PrepareCell) = TypeNormal
            End If
            PrepareDay = DateAdd("d", 1, PrepareDay)
        End If
    End If

    If PrepareStatus = 2 Then
        PrepareControls.ControlDayCelltext(PrepareCell) = ""
        PrepareControls.ControlDayCelltype(PrepareCell) = TypeEmpty
    End If

    PrepareIndex = PrepareIndex + 1
    PrepareCell = PrepareCell + 1

Wend

PrepareControls.ControlLastDay = Day(DateAdd("d", -1, PrepareLast))

End Sub

Public Sub ResizeCalendar(ResizeMe As DateTimeStamp, ResizeControls As DTSControls)

Dim ResizeDay As Boolean
Dim ResizeYear As Boolean
Dim ResizeTime As Boolean
Dim ResizeMonth As Boolean
Dim ResizeFormat As String
Dim ResizeSeconds As Boolean
Dim ResizeFactor As Double
Dim ResizeSize As Double

ResizeFormat = ""
ResizeDay = False
ResizeYear = False
ResizeTime = False
ResizeMonth = False
ResizeSeconds = False
ResizeFactor = 0

ResizeControls.ControlCronoDecimals = 0

If Len(ResizeControls.ControlDateTimeFormat) > 1 Then ResizeFormat = Left$(ResizeControls.ControlDateTimeFormat, Len(ResizeControls.ControlDateTimeFormat) - 1)

Select Case Right(ResizeControls.ControlDateTimeFormat, 1)
    Case "Y"
        ResizeYear = True
    Case "M"
        ResizeYear = True
        ResizeMonth = True
    Case "H"
        ResizeTime = True
        If ResizeFormat <> "R" Then
            ResizeSeconds = True
            If Val(ResizeFormat) > 0 Then ResizeControls.ControlCronoDecimals = Val(ResizeFormat)
        End If
    Case "S", "T"
        ResizeDay = True
        ResizeYear = True
        ResizeTime = True
        ResizeMonth = True
        If ResizeFormat <> "R" Then
            ResizeSeconds = True
            If Val(ResizeFormat) > 0 Then ResizeControls.ControlCronoDecimals = Val(ResizeFormat)
        End If
    Case Else
        ResizeDay = True
        If ResizeFormat <> "F" Then
            ResizeMonth = True
            ResizeYear = True
        End If
End Select

If ResizeControls.ControlCronoDecimals < 1 Or ResizeControls.ControlCronoDecimals > 6 Then ResizeControls.ControlCronoDecimals = 0

If ResizeDay = True Then ResizeFactor = ResizeFactor + 350
If ResizeMonth = True Then ResizeFactor = ResizeFactor + 195
If ResizeYear = True Then ResizeFactor = ResizeFactor + 150 - GetYearMode(ResizeControls, 0) * 75
If ResizeTime = True Then ResizeFactor = ResizeFactor + 200
If ResizeSeconds = True Then ResizeFactor = ResizeFactor + 100
If ResizeControls.ControlCronoDecimals > 0 Then ResizeFactor = ResizeFactor + Choose(ResizeControls.ControlCronoDecimals, 78, 100, 124, 150, 178, 208)

ResizeSize = ResizeMe.ScaleWidth - 15

If ResizeDay = True And ResizeMonth = True Then ResizeSize = ResizeSize - 15
If ResizeMonth = True And ResizeYear = True Then ResizeSize = ResizeSize - 15
If ResizeYear = True And ResizeTime = True Then ResizeSize = ResizeSize - 15
If ResizeTime = True Then ResizeSize = ResizeSize - 15
If ResizeTime = True And ResizeSeconds = True Then ResizeSize = ResizeSize - 15
If ResizeSeconds = True And ResizeControls.ControlCronoDecimals > 0 Then ResizeSize = ResizeSize - 15

ResizeFactor = ResizeSize / ResizeFactor
ResizeSize = ResizeMe.ScaleWidth

ResizeControls.ControlYearCellheight = ResizeMe.ScaleHeight - 15
ResizeControls.ControlHeaderCellheight = ResizeControls.ControlYearCellheight / 7

If ResizeControls.ControlHeaderMode = 0 Then ResizeControls.ControlHeaderCellheight = 0
If ResizeControls.ControlHeaderMode = 1 Then ResizeControls.ControlHeaderCellheight = ResizeControls.ControlHeaderCellheight * 0.75

If ResizeControls.ControlHeaderCellheight < 45 Then ResizeControls.ControlHeaderCellheight = 0

ResizeControls.ControlYearCellheight = ResizeControls.ControlYearCellheight - ResizeControls.ControlHeaderCellheight
ResizeControls.ControlDayCellheight = ResizeControls.ControlYearCellheight / 6
ResizeControls.ControlMonthCellheight = ResizeControls.ControlYearCellheight / 4
ResizeControls.ControlCronoCellheight = ResizeControls.ControlDayCellheight
ResizeControls.ControlYearCellheight = ResizeControls.ControlDayCellheight
ResizeControls.ControlTimeCellheight = ResizeControls.ControlDayCellheight

If ResizeDay = True Then
    ResizeControls.ControlDayCellwidth = ResizeFactor * 50
    ResizeControls.ControlMonthCellleft = ResizeControls.ControlDayCellwidth * 7 + 15
Else
    ResizeControls.ControlDayCellwidth = 0
    ResizeControls.ControlMonthCellleft = 0
End If

If ResizeMonth = True Then
    ResizeControls.ControlMonthCellwidth = ResizeFactor * 65
    ResizeControls.ControlYearCellleft = ResizeControls.ControlMonthCellleft + ResizeControls.ControlMonthCellwidth * 3 + 15
Else
    ResizeControls.ControlMonthCellwidth = 0
    ResizeControls.ControlYearCellleft = ResizeControls.ControlMonthCellleft
End If

If ResizeControls.ControlCronoDecimals > 0 Then
    ResizeControls.ControlCronoCellwidth = ResizeFactor * Choose(ResizeControls.ControlCronoDecimals, 39, 50, 62, 75, 89, 104)
    ResizeControls.ControlCronoCellleft = ResizeSize - ResizeControls.ControlCronoCellwidth * 2 - 15
Else
    ResizeControls.ControlCronoCellleft = ResizeSize
    ResizeControls.ControlCronoCellwidth = 0
End If

If ResizeYear = True Then

    If ResizeSeconds = True Then
        ResizeControls.ControlTimeCellwidth = ResizeFactor * 50
        ResizeControls.ControlTimeCellleft = ResizeControls.ControlCronoCellleft - ResizeControls.ControlTimeCellwidth * 6 - 45
    ElseIf ResizeTime = True Then
        ResizeControls.ControlTimeCellwidth = ResizeFactor * 50
        ResizeControls.ControlTimeCellleft = ResizeControls.ControlCronoCellleft - ResizeControls.ControlTimeCellwidth * 4 - 30
    Else
        ResizeControls.ControlTimeCellwidth = 0
        ResizeControls.ControlTimeCellleft = ResizeControls.ControlCronoCellleft
    End If

    ResizeControls.ControlYearCellwidth = (ResizeControls.ControlTimeCellleft - ResizeControls.ControlYearCellleft - 15) / 2

Else
    
    ResizeControls.ControlYearCellwidth = 0
    ResizeControls.ControlTimeCellleft = ResizeControls.ControlYearCellleft

    If ResizeDay = True Then
        ResizeControls.ControlTimeCellwidth = 0
    ElseIf ResizeSeconds = True Then
        ResizeFactor = ResizeControls.ControlCronoCellleft - ResizeControls.ControlYearCellleft - 45
        ResizeControls.ControlTimeCellwidth = ResizeFactor / 6
    Else
        ResizeFactor = ResizeControls.ControlCronoCellleft - ResizeControls.ControlYearCellleft - 30
        ResizeControls.ControlTimeCellwidth = ResizeFactor / 4
    End If
    
End If

End Sub

Public Sub SplitLabelFields(FieldLabels As String, FieldArray() As String, FieldSize As Integer)

Dim FieldIndex As Integer
Dim FieldPosition As Integer
Dim FieldBuffer As String

FieldBuffer = FieldLabels

FieldIndex = 0
While FieldIndex < FieldSize

    FieldPosition = InStr(FieldLabels + ",", ",")
    FieldArray(FieldIndex) = Left$(FieldBuffer, FieldPosition - 1)
    FieldBuffer = Mid$(FieldBuffer, FieldPosition + 1)

    FieldIndex = FieldIndex + 1
Wend

FieldArray(FieldIndex) = FieldBuffer

End Sub

Public Sub SwapControlColor(NewMe As DateTimeStamp, NewControls As DTSControls, OldColor As OLE_COLOR, NewColor As OLE_COLOR, NewMode As Integer)

If OldColor = NewColor Then Exit Sub

OldColor = NewColor

If (NewMode And 1) > 0 Then Call RedrawAllHeaders(NewMe, NewControls)

If (NewMode And 2) > 0 Then Call RedrawDateCells(NewMe, NewControls)
If (NewMode And 4) > 0 Then Call RedrawYearCells(NewMe, NewControls)
If (NewMode And 8) > 0 Then Call RedrawTimeSecondsCells(NewMe, NewControls)
If (NewMode And 16) > 0 Then Call RedrawCronoCells(NewMe, NewControls)

If (NewMode And 31) > 0 Then Call RedrawBorders(NewMe, NewControls)

End Sub

Public Sub MouseSwapToday(MouseMe As DateTimeStamp, MouseControls As DTSControls)

Dim SwapMode As Integer
Dim SwapDatevalue As DTSValuetext

SwapMode = 0

SwapDatevalue = MouseControls.ControlCurrentValuetext

SwapDatevalue.ValuetextDay = MouseControls.ControlNowDay
SwapDatevalue.ValuetextMonth = MouseControls.ControlNowMonth
SwapDatevalue.ValuetextYear = MouseControls.ControlNowYear

Call VerifyValuedateLimits(MouseControls, SwapDatevalue)

If MouseControls.ControlCurrentValuetext.ValuetextDay <> SwapDatevalue.ValuetextDay Then SwapMode = SwapMode + 1
If MouseControls.ControlCurrentValuetext.ValuetextMonth <> SwapDatevalue.ValuetextMonth Then SwapMode = SwapMode + 2
If MouseControls.ControlCurrentValuetext.ValuetextYear <> SwapDatevalue.ValuetextYear Then SwapMode = SwapMode + 4

MouseControls.ControlCurrentValuetext = SwapDatevalue

If SwapMode > 0 Then Call RecalculateCalendar(MouseControls)

Call SetYearOffset(MouseControls)

If (SwapMode And 3) > 0 Then Call RedrawDateCells(MouseMe, MouseControls)

Call RedrawYearCells(MouseMe, MouseControls)

If SwapMode = 0 Then Exit Sub

Call RedrawBorders(MouseMe, MouseControls)

MouseMe.RaiseDateTimeChange

End Sub

Public Sub VerifyMouseAction(VerifyMe As DateTimeStamp, VerifyControls As DTSControls)

If VerifyControls.ControlMouseY <= VerifyControls.ControlHeaderCellheight Then Exit Sub

If VerifyControls.ControlDayCellwidth > 0 Then
    If VerifyControls.ControlMouseX < VerifyControls.ControlMonthCellleft Then Call MouseDayChange(VerifyMe, VerifyControls, VerifyControls.ControlMouseX, VerifyControls.ControlMouseY)
End If
    
If VerifyControls.ControlMonthCellwidth > 0 Then
    If VerifyControls.ControlMouseX > VerifyControls.ControlMonthCellleft And VerifyControls.ControlMouseX < VerifyControls.ControlYearCellleft Then Call MouseMonthChange(VerifyMe, VerifyControls, VerifyControls.ControlMouseX, VerifyControls.ControlMouseY)
End If

If VerifyControls.ControlYearCellwidth > 0 Then
    If VerifyControls.ControlMouseX > VerifyControls.ControlYearCellleft And VerifyControls.ControlMouseX < VerifyControls.ControlTimeCellleft Then Call MouseGridChange(VerifyMe, VerifyControls, VerifyControls.ControlMouseX, VerifyControls.ControlMouseY, VerifyControls.ControlYearCellleft, VerifyControls.ControlYearCellwidth, VerifyControls.ControlYearCellheight, "Y")
End If

If VerifyControls.ControlTimeCellwidth > 0 Then
    If VerifyControls.ControlMouseX > VerifyControls.ControlTimeCellleft And VerifyControls.ControlMouseX < VerifyControls.ControlCronoCellleft Then Call MouseTimeChange(VerifyMe, VerifyControls, VerifyControls.ControlMouseX, VerifyControls.ControlMouseY)
End If

If VerifyControls.ControlCronoCellwidth > 0 Then
    If VerifyControls.ControlMouseX > VerifyControls.ControlCronoCellleft Then Call MouseGridChange(VerifyMe, VerifyControls, VerifyControls.ControlMouseX, VerifyControls.ControlMouseY, VerifyControls.ControlCronoCellleft, VerifyControls.ControlCronoCellwidth, VerifyControls.ControlCronoCellheight, "E")
End If

End Sub

Public Sub DoScrollDateTime(ScrollMe As DateTimeStamp, ScrollControls As DTSControls, DayChange As Integer, MonthChange As Integer, YearChange As Integer, HourChange As Integer, MinuteChange As Integer, SecondChange As Integer, CronoChange As Integer)

Dim ScrollText As String
Dim ScrollChanged As Integer
Dim ScrollValuetext As DTSValuetext
Dim ScrollMinimum As Date
Dim ScrollMaximum As Date
Dim ScrollSerial As Date
Dim ScrollLenght As Long
Dim ScrollValue As Long

ScrollValuetext = ScrollControls.ControlCurrentValuetext

ScrollSerial = DateSerial(ScrollValuetext.ValuetextYear, ScrollValuetext.ValuetextMonth, ScrollValuetext.ValuetextDay)
ScrollMinimum = DateSerial(ScrollValuetext.ValuetextYear, ScrollValuetext.ValuetextMonth, 1)
ScrollMaximum = DateSerial(ScrollValuetext.ValuetextYear, ScrollValuetext.ValuetextMonth + 1, 0)
ScrollValue = Val("1" + String$(ScrollControls.ControlCronoDecimals, 48))

If ScrollControls.ControlCronoCellwidth > 0 Then

    If CronoChange <> 0 Then ScrollValuetext.ValuetextCrono = ScrollValuetext.ValuetextCrono + CronoChange
    
    If ScrollValuetext.ValuetextCrono < 0 Then

        ScrollLenght = -ScrollValuetext.ValuetextCrono
        ScrollLenght = (ScrollLenght - 1) \ ScrollValue
        ScrollLenght = ScrollLenght + 1

        ScrollValuetext.ValuetextCrono = ScrollValuetext.ValuetextCrono + ScrollLenght * ScrollValue
        
        If ScrollControls.ControlTimeCellwidth > 0 Then ScrollValuetext.ValuetextTime(2) = ScrollValuetext.ValuetextTime(2) - ScrollLenght

    End If

    If ScrollValuetext.ValuetextCrono >= ScrollValue Then

        ScrollLenght = ScrollValuetext.ValuetextCrono \ ScrollValue
    
        ScrollValuetext.ValuetextCrono = ScrollValuetext.ValuetextCrono - ScrollLenght * ScrollValue
        
        If ScrollControls.ControlTimeCellwidth > 0 Then ScrollValuetext.ValuetextTime(2) = ScrollValuetext.ValuetextTime(2) + ScrollLenght
    
    End If

End If

If ScrollControls.ControlTimeCellwidth > 0 Then

    If SecondChange <> 0 Then ScrollValuetext.ValuetextTime(2) = ScrollValuetext.ValuetextTime(2) + SecondChange

    If ScrollValuetext.ValuetextTime(2) < 0 Then
    
        ScrollLenght = -ScrollValuetext.ValuetextTime(2)
        ScrollLenght = (ScrollLenght - 1) \ 60
        ScrollLenght = ScrollLenght + 1
    
        ScrollValuetext.ValuetextTime(2) = ScrollValuetext.ValuetextTime(2) + ScrollLenght * 60
        ScrollValuetext.ValuetextTime(1) = ScrollValuetext.ValuetextTime(1) - ScrollLenght

    End If

    If ScrollValuetext.ValuetextTime(2) > 59 Then

        ScrollLenght = ScrollValuetext.ValuetextTime(2) \ 60

        ScrollValuetext.ValuetextTime(2) = ScrollValuetext.ValuetextTime(2) - ScrollLenght * 60
        ScrollValuetext.ValuetextTime(1) = ScrollValuetext.ValuetextTime(1) + ScrollLenght

    End If

    If MinuteChange <> 0 Then ScrollValuetext.ValuetextTime(1) = ScrollValuetext.ValuetextTime(1) + MinuteChange

    If ScrollValuetext.ValuetextTime(1) < 0 Then

        ScrollLenght = -ScrollValuetext.ValuetextTime(1)
        ScrollLenght = (ScrollLenght - 1) \ 60
        ScrollLenght = ScrollLenght + 1
    
        ScrollValuetext.ValuetextTime(1) = ScrollValuetext.ValuetextTime(1) + ScrollLenght * 60
        ScrollValuetext.ValuetextTime(0) = ScrollValuetext.ValuetextTime(0) - ScrollLenght

    End If

    If ScrollValuetext.ValuetextTime(1) > 59 Then

        ScrollLenght = ScrollValuetext.ValuetextTime(1) \ 60

        ScrollValuetext.ValuetextTime(1) = ScrollValuetext.ValuetextTime(1) - ScrollLenght * 60
        ScrollValuetext.ValuetextTime(0) = ScrollValuetext.ValuetextTime(0) + ScrollLenght

    End If

    If HourChange <> 0 Then ScrollValuetext.ValuetextTime(0) = ScrollValuetext.ValuetextTime(0) + HourChange

    If ScrollValuetext.ValuetextTime(0) < 0 Then
    
        ScrollLenght = -ScrollValuetext.ValuetextTime(0)
        ScrollLenght = (ScrollLenght - 1) \ 24
        ScrollLenght = ScrollLenght + 1
    
        ScrollValuetext.ValuetextTime(0) = ScrollValuetext.ValuetextTime(0) + ScrollLenght * 24
        
        If ScrollControls.ControlDayCellwidth > 0 Then ScrollSerial = ScrollSerial - ScrollLenght
    
    End If

    If ScrollValuetext.ValuetextTime(0) > 23 Then
    
        ScrollLenght = ScrollValuetext.ValuetextTime(0) \ 24
    
        ScrollValuetext.ValuetextTime(0) = ScrollValuetext.ValuetextTime(0) - ScrollLenght * 24
        
        If ScrollControls.ControlDayCellwidth > 0 Then ScrollSerial = ScrollSerial + ScrollLenght

    End If

End If

If ScrollControls.ControlDayCellwidth > 0 Then

    If DayChange <> 0 Then
        ScrollSerial = ScrollSerial + DayChange
        If ScrollControls.ControlDateTimeFormat = "FD" Then
            If ScrollSerial < ScrollMinimum Then ScrollSerial = ScrollMinimum
            If ScrollSerial > ScrollMaximum Then ScrollSerial = ScrollMaximum
        End If
    End If

    ScrollValue = Day(ScrollSerial)

    If ScrollValuetext.ValuetextDay <> ScrollValue Then ScrollValuetext.ValuetextDay = ScrollValue

    ScrollValue = Month(ScrollSerial)

    If ScrollControls.ControlMonthCellwidth > 0 And ScrollValuetext.ValuetextMonth <> ScrollValue Then ScrollValuetext.ValuetextMonth = ScrollValue
    
    ScrollValue = Year(ScrollSerial)

    If ScrollControls.ControlYearCellwidth > 0 And ScrollValuetext.ValuetextYear <> ScrollValue Then ScrollValuetext.ValuetextYear = ScrollValue

End If

If ScrollControls.ControlMonthCellwidth > 0 Then

    If MonthChange <> 0 Then ScrollValuetext.ValuetextMonth = ScrollValuetext.ValuetextMonth + MonthChange

    If ScrollValuetext.ValuetextMonth < 1 Then

        ScrollLenght = -(ScrollValuetext.ValuetextMonth - 1)
        ScrollLenght = (ScrollLenght - 1) \ 12
        ScrollLenght = ScrollLenght + 1

        ScrollValuetext.ValuetextMonth = ScrollValuetext.ValuetextMonth + ScrollLenght * 12
        
        If ScrollControls.ControlYearCellwidth > 0 Then ScrollValuetext.ValuetextYear = ScrollValuetext.ValuetextYear - ScrollLenght

    End If

    If ScrollValuetext.ValuetextMonth > 12 Then

        ScrollLenght = (ScrollValuetext.ValuetextMonth - 1) \ 12
    
        ScrollValuetext.ValuetextMonth = ScrollValuetext.ValuetextMonth - ScrollLenght * 12
        
        If ScrollControls.ControlYearCellwidth > 0 Then ScrollValuetext.ValuetextYear = ScrollValuetext.ValuetextYear + ScrollLenght

    End If

End If

If ScrollControls.ControlYearCellwidth > 0 And YearChange <> 0 Then ScrollValuetext.ValuetextYear = ScrollValuetext.ValuetextYear + YearChange

ScrollSerial = DateSerial(ScrollValuetext.ValuetextYear, ScrollValuetext.ValuetextMonth, ScrollValuetext.ValuetextDay)

If Month(ScrollSerial) <> ScrollValuetext.ValuetextMonth Then ScrollValuetext.ValuetextDay = 1

Call VerifyValuedateLimits(ScrollControls, ScrollValuetext)

ScrollChanged = 0

If ScrollValuetext.ValuetextCrono <> ScrollControls.ControlCurrentValuetext.ValuetextCrono Then ScrollChanged = (ScrollChanged Or 1)
If ScrollValuetext.ValuetextTime(2) <> ScrollControls.ControlCurrentValuetext.ValuetextTime(2) Then ScrollChanged = (ScrollChanged Or 2)
If ScrollValuetext.ValuetextTime(1) <> ScrollControls.ControlCurrentValuetext.ValuetextTime(1) Then ScrollChanged = (ScrollChanged Or 2)
If ScrollValuetext.ValuetextTime(0) <> ScrollControls.ControlCurrentValuetext.ValuetextTime(0) Then ScrollChanged = (ScrollChanged Or 2)
If ScrollValuetext.ValuetextDay <> ScrollControls.ControlCurrentValuetext.ValuetextDay Then ScrollChanged = (ScrollChanged Or 8)
If ScrollValuetext.ValuetextMonth <> ScrollControls.ControlCurrentValuetext.ValuetextMonth Then ScrollChanged = (ScrollChanged Or 8)
If ScrollValuetext.ValuetextYear <> ScrollControls.ControlCurrentValuetext.ValuetextYear Then ScrollChanged = (ScrollChanged Or 4)
    
ScrollControls.ControlCurrentValuetext = ScrollValuetext

If ScrollChanged <> 0 Then Call SwapClickMode(ScrollMe, ScrollControls, "=")

If (ScrollChanged And 12) > 0 Then Call RecalculateCalendar(ScrollControls)

If (ScrollChanged And 12) > 0 Then Call RedrawDateCells(ScrollMe, ScrollControls)
If (ScrollChanged And 12) > 0 Then Call RedrawYearCells(ScrollMe, ScrollControls)
If (ScrollChanged And 14) > 0 Then Call RedrawTimeSecondsCells(ScrollMe, ScrollControls)
If (ScrollChanged And 15) > 0 Then Call RedrawCronoCells(ScrollMe, ScrollControls)

If ScrollChanged = 0 Then Exit Sub

Call RedrawBorders(ScrollMe, ScrollControls)

ScrollMe.RaiseDateTimeChange

End Sub

Public Sub DoScrollMouseWheel(MouseMe As DateTimeStamp, MouseControls As DTSControls, MouseKeys As Long, MouseRotation As Long, MouseOffsetX As Long, MouseOffsetY As Long)

Dim ScrollValue As Long

ScrollValue = MouseRotation \ 120

If MouseControls.ControlReverseScrollFlag = False Then ScrollValue = -ScrollValue

If ScrollValue > 15 Then ScrollValue = 15
If ScrollValue < -15 Then ScrollValue = -15
If ScrollValue = 0 Then Exit Sub

If MouseControls.ControlDayCellwidth > 0 Then
    If MouseOffsetX < MouseControls.ControlMonthCellleft Then Call DoScrollDateTime(MouseMe, MouseControls, (ScrollValue), 0, 0, 0, 0, 0, 0)
End If

If MouseControls.ControlMonthCellwidth > 0 Then
    If MouseOffsetX > MouseControls.ControlMonthCellleft And MouseOffsetX < MouseControls.ControlYearCellleft Then Call DoScrollDateTime(MouseMe, MouseControls, 0, (ScrollValue), 0, 0, 0, 0, 0)
End If

If MouseControls.ControlYearCellwidth > 0 Then
    If MouseOffsetX > MouseControls.ControlYearCellleft And MouseOffsetX < MouseControls.ControlTimeCellleft Then Call DoScrollDateTime(MouseMe, MouseControls, 0, 0, (ScrollValue), 0, 0, 0, 0)
End If

If MouseControls.ControlTimeCellwidth > 0 Then
    If MouseOffsetX > MouseControls.ControlTimeCellleft And MouseOffsetX < MouseControls.ControlCronoCellleft Then
        If MouseOffsetX < MouseControls.ControlTimeCellleft + MouseControls.ControlTimeCellwidth * 2 + 15 Then
            Call DoScrollDateTime(MouseMe, MouseControls, 0, 0, 0, (ScrollValue), 0, 0, 0)
        ElseIf MouseOffsetX < MouseControls.ControlTimeCellleft + MouseControls.ControlTimeCellwidth * 4 + 30 Then
            Call DoScrollDateTime(MouseMe, MouseControls, 0, 0, 0, 0, (ScrollValue), 0, 0)
        Else
            Call DoScrollDateTime(MouseMe, MouseControls, 0, 0, 0, 0, 0, (ScrollValue), 0)
        End If
    End If
End If

If MouseControls.ControlCronoCellwidth > 0 Then
    If MouseOffsetX > MouseControls.ControlCronoCellleft Then Call DoScrollDateTime(MouseMe, MouseControls, 0, 0, 0, 0, 0, 0, (ScrollValue))
End If

End Sub

Public Sub RedrawAllHeaders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)
    
Call RedrawDateHeaders(RedrawMe, RedrawControls)
Call RedrawYearHeaders(RedrawMe, RedrawControls)
Call RedrawTimeSecondsHeaders(RedrawMe, RedrawControls)
Call RedrawCronoHeaders(RedrawMe, RedrawControls)

End Sub

Public Sub RedrawAllCells(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Call SetYearOffset(RedrawControls)

Call RedrawDateCells(RedrawMe, RedrawControls)
Call RedrawYearCells(RedrawMe, RedrawControls)
Call RedrawTimeSecondsCells(RedrawMe, RedrawControls)
Call RedrawCronoCells(RedrawMe, RedrawControls)

End Sub

Public Sub RedrawBorders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim DrawRight As Long
Dim DrawBottom As Long

If RedrawControls.ControlFlatMode = True Then Exit Sub

DrawRight = RedrawMe.ScaleWidth - 15
DrawBottom = RedrawMe.ScaleHeight - 15

RedrawMe.DrawLine 0, 0, DrawRight, 0, &HA0A0A0
RedrawMe.DrawLine 0, 0, 0, DrawBottom, &HA0A0A0
RedrawMe.DrawLine 15, 15, DrawRight - 15, 15, &H696969
RedrawMe.DrawLine 15, 15, 15, DrawBottom - 15, &H696969
RedrawMe.DrawLine DrawRight - 15, 15, DrawRight - 15, DrawBottom, &HE3E3E3
RedrawMe.DrawLine 15, DrawBottom - 15, DrawRight, DrawBottom - 15, &HE3E3E3
RedrawMe.DrawLine DrawRight, 0, DrawRight, DrawBottom, &HFFFFFF
RedrawMe.DrawLine 0, DrawBottom, DrawRight + 15, DrawBottom, &HFFFFFF

End Sub

Public Sub RedrawDateHeaders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawCell As Integer
Dim DrawColor As OLE_COLOR
Dim DrawWeekday As Integer
Dim DrawText As String
Dim DrawX As Long

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlHeaderCellheight <= 0 Then Exit Sub

If RedrawControls.ControlDayCellwidth > 0 Then

    DrawWeekday = RedrawControls.ControlWeekdayStart + 1

    For DrawCell = 0 To 6

        Select Case DrawCell
            Case 0: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6))
            Case 1: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0))
            Case 2: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1))
            Case 3: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2))
            Case 4: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3))
            Case 5: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(5), RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4))
            Case 6: DrawText = Choose(DrawWeekday, RedrawControls.ControlWeekdayNames(6), RedrawControls.ControlWeekdayNames(0), RedrawControls.ControlWeekdayNames(1), RedrawControls.ControlWeekdayNames(2), RedrawControls.ControlWeekdayNames(3), RedrawControls.ControlWeekdayNames(4), RedrawControls.ControlWeekdayNames(5))
        End Select

        DrawLeft = DrawCell * RedrawControls.ControlDayCellwidth
        DrawRight = DrawLeft + RedrawControls.ControlDayCellwidth
    
        DrawColor = RedrawControls.ControlHeaderBackground
        
        RedrawMe.DrawLine DrawLeft, 0, DrawRight, RedrawControls.ControlHeaderCellheight, DTSCellBordercolor, "BF"
        RedrawMe.DrawLine DrawLeft + 15, 15, DrawRight - 15, RedrawControls.ControlHeaderCellheight - 30, DrawColor, "BF"

        DrawX = DrawLeft + (RedrawControls.ControlDayCellwidth - RedrawMe.TextWidth(DrawText)) / 2

        While DrawX < (DrawLeft + 30) And DrawText <> ""
    
            DrawText = Left$(DrawText, Len(DrawText) - 1)
            DrawX = DrawLeft + (RedrawControls.ControlDayCellwidth - RedrawMe.TextWidth(DrawText)) / 2
    
        Wend
    
        RedrawMe.DrawPrint DrawX, (RedrawControls.ControlHeaderCellheight - RedrawMe.TextHeight(DrawText)) / 2, RedrawControls.ControlHeaderForeground, DrawText

    Next

End If

If RedrawControls.ControlMonthCellwidth > 0 Then

    DrawRight = RedrawControls.ControlMonthCellleft + RedrawControls.ControlMonthCellwidth * 3

    DrawColor = RedrawControls.ControlHeaderBackground
    
    RedrawMe.DrawLine RedrawControls.ControlMonthCellleft, 0, DrawRight, RedrawControls.ControlHeaderCellheight, DTSCellBordercolor, "BF"
    RedrawMe.DrawLine RedrawControls.ControlMonthCellleft + 15, 15, DrawRight - 15, RedrawControls.ControlHeaderCellheight - 30, DrawColor, "BF"

    DrawText = RedrawControls.ControlHeaderMonth

    RedrawMe.DrawPrint RedrawControls.ControlMonthCellleft + (RedrawControls.ControlMonthCellwidth * 3 - RedrawMe.TextWidth(DrawText)) / 2, (RedrawControls.ControlHeaderCellheight - RedrawMe.TextHeight(DrawText)) / 2, RedrawControls.ControlHeaderForeground, DrawText

End If

End Sub

Public Sub RedrawDateCells(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim RedrawCell As Integer

If RedrawControls.ControlInitialized = False Then Exit Sub

If RedrawControls.ControlDayCellwidth > 0 Then
    For RedrawCell = 0 To 41
        Call DrawDayCell(RedrawMe, RedrawControls, RedrawCell)
    Next
End If

If RedrawControls.ControlMonthCellwidth > 0 Then
    For RedrawCell = 0 To 11
        Call DrawMonthCell(RedrawMe, RedrawControls, RedrawCell)
    Next
End If

End Sub

Private Sub RedrawTimeSecondsCells(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim RedrawCell As Integer
Dim RedrawTotal As Integer

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlTimeCellwidth <= 0 Then Exit Sub

RedrawTotal = 2

If Len(RedrawControls.ControlDateTimeFormat) > 1 Then
    If Left$(RedrawControls.ControlDateTimeFormat, Len(RedrawControls.ControlDateTimeFormat) - 1) = "R" Then RedrawTotal = 1
End If

For RedrawCell = 0 To RedrawTotal
    Call DrawTimeCell(RedrawMe, RedrawControls, RedrawCell)
Next

Call DrawGridCells(RedrawMe, RedrawControls, RedrawControls.ControlTimeCellleft, RedrawControls.ControlTimeCellwidth, RedrawControls.ControlTimeCellheight, "H")
Call DrawGridCells(RedrawMe, RedrawControls, RedrawControls.ControlTimeCellleft + RedrawControls.ControlTimeCellwidth * 2 + 15, RedrawControls.ControlTimeCellwidth, RedrawControls.ControlTimeCellheight, "M")

If RedrawTotal = 2 Then Call DrawGridCells(RedrawMe, RedrawControls, RedrawControls.ControlTimeCellleft + RedrawControls.ControlTimeCellwidth * 4 + 30, RedrawControls.ControlTimeCellwidth, RedrawControls.ControlTimeCellheight, "S")

End Sub

Private Sub RedrawYearCells(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlYearCellwidth <= 0 Then Exit Sub

If GetYearMode(RedrawControls, 2) = 0 Then

    Call DrawYearItem(RedrawMe, RedrawControls, False, 0, RedrawControls.ControlCurrentValuetext.ValuetextYear)
    Call DrawGridCells(RedrawMe, RedrawControls, RedrawControls.ControlYearCellleft, RedrawControls.ControlYearCellwidth, RedrawControls.ControlYearCellheight, "Y")

Else

    Call DrawYearList(RedrawMe, RedrawControls)
    
End If

End Sub

Private Sub RedrawCronoCells(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlCronoCellwidth <= 0 Then Exit Sub

Call DrawCronoCell(RedrawMe, RedrawControls)
Call DrawGridCells(RedrawMe, RedrawControls, RedrawControls.ControlCronoCellleft, RedrawControls.ControlCronoCellwidth, RedrawControls.ControlCronoCellheight, "E")

End Sub

Private Sub RedrawYearHeaders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim DrawRight As Long
Dim DrawColor As OLE_COLOR
Dim DrawText As String

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlHeaderCellheight <= 0 Then Exit Sub
If RedrawControls.ControlYearCellwidth <= 0 Then Exit Sub

DrawRight = RedrawControls.ControlYearCellleft + RedrawControls.ControlYearCellwidth * 2

DrawColor = RedrawControls.ControlHeaderBackground

RedrawMe.DrawLine RedrawControls.ControlYearCellleft, 0, DrawRight, RedrawControls.ControlHeaderCellheight, DTSCellBordercolor, "BF"
RedrawMe.DrawLine RedrawControls.ControlYearCellleft + 15, 15, DrawRight - 15, RedrawControls.ControlHeaderCellheight - 30, DrawColor, "BF"

DrawText = RedrawControls.ControlHeaderYear

RedrawMe.DrawPrint RedrawControls.ControlYearCellleft + RedrawControls.ControlYearCellwidth - RedrawMe.TextWidth(DrawText) / 2, (RedrawControls.ControlHeaderCellheight - RedrawMe.TextHeight(DrawText)) / 2, RedrawControls.ControlHeaderForeground, DrawText

End Sub

Private Sub RedrawCronoHeaders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim DrawRight As Long
Dim DrawColor As OLE_COLOR
Dim DrawText As String

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlHeaderCellheight <= 0 Then Exit Sub
If RedrawControls.ControlCronoCellwidth <= 0 Then Exit Sub

DrawRight = RedrawControls.ControlCronoCellleft + RedrawControls.ControlCronoCellwidth * 2

DrawColor = RedrawControls.ControlHeaderBackground

RedrawMe.DrawLine RedrawControls.ControlCronoCellleft, 0, DrawRight, RedrawControls.ControlHeaderCellheight, DTSCellBordercolor, "BF"
RedrawMe.DrawLine RedrawControls.ControlCronoCellleft + 15, 15, DrawRight - 15, RedrawControls.ControlHeaderCellheight - 30, DrawColor, "BF"

DrawText = "1/1" + String(RedrawControls.ControlCronoDecimals, 48)

RedrawMe.DrawPrint RedrawControls.ControlCronoCellleft + RedrawControls.ControlCronoCellwidth - RedrawMe.TextWidth(DrawText) / 2, (RedrawControls.ControlHeaderCellheight - RedrawMe.TextHeight(DrawText)) / 2, RedrawControls.ControlHeaderForeground, DrawText

End Sub

Private Sub RedrawTimeSecondsHeaders(RedrawMe As DateTimeStamp, RedrawControls As DTSControls)

Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawMode As Integer
Dim DrawColor As OLE_COLOR
Dim DrawTotal As Integer
Dim DrawText As String

If RedrawControls.ControlInitialized = False Then Exit Sub
If RedrawControls.ControlHeaderCellheight <= 0 Then Exit Sub
If RedrawControls.ControlTimeCellwidth <= 0 Then Exit Sub
   
DrawRight = RedrawControls.ControlTimeCellleft - 15
DrawTotal = 2

If Len(RedrawControls.ControlDateTimeFormat) > 1 Then
    If Left$(RedrawControls.ControlDateTimeFormat, Len(RedrawControls.ControlDateTimeFormat) - 1) = "R" Then DrawTotal = 1
End If

DrawColor = RedrawControls.ControlHeaderBackground

For DrawMode = 0 To 2

    DrawLeft = DrawRight + 15
    DrawRight = DrawLeft + RedrawControls.ControlTimeCellwidth * 2

    RedrawMe.DrawLine DrawLeft, 0, DrawRight, RedrawControls.ControlHeaderCellheight, DTSCellBordercolor, "BF"
    RedrawMe.DrawLine DrawLeft + 15, 15, DrawRight - 15, RedrawControls.ControlHeaderCellheight - 30, DrawColor, "BF"

    DrawText = RedrawControls.ControlHeaderTime(DrawMode)

    RedrawMe.DrawPrint DrawLeft + RedrawControls.ControlTimeCellwidth - RedrawMe.TextWidth(DrawText) / 2, (RedrawControls.ControlHeaderCellheight - RedrawMe.TextHeight(DrawText)) / 2, RedrawControls.ControlHeaderForeground, DrawText

Next

End Sub

Private Sub DrawTimeChange(DrawMe As DateTimeStamp, DrawControls As DTSControls, DrawCell As Integer)

Call DrawTimeCell(DrawMe, DrawControls, DrawCell)

If DrawCell < 1 Then Call DrawTimeCell(DrawMe, DrawControls, 1)

If Len(DrawControls.ControlDateTimeFormat) > 1 Then
    If Left$(DrawControls.ControlDateTimeFormat, Len(DrawControls.ControlDateTimeFormat) - 1) = "R" Then Exit Sub
End If

If DrawCell < 2 Then Call DrawTimeCell(DrawMe, DrawControls, 2)

If DrawControls.ControlCronoCellwidth > 0 Then Call DrawCronoCell(DrawMe, DrawControls)

End Sub

Private Sub DrawTimeCell(DrawMe As DateTimeStamp, DrawControls As DTSControls, DrawCell As Integer)

Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawColor As OLE_COLOR
Dim DrawText As String

DrawBottom = DrawControls.ControlHeaderCellheight + DrawControls.ControlTimeCellheight
DrawLeft = DrawControls.ControlTimeCellleft + (DrawControls.ControlTimeCellwidth * 2 + 15) * DrawCell
DrawRight = DrawLeft + DrawControls.ControlTimeCellwidth * 2

DrawColor = DrawControls.ControlSelectedBackground

DrawMe.DrawLine DrawLeft, DrawControls.ControlHeaderCellheight, DrawRight, DrawBottom, DTSCellBordercolor, "BF"
DrawMe.DrawLine DrawLeft + 15, DrawControls.ControlHeaderCellheight + 15, DrawRight - 15, DrawBottom - 15, DrawColor, "BF"

DrawText = DatePartFormat(DrawControls.ControlCurrentValuetext.ValuetextTime(DrawCell))

DrawMe.DrawPrint DrawLeft + DrawControls.ControlTimeCellwidth - DrawMe.TextWidth(DrawText) / 2, DrawControls.ControlHeaderCellheight + (DrawControls.ControlTimeCellheight - DrawMe.TextHeight(DrawText)) / 2, DrawControls.ControlSelectedForeground, DrawText

End Sub

Private Sub DrawCronoCell(DrawMe As DateTimeStamp, DrawControls As DTSControls)

Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawColor As OLE_COLOR
Dim DrawText As String

DrawBottom = DrawControls.ControlHeaderCellheight + DrawControls.ControlCronoCellheight
DrawRight = DrawControls.ControlCronoCellleft + DrawControls.ControlCronoCellwidth * 2

DrawColor = DrawControls.ControlSelectedBackground

DrawMe.DrawLine DrawControls.ControlCronoCellleft, DrawControls.ControlHeaderCellheight, DrawRight, DrawBottom, DTSCellBordercolor, "B"
DrawMe.DrawLine DrawControls.ControlCronoCellleft + 15, DrawControls.ControlHeaderCellheight + 15, DrawRight - 15, DrawBottom - 15, DrawColor, "BF"

DrawText = Right$(String$(DrawControls.ControlCronoDecimals, 48) + Trim$(Str$(DrawControls.ControlCurrentValuetext.ValuetextCrono)), DrawControls.ControlCronoDecimals)

DrawMe.DrawPrint DrawControls.ControlCronoCellleft + DrawControls.ControlCronoCellwidth - DrawMe.TextWidth(DrawText) / 2, DrawControls.ControlHeaderCellheight + (DrawControls.ControlCronoCellheight - DrawMe.TextHeight(DrawText)) / 2, DrawControls.ControlSelectedForeground, DrawText

End Sub

Private Sub DrawYearGrid(DrawMe As DateTimeStamp, DrawControls As DTSControls)

If DrawControls.ControlYearCellwidth <= 0 Then Exit Sub

If GetYearMode(DrawControls, 2) = 0 Then

    Call DrawGridCells(DrawMe, DrawControls, DrawControls.ControlYearCellleft, DrawControls.ControlYearCellwidth, DrawControls.ControlYearCellheight, "Y")

Else

    Call DrawYearList(DrawMe, DrawControls)
    
End If

End Sub

Private Sub DrawYearList(DrawMe As DateTimeStamp, DrawControls As DTSControls)

Dim DrawCell As Integer
Dim DrawYearValue As Integer
Dim DrawTotal As Integer

DrawYearValue = DrawControls.ControlYearStart

If DrawControls.ControlYearEnd - DrawControls.ControlYearStart > 5 Then
    DrawTotal = 11
Else
    DrawTotal = 5
End If

For DrawCell = 0 To DrawTotal
    
    If DrawTotal = 5 Then
        Call DrawYearItem(DrawMe, DrawControls, False, DrawCell, DrawYearValue)
    Else
        Call DrawYearItem(DrawMe, DrawControls, True, DrawCell, DrawYearValue)
    End If
    
    Select Case DrawYearValue
        Case DrawControls.ControlYearEnd: DrawYearValue = 0
        Case Is <> 0:                     DrawYearValue = DrawYearValue + 1
    End Select
    
Next

End Sub

Private Sub DrawYearItem(DrawMe As DateTimeStamp, DrawControls As DTSControls, DrawDouble As Boolean, DrawCell As Integer, DrawYearValue As Integer)

Dim DrawTop As Long
Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawPrintcolor As OLE_COLOR
Dim DrawLinecolor As OLE_COLOR
Dim DrawText As String

If DrawDouble = False Then
    
    DrawLeft = DrawControls.ControlYearCellleft
    DrawRight = DrawLeft + DrawControls.ControlYearCellwidth * 2
    DrawTop = DrawControls.ControlHeaderCellheight + DrawControls.ControlYearCellheight * DrawCell

Else

    If (DrawCell And 1) = 0 Then
        DrawLeft = DrawControls.ControlYearCellleft
    Else
        DrawLeft = DrawControls.ControlYearCellleft + DrawControls.ControlYearCellwidth
    End If
    
    DrawTop = DrawControls.ControlHeaderCellheight + DrawControls.ControlYearCellheight * (DrawCell \ 2)
    DrawRight = DrawLeft + DrawControls.ControlYearCellwidth

End If

DrawBottom = DrawTop + DrawControls.ControlYearCellheight

If DrawYearValue = 0 Then
    DrawLinecolor = DrawControls.ControlEmptyBackground
    DrawPrintcolor = DrawControls.ControlSelectedForeground
    DrawText = ""
Else
    DrawText = Right$(Str$(10000 + DrawYearValue), 4)
    If DrawYearValue = DrawControls.ControlCurrentValuetext.ValuetextYear Then
        DrawPrintcolor = DrawControls.ControlSelectedForeground
        DrawLinecolor = DrawControls.ControlSelectedBackground
    Else
        DrawPrintcolor = DrawControls.ControlNormalForeground
        DrawLinecolor = DrawControls.ControlNormalBackground
    End If
End If

DrawMe.DrawLine DrawLeft, DrawTop, DrawRight, DrawBottom, DTSCellBordercolor, "B"
DrawMe.DrawLine DrawLeft + 15, DrawTop + 15, DrawRight - 15, DrawBottom - 15, DrawLinecolor, "BF"

DrawMe.DrawPrint (DrawLeft + DrawRight - DrawMe.TextWidth(DrawText)) / 2, DrawTop + (DrawControls.ControlYearCellheight - DrawMe.TextHeight(DrawText)) / 2, DrawPrintcolor, DrawText

End Sub

Private Sub DrawGridCells(DrawMe As DateTimeStamp, DrawControls As DTSControls, RedrawLeft As Long, DrawCellwidth As Long, DrawCellheight As Long, DrawMode As String)

Dim DrawTop As Long
Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawLimit As Integer
Dim DrawColor As OLE_COLOR
Dim DrawMaximum As Integer
Dim DrawValue As Integer
Dim DrawRow As Integer
Dim DrawCol As Integer
Dim DrawText As String
Dim DrawStart As String
Dim DrawEnd As String

DrawColor = DrawControls.ControlNormalBackground

DrawStart = ""
DrawEnd = ""

Select Case DrawMode
    Case "Y"
        If DrawControls.ControlClickMode = "Y" Then
            DrawValue = DrawControls.ControlCurrentValuetext.ValuetextYear \ 10
            DrawValue = DrawValue * 10
            DrawMaximum = DrawValue + 9
            DrawLimit = 34
        Else
            DrawLimit = 33
            DrawMaximum = 299
            DrawValue = DrawControls.ControlYearOffset
            DrawEnd = "x"
        End If
    Case "H"
        If DrawControls.ControlClickMode = "H" Then
            DrawValue = DrawControls.ControlCurrentValuetext.ValuetextTime(0) \ 10
            DrawValue = DrawValue * 10
            DrawMaximum = 23
            DrawLimit = 42
        Else
            DrawLimit = 41
            DrawMaximum = 2
            DrawValue = 0
            DrawEnd = "x"
        End If
    Case "M"
        If DrawControls.ControlClickMode = "M" Then
            DrawValue = DrawControls.ControlCurrentValuetext.ValuetextTime(1) \ 10
            DrawValue = DrawValue * 10
            DrawMaximum = 59
            DrawLimit = 52
        Else
            DrawLimit = 51
            DrawMaximum = 5
            DrawValue = 0
            DrawEnd = "x"
        End If
    Case "S"
        If DrawControls.ControlClickMode = "S" Then
            DrawValue = DrawControls.ControlCurrentValuetext.ValuetextTime(2) \ 10
            DrawValue = DrawValue * 10
            DrawMaximum = 59
            DrawLimit = 62
        Else
            DrawLimit = 61
            DrawMaximum = 5
            DrawValue = 0
            DrawEnd = "x"
        End If
    Case "E"
        DrawValue = Val(DrawControls.ControlClickMode)
        DrawStart = Left$(Right$(String$(DrawControls.ControlCronoDecimals, 48) + Trim$(Str$(DrawControls.ControlCurrentValuetext.ValuetextCrono)), DrawControls.ControlCronoDecimals), DrawValue)
        DrawValue = DrawControls.ControlCronoDecimals - DrawValue - 1
        If DrawValue > 0 Then DrawEnd = String$(DrawValue, 120)
        DrawLimit = 76 - DrawValue
        DrawMaximum = 9
        DrawValue = 0
    Case Else
        DrawMaximum = 0
        DrawValue = 1
End Select

For DrawRow = 0 To 4
    
    DrawTop = DrawControls.ControlHeaderCellheight + (DrawRow + 1) * DrawCellheight
    DrawBottom = DrawTop + DrawCellheight

    For DrawCol = 0 To 1
    
        DrawLeft = RedrawLeft + DrawCol * DrawCellwidth
        DrawRight = DrawLeft + DrawCellwidth
        
        If DrawValue > DrawMaximum Then
            DrawColor = DrawControls.ControlEmptyBackground
        Else
            DrawColor = ReturnLimitColor(DrawControls, DrawLimit, Val(DrawStart + Trim$(Str$(DrawValue))))
        End If
        
        DrawMe.DrawLine DrawLeft, DrawTop, DrawRight, DrawBottom, DTSCellBordercolor, "B"
        DrawMe.DrawLine DrawLeft + 15, DrawTop + 15, DrawRight - 15, DrawBottom - 15, DrawColor, "BF"
        
        If DrawValue <= DrawMaximum Then
        
            DrawText = DrawStart + Trim$(Str$(DrawValue)) + DrawEnd
            
            If DrawMode = "Y" And DrawLimit = 33 Then
                If DrawRow = 0 And DrawCol = 0 And DrawControls.ControlYearOffset > 100 Then DrawText = "<<"
                If DrawRow = 4 And DrawCol = 1 And DrawControls.ControlYearOffset < 290 Then DrawText = ">>"
            End If
            
            DrawMe.DrawPrint DrawLeft + (DrawCellwidth - DrawMe.TextWidth(DrawText)) / 2, DrawTop + (DrawCellheight - DrawMe.TextHeight(DrawText)) / 2, DrawControls.ControlNormalForeground, DrawText
    
        End If
        
        DrawValue = DrawValue + 1
    Next

Next

End Sub

Private Sub DrawMonthCell(DrawMe As DateTimeStamp, DrawControls As DTSControls, ByVal DrawCell As Integer)

Dim DrawTop As Long
Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawPrintcolor As OLE_COLOR
Dim DrawLinecolor As OLE_COLOR
Dim DrawText As String
Dim DrawX As Long

DrawTop = (DrawCell \ 3) * DrawControls.ControlMonthCellheight + DrawControls.ControlHeaderCellheight
DrawLeft = (DrawCell Mod 3) * DrawControls.ControlMonthCellwidth + DrawControls.ControlMonthCellleft
DrawBottom = DrawTop + DrawControls.ControlMonthCellheight
DrawRight = DrawLeft + DrawControls.ControlMonthCellwidth

If DrawCell + 1 = DrawControls.ControlCurrentValuetext.ValuetextMonth Then
    DrawLinecolor = DrawControls.ControlSelectedBackground
    DrawPrintcolor = DrawControls.ControlSelectedForeground
    DrawControls.ControlMonthCellnumber = DrawCell + 1
Else
    DrawLinecolor = ReturnLimitColor(DrawControls, 22, DrawCell + 1)
    DrawPrintcolor = DrawControls.ControlNormalForeground
End If

DrawMe.DrawLine DrawLeft, DrawTop, DrawRight, DrawBottom, DTSCellBordercolor, "B"
DrawMe.DrawLine DrawLeft + 15, DrawTop + 15, DrawRight - 15, DrawBottom - 15, DrawLinecolor, "BF"

DrawText = DatePartFormat(DrawCell + 1)

DrawMe.DrawPrint DrawLeft + (DrawControls.ControlMonthCellwidth - DrawMe.TextWidth(DrawText)) / 2, DrawTop + DrawControls.ControlMonthCellheight / 3 - DrawMe.TextHeight(DrawText) / 2, DrawPrintcolor, DrawText

DrawText = DrawControls.ControlMonthNames(DrawCell)

DrawX = DrawLeft + (DrawControls.ControlMonthCellwidth - DrawMe.TextWidth(DrawText)) / 2

While DrawX < (DrawLeft + 30) And DrawText <> ""
    
    DrawText = Left$(DrawText, Len(DrawText) - 1)
    
    DrawX = DrawLeft + (DrawControls.ControlMonthCellwidth - DrawMe.TextWidth(DrawText)) / 2
    
Wend

DrawMe.DrawPrint DrawX, DrawTop + (DrawControls.ControlMonthCellheight / 3) * 2 - DrawMe.TextHeight(DrawText) / 2, DrawPrintcolor, DrawText

End Sub

Private Sub DrawDayCell(DrawMe As DateTimeStamp, DrawControls As DTSControls, ByVal DrawCell As Integer)

Dim DrawTop As Long
Dim DrawLeft As Long
Dim DrawRight As Long
Dim DrawBottom As Long
Dim DrawPrintcolor As OLE_COLOR
Dim DrawLinecolor As OLE_COLOR
Dim DrawNowflag As Boolean
Dim DrawText As String
Dim DrawMarker As Long

DrawNowflag = DrawControls.ControlNowFlag
DrawTop = (DrawCell \ 7) * DrawControls.ControlDayCellheight + DrawControls.ControlHeaderCellheight
DrawBottom = DrawTop + DrawControls.ControlDayCellheight

If DrawControls.ControlDateTimeFormat = "FD" Then
    If DrawControls.ControlCurrentValuetext.ValuetextMonth <> DrawControls.ControlNowMonth Or DrawControls.ControlCurrentValuetext.ValuetextYear <> DrawControls.ControlNowYear Then DrawNowflag = False
End If

DrawText = Trim$(DrawControls.ControlDayCelltext(DrawCell))
DrawLeft = (DrawCell Mod 7) * DrawControls.ControlDayCellwidth
DrawRight = DrawLeft + DrawControls.ControlDayCellwidth
DrawMarker = 0

If DrawCell = 41 And DrawNowflag = True Then
    DrawLinecolor = DrawControls.ControlNormalBackground
    DrawPrintcolor = DrawControls.ControlNormalForeground
Else
    Select Case DrawControls.ControlDayCelltype(DrawCell)
        Case TypeNormal
            DrawMarker = Val(DrawText)
            DrawLinecolor = ReturnLimitColor(DrawControls, 12, DrawMarker)
            DrawPrintcolor = DrawControls.ControlNormalForeground
        Case TypeSelected
            DrawMarker = Val(DrawText)
            DrawLinecolor = DrawControls.ControlSelectedBackground
            DrawPrintcolor = DrawControls.ControlSelectedForeground
            DrawControls.ControlDayCellnumber = DrawCell
        Case TypeEmpty
            DrawLinecolor = DrawControls.ControlEmptyBackground
            DrawPrintcolor = DrawControls.ControlNormalForeground
    End Select
End If

DrawMe.DrawLine DrawLeft, DrawTop, DrawRight, DrawBottom, DTSCellBordercolor, "B"
DrawMe.DrawLine DrawLeft + 15, DrawTop + 15, DrawRight - 15, DrawBottom - 15, DrawLinecolor, "BF"

If DrawCell = 41 And DrawNowflag = True Then

    DrawLeft = DrawLeft + 60
    DrawRight = DrawRight - 45
    DrawBottom = DrawBottom - 60
    DrawTop = DrawTop + 60
    
    DrawMe.DrawLine DrawLeft, DrawTop + 30, DrawRight, DrawTop + 30, DTSCellBordercolor
    DrawMe.DrawLine DrawLeft, DrawBottom - 45, DrawRight, DrawBottom - 45, DTSCellBordercolor
    DrawMe.DrawLine DrawRight - 45, DrawTop, DrawRight - 45, DrawBottom, DTSCellBordercolor
    DrawMe.DrawLine DrawLeft + 30, DrawTop, DrawLeft + 30, DrawBottom, DTSCellBordercolor
    
    DrawMe.DrawLine DrawLeft + 45, DrawTop + 45, DrawRight - 60, DrawBottom - 60, DrawControls.ControlNowColor, "B"
    DrawMe.DrawLine DrawLeft + 60, DrawTop + 60, DrawRight - 75, DrawBottom - 75, DrawControls.ControlNowColor, "B"
    
Else

    If DrawControls.ControlNowFlag = True Then
        If DrawMarker = DrawControls.ControlNowDay And DrawControls.ControlCurrentValuetext.ValuetextMonth = DrawControls.ControlNowMonth And DrawControls.ControlCurrentValuetext.ValuetextYear = DrawControls.ControlNowYear Then
            DrawMe.DrawLine DrawLeft + 30, DrawTop + 30, DrawRight - 30, DrawBottom - 30, DrawControls.ControlNowColor, "B"
            DrawMe.DrawLine DrawLeft + 15, DrawTop + 15, DrawRight - 15, DrawBottom - 15, DrawControls.ControlNowColor, "B"
        End If
    End If

    DrawMe.DrawPrint DrawLeft + (DrawControls.ControlDayCellwidth - DrawMe.TextWidth(DrawText)) / 2, DrawTop + (DrawControls.ControlDayCellheight - DrawMe.TextHeight(DrawText)) / 2, DrawPrintcolor, DrawText

End If

End Sub

Private Sub DrawForLimitCells(DrawMe As DateTimeStamp, DrawControls As DTSControls, DrawStatus As Integer)

Dim DrawFlag As Boolean

If DrawControls.ControlLimitMode = 0 Then Exit Sub

If DrawControls.ControlTimeCellwidth > 0 Then
    
    If DrawStatus < 4 Then Call DrawGridCells(DrawMe, DrawControls, DrawControls.ControlTimeCellleft, DrawControls.ControlTimeCellwidth, DrawControls.ControlTimeCellheight, "H")
    If DrawStatus < 5 Then Call DrawGridCells(DrawMe, DrawControls, DrawControls.ControlTimeCellleft + DrawControls.ControlTimeCellwidth * 2 + 15, DrawControls.ControlTimeCellwidth, DrawControls.ControlTimeCellheight, "M")

    DrawFlag = False
    If Len(DrawControls.ControlDateTimeFormat) > 1 Then
        If Left$(DrawControls.ControlDateTimeFormat, Len(DrawControls.ControlDateTimeFormat) - 1) = "R" Then DrawFlag = True
    End If

    If DrawFlag = False Then
        If DrawStatus < 6 Then Call DrawGridCells(DrawMe, DrawControls, DrawControls.ControlTimeCellleft + DrawControls.ControlTimeCellwidth * 4 + 30, DrawControls.ControlTimeCellwidth, DrawControls.ControlTimeCellheight, "S")
    End If

End If

If DrawControls.ControlCronoCellwidth > 0 Then
    Call DrawGridCells(DrawMe, DrawControls, DrawControls.ControlCronoCellleft, DrawControls.ControlCronoCellwidth, DrawControls.ControlCronoCellheight, "E")
End If

End Sub

Private Sub SwapClickMode(NewMe As DateTimeStamp, NewControls As DTSControls, NewMode As String)

Dim OldMode As String

OldMode = NewControls.ControlClickMode
NewControls.ControlClickMode = NewMode

If NewMode <> "Y" Then
    Call SetYearOffset(NewControls)
    If OldMode <> "Y" Then Call DrawYearGrid(NewMe, NewControls)
End If

Select Case OldMode
    Case "Y"
        Call DrawYearGrid(NewMe, NewControls)
    Case "H"
        Call DrawGridCells(NewMe, NewControls, NewControls.ControlTimeCellleft, NewControls.ControlTimeCellwidth, NewControls.ControlTimeCellheight, "H")
    Case "M"
        Call DrawGridCells(NewMe, NewControls, NewControls.ControlTimeCellleft + NewControls.ControlTimeCellwidth * 2 + 15, NewControls.ControlTimeCellwidth, NewControls.ControlTimeCellheight, "M")
    Case "S"
        Call DrawGridCells(NewMe, NewControls, NewControls.ControlTimeCellleft + NewControls.ControlTimeCellwidth * 4 + 30, NewControls.ControlTimeCellwidth, NewControls.ControlTimeCellheight, "S")
    Case "1" To "5"
        If NewMode < "1" Or NewMode > "5" Then Call DrawGridCells(NewMe, NewControls, NewControls.ControlCronoCellleft, NewControls.ControlCronoCellwidth, NewControls.ControlCronoCellheight, "E")
End Select

End Sub

Private Sub VerifyMonthChange(VerifyMe As DateTimeStamp, VerifyControls As DTSControls)

Dim VerifyDate As Date

VerifyDate = DateSerial(VerifyControls.ControlCurrentValuetext.ValuetextYear, VerifyControls.ControlCurrentValuetext.ValuetextMonth, VerifyControls.ControlCurrentValuetext.ValuetextDay)

If Month(VerifyDate) <> VerifyControls.ControlCurrentValuetext.ValuetextMonth Then VerifyControls.ControlCurrentValuetext.ValuetextDay = 1

Call RecalculateCalendar(VerifyControls)

Call RedrawAllCells(VerifyMe, VerifyControls)
Call RedrawBorders(VerifyMe, VerifyControls)

VerifyMe.RaiseDateTimeChange

End Sub

Private Sub MouseDayChange(MouseMe As DateTimeStamp, MouseControls As DTSControls, MouseX As Single, MouseY As Single)

Dim MouseCell As Long
Dim MouseNowFlag As Boolean
Dim MouseDatevalue As DTSValuetext
Dim MouseDay As Long

MouseCell = (MouseY - MouseControls.ControlHeaderCellheight) \ MouseControls.ControlDayCellheight
MouseCell = MouseCell * 7 + (MouseX \ MouseControls.ControlDayCellwidth)

MouseNowFlag = MouseControls.ControlNowFlag

If MouseControls.ControlDateTimeFormat = "FD" Then
    If MouseControls.ControlCurrentValuetext.ValuetextMonth <> MouseControls.ControlNowMonth Or MouseControls.ControlCurrentValuetext.ValuetextYear <> MouseControls.ControlNowYear Then MouseNowFlag = False
End If

If MouseCell = 41 And MouseNowFlag = True Then Call MouseSwapToday(MouseMe, MouseControls)

If MouseCell > MouseControls.ControlLastDay + MouseControls.ControlDayOffset Or MouseControls.ControlDayCelltype(MouseCell) = TypeEmpty Then Exit Sub
If MouseCell = MouseControls.ControlDayCellnumber Then Exit Sub

MouseDay = Val(MouseControls.ControlDayCelltext(MouseCell))

MouseDatevalue = MouseControls.ControlCurrentValuetext

MouseDatevalue.ValuetextDay = MouseDay

Call VerifyValuedateLimits(MouseControls, MouseDatevalue)

If MouseDatevalue.ValuetextDay <> MouseDay Then Exit Sub

MouseControls.ControlCurrentValuetext = MouseDatevalue

If MouseControls.ControlDayCellnumber >= 0 Then

    MouseControls.ControlDayCelltype(MouseControls.ControlDayCellnumber) = TypeNormal

    Call DrawDayCell(MouseMe, MouseControls, MouseControls.ControlDayCellnumber)

End If

MouseControls.ControlDayCellnumber = MouseCell

MouseControls.ControlDayCelltype(MouseCell) = TypeSelected

Call SwapClickMode(MouseMe, MouseControls, "=")
Call DrawDayCell(MouseMe, MouseControls, MouseCell)
Call DrawForLimitCells(MouseMe, MouseControls, 3)
Call RedrawBorders(MouseMe, MouseControls)

MouseMe.RaiseDateTimeChange

End Sub

Private Sub MouseMonthChange(MouseMe As DateTimeStamp, MouseControls As DTSControls, MouseX As Single, MouseY As Single)

Dim MouseCell As Long
Dim MouseDatevalue As DTSValuetext

MouseCell = (MouseY - MouseControls.ControlHeaderCellheight) \ MouseControls.ControlMonthCellheight
MouseCell = MouseCell * 3 + ((MouseX - MouseControls.ControlMonthCellleft) \ MouseControls.ControlMonthCellwidth) + 1

If MouseCell = MouseControls.ControlMonthCellnumber Then Exit Sub

MouseDatevalue = MouseControls.ControlCurrentValuetext

MouseDatevalue.ValuetextMonth = MouseCell

Call VerifyValuedateLimits(MouseControls, MouseDatevalue)

MouseControls.ControlCurrentValuetext = MouseDatevalue

Call SwapClickMode(MouseMe, MouseControls, "=")
Call VerifyMonthChange(MouseMe, MouseControls)

End Sub

Private Sub MouseTimeChange(MouseMe As DateTimeStamp, MouseControls As DTSControls, MouseX As Single, MouseY As Single)

Dim MouseMode As Integer

MouseMode = MouseX - MouseControls.ControlTimeCellleft
MouseMode = MouseMode \ MouseControls.ControlTimeCellwidth

Select Case MouseMode
    Case 0, 1
        Call MouseGridChange(MouseMe, MouseControls, MouseX, MouseY, MouseControls.ControlTimeCellleft, MouseControls.ControlTimeCellwidth, MouseControls.ControlTimeCellheight, "H")
        Call DrawForLimitCells(MouseMe, MouseControls, 4)
    Case 2, 3
        Call MouseGridChange(MouseMe, MouseControls, MouseX, MouseY, MouseControls.ControlTimeCellleft + MouseControls.ControlTimeCellwidth * 2 + 15, MouseControls.ControlTimeCellwidth, MouseControls.ControlTimeCellheight, "M")
        Call DrawForLimitCells(MouseMe, MouseControls, 5)
    Case 4, 5
        Call MouseGridChange(MouseMe, MouseControls, MouseX, MouseY, MouseControls.ControlTimeCellleft + MouseControls.ControlTimeCellwidth * 4 + 30, MouseControls.ControlTimeCellwidth, MouseControls.ControlTimeCellheight, "S")
        Call DrawForLimitCells(MouseMe, MouseControls, 6)
End Select

Call RedrawBorders(MouseMe, MouseControls)

End Sub

Private Sub MouseGridChange(MouseMe As DateTimeStamp, MouseControls As DTSControls, MouseX As Single, MouseY As Single, RedrawLeft As Long, RedrawCellwidth As Long, RedrawCellheight As Long, RedrawMode As String)

Dim MouseCell As Long
Dim MouseMaximum As Integer
Dim MousePosition As Integer
Dim MouseDatevalue As DTSValuetext
Dim MouseLimit As Integer
Dim MouseMode As Integer
Dim MouseText As String

If RedrawMode <> "Y" Or GetYearMode(MouseControls, 2) = 0 Then
        
    MouseCell = MouseY - MouseControls.ControlHeaderCellheight - RedrawCellheight

    If MouseCell < 0 Then Exit Sub

    MouseCell = MouseCell \ RedrawCellheight

    If MouseCell > 5 Then Exit Sub

    MouseCell = MouseCell * 2

    If MouseX > RedrawLeft + RedrawCellwidth Then MouseCell = MouseCell + 1

    Select Case RedrawMode

        Case "Y"
    
            If MouseControls.ControlClickMode = "Y" Then
    
                MouseCell = (MouseControls.ControlCurrentValuetext.ValuetextYear \ 10) * 10 + MouseCell
            
                If ValidCurrentChangeLimits(MouseControls, 34, MouseCell) = False Then Exit Sub
            
                MouseControls.ControlCurrentValuetext.ValuetextYear = MouseCell

                Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)

                Call SwapClickMode(MouseMe, MouseControls, "=")
                Call VerifyMonthChange(MouseMe, MouseControls)
        
            ElseIf (MouseCell = 0 And MouseControls.ControlYearOffset > 100 And ValidCurrentChangeLimits(MouseControls, 33, MouseControls.ControlYearOffset)) Or (MouseCell = 9 And MouseControls.ControlYearOffset < 290 And ValidCurrentChangeLimits(MouseControls, 33, MouseControls.ControlYearOffset + 9)) Then
            
                If MouseCell = 0 Then
                    MouseControls.ControlYearOffset = MouseControls.ControlYearOffset - 8
                    If MouseControls.ControlYearOffset < 100 Then MouseControls.ControlYearOffset = 100
                Else
                    MouseControls.ControlYearOffset = MouseControls.ControlYearOffset + 8
                    If MouseControls.ControlYearOffset > 290 Then MouseControls.ControlYearOffset = 290
                End If
            
                Call SwapClickMode(MouseMe, MouseControls, "Y")
            
                MouseControls.ControlClickMode = "="
            
                Call DrawGridCells(MouseMe, MouseControls, MouseControls.ControlYearCellleft, MouseControls.ControlYearCellwidth, MouseControls.ControlYearCellheight, "Y")
        
                Call RedrawBorders(MouseMe, MouseControls)
        
            Else
            
                MouseCell = MouseControls.ControlYearOffset + MouseCell
            
                If ValidCurrentChangeLimits(MouseControls, 33, MouseCell) = False Then Exit Sub
            
                MouseCell = MouseCell * 10
            
                MouseDatevalue = MouseControls.ControlCurrentValuetext
                
                MouseDatevalue.ValuetextYear = MouseCell
                
                Call VerifyValuedateLimits(MouseControls, MouseDatevalue)
                
                MouseControls.ControlCurrentValuetext.ValuetextYear = MouseDatevalue.ValuetextYear
    
                Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)
            
                Call SwapClickMode(MouseMe, MouseControls, "Y")
                Call DrawGridCells(MouseMe, MouseControls, MouseControls.ControlYearCellleft, MouseControls.ControlYearCellwidth, MouseControls.ControlYearCellheight, "Y")
                Call VerifyMonthChange(MouseMe, MouseControls)

            End If

        Case "E"
    
            MousePosition = Val(MouseControls.ControlClickMode) + 1
            MouseLimit = 76 - (MouseControls.ControlCronoDecimals - MousePosition)
            MouseText = Right$(String$(MouseControls.ControlCronoDecimals, 48) + Trim$(Str$(MouseControls.ControlCurrentValuetext.ValuetextCrono)), MouseControls.ControlCronoDecimals)
            MouseText = Left$(MouseText, MousePosition - 1) + Chr$(48 + MouseCell) + String$(Len(MouseText) - MousePosition, 48)
            MouseCell = Val(Left$(MouseText, MousePosition))
        
            If ValidCurrentChangeLimits(MouseControls, MouseLimit, MouseCell) = False Then Exit Sub
        
            MouseControls.ControlCurrentValuetext.ValuetextCrono = Val(MouseText)
        
            Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)
        
            If MousePosition < MouseControls.ControlCronoDecimals Then
                Call SwapClickMode(MouseMe, MouseControls, Chr$(48 + MousePosition))
            Else
                Call SwapClickMode(MouseMe, MouseControls, "=")
            End If
        
            Call DrawCronoCell(MouseMe, MouseControls)
            Call DrawGridCells(MouseMe, MouseControls, MouseControls.ControlCronoCellleft, MouseControls.ControlCronoCellwidth, MouseControls.ControlCronoCellheight, "E")

            Call RedrawBorders(MouseMe, MouseControls)
    
            MouseMe.RaiseDateTimeChange

        Case Else
    
            Select Case RedrawMode
                Case "H"
                    MouseLimit = 42
                    MouseMaximum = 23
                    MouseMode = 0
                Case "M"
                    MouseLimit = 52
                    MouseMaximum = 59
                    MouseMode = 1
                Case "S"
                    MouseLimit = 62
                    MouseMaximum = 59
                    MouseMode = 2
                Case Else
                    Exit Sub
            End Select
    
            If MouseControls.ControlClickMode = RedrawMode Then
    
                MouseCell = (MouseControls.ControlCurrentValuetext.ValuetextTime(MouseMode) \ 10) * 10 + MouseCell
        
                If MouseCell > MouseMaximum Then Exit Sub
                If ValidCurrentChangeLimits(MouseControls, MouseLimit, MouseCell) = False Then Exit Sub
            
                MouseControls.ControlCurrentValuetext.ValuetextTime(MouseMode) = MouseCell
        
                Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)
                
                Call DrawTimeChange(MouseMe, MouseControls, MouseMode)
                Call SwapClickMode(MouseMe, MouseControls, "=")

            Else
    
                MouseLimit = MouseLimit - 1
            
                If ValidCurrentChangeLimits(MouseControls, MouseLimit, MouseCell) = False Then Exit Sub
            
                MouseCell = MouseCell * 10
            
                If MouseCell > MouseMaximum Then Exit Sub
            
                MouseDatevalue = MouseControls.ControlCurrentValuetext
                
                MouseDatevalue.ValuetextTime(MouseMode) = MouseCell
                
                Call VerifyValuedateLimits(MouseControls, MouseDatevalue)
                
                MouseControls.ControlCurrentValuetext.ValuetextTime(MouseMode) = MouseDatevalue.ValuetextTime(MouseMode)
    
                Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)
            
                Call DrawTimeChange(MouseMe, MouseControls, MouseMode)
                Call SwapClickMode(MouseMe, MouseControls, RedrawMode)
    
            End If
    
            Call DrawGridCells(MouseMe, MouseControls, RedrawLeft, RedrawCellwidth, RedrawCellheight, RedrawMode)
        
            Call RedrawBorders(MouseMe, MouseControls)
    
            MouseMe.RaiseDateTimeChange

    End Select

Else

    MouseCell = MouseY - MouseControls.ControlHeaderCellheight

    If MouseCell < 0 Then Exit Sub

    MouseCell = MouseCell \ RedrawCellheight
    
    If MouseControls.ControlYearEnd - MouseControls.ControlYearStart > 5 Then
    
        MouseCell = MouseCell * 2

        If MouseX > RedrawLeft + RedrawCellwidth Then MouseCell = MouseCell + 1
    
    End If
    
    MouseCell = MouseCell + MouseControls.ControlYearStart
    
    If MouseCell > MouseControls.ControlYearEnd Then Exit Sub
    
    MouseControls.ControlCurrentValuetext.ValuetextYear = MouseCell

    Call VerifyValuedateLimits(MouseControls, MouseControls.ControlCurrentValuetext)

    Call SwapClickMode(MouseMe, MouseControls, "=")
    Call VerifyMonthChange(MouseMe, MouseControls)
    
End If

End Sub

Private Function DatePartFormat(PartValue As Integer)

DatePartFormat = Right$(Str$(100 + PartValue), 2)

End Function

Private Sub SetYearOffset(SetControls As DTSControls)

SetControls.ControlYearOffset = SetControls.ControlCurrentValuetext.ValuetextYear \ 10 - 4

If SetControls.ControlYearOffset < 100 Then SetControls.ControlYearOffset = 100
If SetControls.ControlYearOffset > 290 Then SetControls.ControlYearOffset = 290

End Sub

Private Function ReturnLimitColor(ReturnControls As DTSControls, ReturnMode As Integer, ReturnCheck As Long) As OLE_COLOR

If ValidCurrentChangeLimits(ReturnControls, ReturnMode, ReturnCheck) = False Then
    ReturnLimitColor = ReturnControls.ControlDisabledBackground
Else
    ReturnLimitColor = ReturnControls.ControlNormalBackground
End If

End Function

Private Sub VerifyValuedateLimits(VerifyControls As DTSControls, VerifyTestdate As DTSValuetext)

If ValidValuedateLimit(VerifyControls, 88, 8888, VerifyTestdate, VerifyControls.ControlMinimumValuetext, 1) = False Then VerifyTestdate = VerifyControls.ControlMinimumValuetext
If ValidValuedateLimit(VerifyControls, 88, 8888, VerifyTestdate, VerifyControls.ControlMaximumValuetext, 2) = False Then VerifyTestdate = VerifyControls.ControlMaximumValuetext

End Sub

Private Function ValidCurrentChangeLimits(VerifyControls As DTSControls, VerifyMode As Integer, VerifyCheck As Long) As Boolean

Dim VerifyResult As Boolean

    VerifyResult = False
    
    If ValidValuedateLimit(VerifyControls, VerifyMode, VerifyCheck, VerifyControls.ControlCurrentValuetext, VerifyControls.ControlMinimumValuetext, 1) = False Then GoTo endvalidcurrentchangelimits
    If ValidValuedateLimit(VerifyControls, VerifyMode, VerifyCheck, VerifyControls.ControlCurrentValuetext, VerifyControls.ControlMaximumValuetext, 2) = False Then GoTo endvalidcurrentchangelimits
    
    VerifyResult = True

endvalidcurrentchangelimits:

    ValidCurrentChangeLimits = VerifyResult

End Function

Private Function ValidValuedateLimit(ReturnControls As DTSControls, ReturnMode As Integer, ReturnCheck As Long, ReturnTestdate As DTSValuetext, ReturnValuedate As DTSValuetext, ReturnCondition As Integer) As Boolean

Dim ReturnResult As Long
Dim ReturnSubcheck As Long
Dim ReturnSubmode As Integer
Dim ReturnFlag As Boolean

    Select Case ReturnMode
        Case 33, 41, 51, 61, 75
            ReturnSubmode = ReturnMode + 1
            ReturnSubcheck = ReturnCheck * 10
            If ReturnCondition = 1 Then ReturnSubcheck = ReturnSubcheck + 9
        Case 74
            ReturnSubmode = 76
            ReturnSubcheck = ReturnCheck * 100
            If ReturnCondition = 1 Then ReturnSubcheck = ReturnSubcheck + 99
        Case 73
            ReturnSubmode = 76
            ReturnSubcheck = ReturnCheck * 1000
            If ReturnCondition = 1 Then ReturnSubcheck = ReturnSubcheck + 999
        Case 72
            ReturnSubmode = 76
            ReturnSubcheck = ReturnCheck * 10000
            If ReturnCondition = 1 Then ReturnSubcheck = ReturnSubcheck + 9999
        Case 71
            ReturnSubmode = 76
            ReturnSubcheck = ReturnCheck * 100000
            If ReturnCondition = 1 Then ReturnSubcheck = ReturnSubcheck + 99999
        Case Else
            ReturnSubcheck = ReturnCheck
            ReturnSubmode = ReturnMode
    End Select
    
    ReturnResult = 0
    ReturnFlag = False

    If ReturnSubmode = 34 Then
        
        ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextYear

    Else

        If ReturnControls.ControlDayCellwidth > 0 Or ReturnControls.ControlYearCellwidth > 0 Then
            ReturnResult = ReturnTestdate.ValuetextYear - ReturnValuedate.ValuetextYear
            If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit
        End If

        If ReturnSubmode = 22 Then

            ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextMonth

        Else

            If ReturnControls.ControlDayCellwidth > 0 Or ReturnControls.ControlMonthCellwidth > 0 Then
                ReturnResult = ReturnTestdate.ValuetextMonth - ReturnValuedate.ValuetextMonth
                If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit
            End If

            If ReturnSubmode = 12 Then

                ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextDay

            Else

                If ReturnControls.ControlDayCellwidth > 0 Then
                    ReturnResult = ReturnTestdate.ValuetextDay - ReturnValuedate.ValuetextDay
                    If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit
                End If

                If ReturnSubmode = 42 Then

                    ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextTime(0)
                
                Else

                    If ReturnControls.ControlTimeCellwidth > 0 Then

                        ReturnResult = ReturnTestdate.ValuetextTime(0) - ReturnValuedate.ValuetextTime(0)
                        If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit

                        If ReturnSubmode = 52 Then

                            ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextTime(1)

                        Else

                            ReturnResult = ReturnTestdate.ValuetextTime(1) - ReturnValuedate.ValuetextTime(1)
                            If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit

                            If ReturnSubmode = 62 Then

                                ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextTime(2)

                            Else

                                If Len(ReturnControls.ControlDateTimeFormat) > 1 Then
                                    If Left$(ReturnControls.ControlDateTimeFormat, Len(ReturnControls.ControlDateTimeFormat) - 1) = "R" Then ReturnFlag = True
                                End If

                                If ReturnFlag = False Then
                                    ReturnResult = ReturnTestdate.ValuetextTime(2) - ReturnValuedate.ValuetextTime(2)
                                    If ReturnResult <> 0 Then GoTo endvalidvaluedatelimit
                                End If

                                If ReturnSubmode = 76 Then

                                    ReturnResult = ReturnSubcheck - ReturnValuedate.ValuetextCrono

                                Else

                                    If ReturnControls.ControlCronoCellwidth > 0 Then
                                        ReturnResult = ReturnTestdate.ValuetextCrono - ReturnValuedate.ValuetextCrono
                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End If

    End If

endvalidvaluedatelimit:

    ReturnFlag = True
    
    If ReturnCondition = 1 Then
        If ReturnResult < 0 Then ReturnFlag = False
    Else
        If ReturnResult > 0 Then ReturnFlag = False
    End If

    ValidValuedateLimit = ReturnFlag
    
End Function

Private Function DoGetDateTimeControldata(GetControls As DTSControls, GetControldata As DTSValuetext)

Dim GetResult As String
Dim GetReduced As Boolean

GetResult = ""
GetReduced = False

If GetControls.ControlDayCellwidth > 0 Then
    GetResult = GetResult + DatePartFormat(GetControldata.ValuetextDay) + DTSDateSeparator + DatePartFormat(GetControldata.ValuetextMonth) + DTSDateSeparator
Else
    If GetControls.ControlMonthCellwidth > 0 Then GetResult = DatePartFormat(GetControldata.ValuetextMonth) + DTSDateSeparator
End If

If GetControls.ControlDayCellwidth > 0 Or GetControls.ControlYearCellwidth > 0 Then GetResult = GetResult + Right$(Str$(10000 + GetControldata.ValuetextYear), 4)

If GetControls.ControlTimeCellwidth > 0 Then
   
   If GetResult <> "" Then GetResult = GetResult + " "
   
   GetResult = GetResult + DatePartFormat(GetControldata.ValuetextTime(0)) + DTSHourSeparator + DatePartFormat(GetControldata.ValuetextTime(1))

    If Len(GetControls.ControlDateTimeFormat) > 1 Then
        If Left$(GetControls.ControlDateTimeFormat, Len(GetControls.ControlDateTimeFormat) - 1) = "R" Then GetReduced = True
    End If

    If GetReduced = False Then GetResult = GetResult + DTSSecondSeparator + DatePartFormat(GetControldata.ValuetextTime(2))

End If

If GetControls.ControlCronoCellwidth > 0 Then GetResult = GetResult + DTSTimingSeparator + Right$(String$(GetControls.ControlCronoDecimals, 48) + Trim$(Str$(GetControldata.ValuetextCrono)), GetControls.ControlCronoDecimals)

DoGetDateTimeControldata = GetResult

End Function

Private Sub DoLetNewControldata(NewControls As DTSControls, NewControldata As DTSValuetext, NewValuetext As String, NewDefaultmode As Integer)

Dim LetReduced As Boolean

If NewControls.ControlDayCellwidth > 0 Then Call GetValueTextvalue(NewValuetext, DTSDateSeparator, NewControldata.ValuetextDay, 1, 31, NewDefaultmode)
If NewControls.ControlMonthCellwidth > 0 Or NewControls.ControlDateTimeFormat = "FD" Then Call GetValueTextvalue(NewValuetext, DTSDateSeparator, NewControldata.ValuetextMonth, 1, 12, NewDefaultmode)
If NewControls.ControlYearCellwidth > 0 Or NewControls.ControlDateTimeFormat = "FD" Then Call GetValueTextvalue(NewValuetext, " ", NewControldata.ValuetextYear, 1000, 2999, NewDefaultmode)

If NewControls.ControlTimeCellwidth > 0 Then

    Call GetValueTextvalue(NewValuetext, DTSHourSeparator, NewControldata.ValuetextTime(0), 0, 23, NewDefaultmode)
    Call GetValueTextvalue(NewValuetext, DTSSecondSeparator, NewControldata.ValuetextTime(1), 0, 59, NewDefaultmode)

    LetReduced = False

    If Len(NewControls.ControlDateTimeFormat) > 1 Then
        If Left$(NewControls.ControlDateTimeFormat, Len(NewControls.ControlDateTimeFormat) - 1) = "R" Then LetReduced = True
    End If

    If LetReduced = False Then Call GetValueTextvalue(NewValuetext, DTSTimingSeparator, NewControldata.ValuetextTime(2), 0, 59, NewDefaultmode)

End If

If NewControls.ControlCronoCellwidth > 0 Then
    
    If NewValuetext <> "" Then
        NewValuetext = Left$(Trim$(NewValuetext) + String$(NewControls.ControlCronoDecimals, 48), NewControls.ControlCronoDecimals)
    ElseIf NewDefaultmode = 0 Then
        NewValuetext = String(NewControls.ControlCronoDecimals, 48)
    Else
        NewValuetext = String(NewControls.ControlCronoDecimals, 57)
    End If
    
    NewControldata.ValuetextCrono = Val(NewValuetext)

End If

End Sub

Private Sub GetValueTextvalue(NewValuetext As String, NewSeparator As String, ReturnResult As Integer, ReturnMinimum As Integer, ReturnMaximum As Integer, ReturnDefaultmode As Integer)

Dim GetPosition As Integer
Dim GetValue As Long

GetPosition = InStr(NewValuetext + NewSeparator, NewSeparator)

If GetPosition = 1 Then
    GetValue = ReturnDefaultmode
Else
    GetValue = Val(Left$(NewValuetext, GetPosition - 1))
    NewValuetext = Mid$(NewValuetext, GetPosition + 1)
End If

If ReturnMinimum = 1000 And ReturnMaximum = 2999 Then
    Select Case GetValue
        Case 0 To 49:  GetValue = GetValue + 2000
        Case 50 To 99: GetValue = GetValue + 1900
    End Select
End If

If GetValue > ReturnMaximum Then GetValue = ReturnMaximum
If GetValue < ReturnMinimum Then GetValue = ReturnMinimum

ReturnResult = GetValue

End Sub

Private Sub ResizeAndRedraw(UpdateMe As DateTimeStamp, UpdateControls As DTSControls)

UpdateControls.ControlClickMode = "="
    
Call ResizeCalendar(UpdateMe, UpdateControls)
    
Call RedrawAllHeaders(UpdateMe, UpdateControls)
Call RedrawAllCells(UpdateMe, UpdateControls)
Call RedrawBorders(UpdateMe, UpdateControls)
    
End Sub

Private Function GetYearMode(GetControls As DTSControls, GetReturnMode As Integer) As Integer

If GetControls.ControlYearStart = 0 Or GetControls.ControlYearEnd = 0 Then
    GetYearMode = 0
ElseIf GetControls.ControlYearEnd - GetControls.ControlYearStart > 5 Then
    GetYearMode = GetReturnMode
Else
    GetYearMode = 1
End If

End Function

