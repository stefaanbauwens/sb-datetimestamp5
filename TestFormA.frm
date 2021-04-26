VERSION 5.00
Object = "*\AControlProject.vbp"
Begin VB.Form TestFormA 
   Caption         =   "DateTimeStamp TestProgram"
   ClientHeight    =   4935
   ClientLeft      =   1020
   ClientTop       =   1740
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8895
   Begin VB.CommandButton Command1 
      Caption         =   "Today"
      Height          =   315
      Left            =   6000
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Reverse"
      Height          =   315
      Left            =   6870
      TabIndex        =   27
      Top             =   840
      Width           =   945
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Flat"
      Height          =   315
      Left            =   7200
      TabIndex        =   26
      Top             =   480
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   13
      Left            =   6360
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   12
      Left            =   6000
      TabIndex        =   24
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   11
      Left            =   5400
      TabIndex        =   23
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   10
      Left            =   5040
      TabIndex        =   22
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   9
      Left            =   4560
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   8
      Left            =   4200
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   7
      Left            =   3720
      TabIndex        =   19
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   6
      Left            =   3360
      TabIndex        =   18
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   5
      Left            =   2760
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   4
      Left            =   2400
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bScroll 
      Caption         =   "-"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin sbdts5.DateTimeStamp stCalendar1 
      Height          =   3615
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6376
      DateTimeValuetext=   "01/01/1950 00:00:00.000000"
      DateTimeFormat  =   "6T"
      FlatMode        =   -1  'True
      BeginProperty DateTimeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   315
      Left            =   7920
      Max             =   2
      TabIndex        =   10
      Top             =   840
      Value           =   2
      Width           =   855
   End
   Begin VB.ComboBox cmbTime3 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   750
   End
   Begin VB.ComboBox cmbTime2 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   750
   End
   Begin VB.ComboBox cmbTime1 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   750
   End
   Begin VB.ComboBox cmbDay 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   750
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   315
      Left            =   7920
      Max             =   15
      Min             =   1
      TabIndex        =   3
      Top             =   120
      Value           =   1
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   7920
      Max             =   6
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "TestFormA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim UpdateX As Long
Dim UpdateY As Long
Dim UpdateButton As Long

Option Explicit

Private Sub Check1_Click()

If Check1.Value = 0 Then
    stCalendar1.FlatMode = False
Else
    stCalendar1.FlatMode = True
End If

End Sub

Private Sub Check2_Click()

If Check2.Value = 0 Then
    stCalendar1.ReverseScroll = False
Else
    stCalendar1.ReverseScroll = True
End If

End Sub

Private Sub bScroll_Click(ClickIndex As Integer)

Select Case ClickIndex
    Case 0:  Call stCalendar1.ScrollDateTime(-1, 0, 0, 0, 0, 0, 0)
    Case 1:  Call stCalendar1.ScrollDateTime(1, 0, 0, 0, 0, 0, 0)
    Case 2:  Call stCalendar1.ScrollDateTime(0, -1, 0, 0, 0, 0, 0)
    Case 3:  Call stCalendar1.ScrollDateTime(0, 1, 0, 0, 0, 0, 0)
    Case 4:  Call stCalendar1.ScrollDateTime(0, 0, -1, 0, 0, 0, 0)
    Case 5:  Call stCalendar1.ScrollDateTime(0, 0, 1, 0, 0, 0, 0)
    Case 6:  Call stCalendar1.ScrollDateTime(0, 0, 0, -1, 0, 0, 0)
    Case 7:  Call stCalendar1.ScrollDateTime(0, 0, 0, 1, 0, 0, 0)
    Case 8:  Call stCalendar1.ScrollDateTime(0, 0, 0, 0, -1, 0, 0)
    Case 9:  Call stCalendar1.ScrollDateTime(0, 0, 0, 0, 1, 0, 0)
    Case 10: Call stCalendar1.ScrollDateTime(0, 0, 0, 0, 0, -1, 0)
    Case 11: Call stCalendar1.ScrollDateTime(0, 0, 0, 0, 0, 1, 0)
    Case 12: Call stCalendar1.ScrollDateTime(0, 0, 0, 0, 0, 0, -1)
    Case 13: Call stCalendar1.ScrollDateTime(0, 0, 0, 0, 0, 0, 1)
End Select

End Sub

Private Sub Command1_Click()

stCalendar1.SetToday

End Sub

Private Sub Form_Load()

Dim LoadIndex As Long
Dim LoadYear As Long

For LoadIndex = 101 To 131
    cmbDay.AddItem Right$(Str$(LoadIndex), 2)
Next LoadIndex
cmbDay.ListIndex = Day(Now) - 1

For LoadIndex = 1 To 12
    cmbMonth.AddItem Format(CDate("1/" & LoadIndex & "/1999"), "mmmm")
Next LoadIndex
cmbMonth.ListIndex = Month(Now) - 1

LoadYear = Val(Format(Now, "yyyy"))

For LoadIndex = LoadYear - 6 To LoadYear + 6
    cmbYear.AddItem Str(LoadIndex)
Next LoadIndex
cmbYear.ListIndex = 6

For LoadIndex = 100 To 123
    cmbTime1.AddItem Right$(Str$(LoadIndex), 2)
Next LoadIndex
cmbTime1.ListIndex = Hour(Now)

For LoadIndex = 100 To 159
    cmbTime2.AddItem Right$(Str$(LoadIndex), 2)
Next LoadIndex
cmbTime2.ListIndex = Minute(Now)

For LoadIndex = 100 To 159
    cmbTime3.AddItem Right$(Str$(LoadIndex), 2)
Next LoadIndex
cmbTime3.ListIndex = Second(Now)

HScroll2.Value = 12
HScroll3.Value = 2

Call UpdateDateTime

Load TestFormB

End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub

If Me.WindowState = 0 Then
    If Me.Width < 3090 Then Me.Width = 3090
    If Me.Height < 3090 Then Me.Height = 3090
End If

stCalendar1.Width = Me.ScaleWidth - stCalendar1.Left - 120
stCalendar1.Height = Me.ScaleHeight - stCalendar1.Top - 120

End Sub

Private Sub HScroll1_Change()

stCalendar1.StartingWeekday = HScroll1.Value

End Sub

Private Sub HScroll2_Change()

Select Case HScroll2.Value
    Case 1:  Label1.Caption = "Y"
    Case 2:  Label1.Caption = "M"
    Case 3:  Label1.Caption = "D"
    Case 4:  Label1.Caption = "FD"
    Case 5:  Label1.Caption = "RT"
    Case 6:  Label1.Caption = "T"
    Case 7:  Label1.Caption = "1T"
    Case 8:  Label1.Caption = "2T"
    Case 9:  Label1.Caption = "3T"
    Case 10: Label1.Caption = "4T"
    Case 11: Label1.Caption = "5T"
    Case 12: Label1.Caption = "6T"
    Case 13: Label1.Caption = "6H"
    Case 14: Label1.Caption = "H"
    Case 15: Label1.Caption = "RH"
End Select

stCalendar1.DateTimeFormat = Label1.Caption

Call UpdateDateTime

End Sub

Private Sub HScroll3_Change()

stCalendar1.HeaderMode = HScroll3.Value

End Sub

Private Sub cmbDay_Click()

Call UpdateDateTime

End Sub

Private Sub cmbMonth_Click()

Call UpdateDateTime

End Sub

Private Sub cmbYear_Click()

Call UpdateDateTime

End Sub

Private Sub cmbTime1_Click()

Call UpdateDateTime

End Sub

Private Sub cmbTime2_Click()

Call UpdateDateTime

End Sub

Private Sub cmbTime3_Click()

Call UpdateDateTime

End Sub

Private Sub UpdateDateTime()

Dim SetDate As String
Dim SetYear As String
Dim SetReduced As String
Dim SetSecond As String
Dim SetCrono As String
Dim SetText As String

SetYear = Trim$(cmbYear.List(cmbYear.ListIndex))
SetDate = Right$(Str$(cmbDay.ListIndex + 101), 2) + "/" + Right$(Str$(cmbMonth.ListIndex + 101), 2) + "/" + SetYear
SetReduced = Right$(Str$(cmbTime1.ListIndex + 100), 2) + ":" + Right$(Str$(cmbTime2.ListIndex + 100), 2)
SetSecond = ":" + Right$(Str$(cmbTime3.ListIndex + 100), 2)
SetCrono = "."

Randomize Timer
While Len(SetCrono) < 8
    SetCrono = SetCrono + Chr$(48 + Int(Rnd(Timer) * 9.9))
Wend

Select Case Label1.Caption
    Case "Y":  SetText = SetYear
    Case "M":  SetText = Mid$(SetDate, 4)
    Case "D":  SetText = SetDate
    Case "RT": SetText = SetDate + " " + SetReduced
    Case "T":  SetText = SetDate + " " + SetReduced + SetSecond
    Case "6T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 7)
    Case "5T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 6)
    Case "4T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 5)
    Case "3T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 4)
    Case "2T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 3)
    Case "1T": SetText = SetDate + " " + SetReduced + SetSecond + Left$(SetCrono, 2)
    Case "RH": SetText = SetReduced
    Case "H":  SetText = SetReduced + SetSecond
    Case "6H": SetText = SetReduced + SetSecond + Left$(SetCrono, 7)
    Case Else: Exit Sub
End Select

stCalendar1.DateTimeValuetext = SetText

End Sub

Private Sub stCalendar1_DateTimeChange()

Label2.Caption = stCalendar1.DateTimeValuetext

End Sub

Private Sub stCalendar1_DateTimeKeyPress(KeyAscii As Integer)

Debug.Print "DateTimeKeyPress " & KeyAscii

End Sub

Private Sub stCalendar1_DateTimeMouseUp(MouseButton As Integer)

Debug.Print "DateTimeMouseUp " & MouseButton

End Sub

Private Sub stCalendar1_HoverChange(IsHovered As Boolean)

Debug.Print "HoverChange " & IsHovered

End Sub

Private Sub stCalendar1_MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

Debug.Print "MouseMove " & MouseState & " " & MouseX & " " & MouseY

End Sub

