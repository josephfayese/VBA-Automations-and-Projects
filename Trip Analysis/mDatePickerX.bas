Attribute VB_Name = "mDatePickerX"
'+------------------------------------------------------------+
'| VbaA2z - DatePickerX 1.0 | 10/4/2020                       |
'| Compatible with 32-Bit and 64-Bit Office                   |
'| Author: L. Pamai (VbaA2z.Team@gmail.com)                   |
'| Visit channel: Youtube.com/VbaA2z                          |
'| More download: VbaA2z.Blogspot.com                         |
'+------------------------------------------------------------+
'| Free for personal and commercial use.                      |
'| This code comes with no warrantees. Unsupported product.   |
'+------------------------------------------------------------+

Option Explicit

Public sDate As Date
Public activeCtrl As String
Public ActiveUF As String
Public ActiveCtr1 As String
Public ActiveCtr2 As String
Public Const DatePickerX_DateFormat = "mm/dd/yyyy" '"d/m/yyyy"

'-------------------COLOR CONTROL
   
   'color apply sample
   'DatePickerX_Back = RGB(Red,Green,Blue)
   'DatePickerX_Back = &H404040
   'DatePickerX_Back = vbBlack
   
   
   'Calendar
   Public Const DatePickerX_Back = &H404040
   Public Const DatePickerX_Font = &HC0C0C0
   Public Const DatePickerX_Title_Font = &H80FF80
   
   'normal control
   Public Const Color_Dates_Back = &HC0C0C0
   Public Const Color_Dates_Font = vbBlack
   
   'current date
   Public Const Color_CDate_Backcolor = &HC0C0FF
   Public Const Color_CDate_Font = 1
   
   'on hover
   Public Const Color_HoverColor_Back = &HE0E0E0
   Public Const Color_HoverColor_Font = 1 'NOT USED
   
   'dates falling outside the month
   Public Const Color_ODates_Font = &H808080 'vbGreen
'-------------------


Function tempUF() As Object
Dim ufr As Object
For Each ufr In UserForms
If ufr.Name = ActiveUF Then
      Set tempUF = ufr
  Exit For
End If
Next ufr
End Function
Function eCalCtrlRESET()
Dim tempUFXj As Object
Set tempUFXj = tempUF
Dim ctl As Control

For Each ctl In tempUFXj.Controls
If ctl.Tag = "daysbg" Or ctl.Tag = "days" Then
   If CDate(ctl.ControlTipText) <> Date Then
      ctl.BackColor = Color_Dates_Back
   Else
      ctl.BackColor = Color_CDate_Backcolor
   End If
End If
Next ctl
Set tempUFXj = Nothing
End Function
Sub SetDate()
Dim tempUFX As Object
Set tempUFX = tempUF
With tempUFX
   .Controls(activeCtrl).Value = Format(sDate, DatePickerX_DateFormat)
   .Controls(activeCtrl).SetFocus
   .Controls("DatePickerX").Visible = False
End With
End Sub
Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
If dtmDate = 0 Then
dtmDate = Date
End If
dhLastDayInMonth = DateSerial(Year(dtmDate), Month(dtmDate) + 1, 0)
End Function
Function daybyNo(dayX As String, Optional s As Boolean) As String
Select Case dayX
Case Is = "Mon"
daybyNo = IIf(s = True, 1, 2)
Case Is = "Tue"
daybyNo = IIf(s = True, 2, 3)
Case Is = "Wed"
daybyNo = IIf(s = True, 3, 4)
Case Is = "Thu"
daybyNo = IIf(s = True, 4, 5)
Case Is = "Fri"
daybyNo = IIf(s = True, 5, 6)
Case Is = "Sat"
daybyNo = IIf(s = True, 6, 7)
Case Is = "Sun"
daybyNo = IIf(s = True, 7, 1)
End Select
End Function
Function mthnobytext(mthX As String) As Byte
Select Case mthX
Case Is = "Jan", "January"
mthnobytext = 1
Case Is = "Feb", "February"
mthnobytext = 2
Case Is = "Mar", "March"
mthnobytext = 3
Case Is = "Apr", "April"
mthnobytext = 4
Case Is = "May", "May"
mthnobytext = 5
Case Is = "Jun", "June"
mthnobytext = 6
Case Is = "Jul", "July"
mthnobytext = 7
Case Is = "Aug", "August"
mthnobytext = 8
Case Is = "Sep", "September"
mthnobytext = 9
Case Is = "Oct", "October"
mthnobytext = 10
Case Is = "Nov", "November"
mthnobytext = 11
Case Is = "Dec", "December"
mthnobytext = 12
End Select
End Function
