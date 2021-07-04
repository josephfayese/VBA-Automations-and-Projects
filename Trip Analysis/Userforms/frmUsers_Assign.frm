VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsers_Assign 
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885.001
   OleObjectBlob   =   "frmUsers_Assign.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUsers_Assign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub UserForm_Initialize()
'DatePickerX_Ini'-------------< this line goes to your UserForm_Initialize
'End Sub

'show----------------------------
'Private Sub TextBox1_Enter()
'PX_hide
'End Sub

'hide----------------------------
'Private Sub TextBox2_Enter()
'GetDate
'End Sub

'Private Sub Image1_MouseDown(******
'TextBox5.SetFocus
'GetDate
'End Sub

Dim DatePickerX_Ctrls() As cDatePickerX

'Private Sub UserForm_Initialize()
'DatePickerX_Ini
'End Sub

Sub PX_hide()

On Error Resume Next
DatePickerX.Visible = False
On Error GoTo 0

End Sub
Function GetDate()
'date picker loader

Dim k As Control

'---------------------------------------------prep notes
'If DatePicker is located in MultiPage then update MultiPage1 to name of your MultiPage accordingly
'activeCtrl = Me.MultiPage1(Me.MultiPage1.Value).ActiveControl.Name
'else use below if located in UserForm itself.
'activeCtrl = Me.ActiveControl.Name

'activeCtrl = Me.MultiPage1(Me.MultiPage1.Value).ActiveControl.Name
activeCtrl = Me.ActiveControl.Name
'---------------------------------------------

Set k = Me.Controls(activeCtrl)

With Me.DatePickerX
   .Left = k.Left + k.Width + 10
   .Top = (k.Top + k.Height) - k.Height
   
   If Me.Height < (k.Top + .Height) Then
         .Left = (k.Left + k.Width) + 2
         .Top = (k.Top - .Height) + k.Height
   End If
   .Visible = True
End With
Set k = Nothing

End Function


Function DatePickerX_PrevNext(showNxt As Boolean)
Dim tmpDate As Date, vNewMonthDate As Date
   tmpDate = DateSerial(Me.mem_year.Value, Me.mem_mth.Value, 1)
   If showNxt = True Then
   vNewMonthDate = DateAdd("m", 1, tmpDate)
   Else
   vNewMonthDate = DateAdd("m", -1, tmpDate)
   End If
   Call LoadDates(Month(vNewMonthDate), Year(vNewMonthDate))
End Function

Private Sub cmbRole_Change()
PX_hide
End Sub

Private Sub cmbRole_Enter()
PX_hide
End Sub

Private Sub CommandButton1_Click()
    Call mMain.Create_Users
    'Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    Call mMain.User_Activity
End Sub

Private Sub CommandButton4_Click()
    Call mMain.Create_Trip_Report
End Sub

Private Sub CommandButton5_Click()
frmFareRate.Show
End Sub

Private Sub eCalNextMonth_Click()
   DatePickerX_PrevNext True
End Sub

Private Sub eCalPrevMonth_Click()
   DatePickerX_PrevNext False
End Sub

Private Sub eCalTitle_Click()

   tbYear.Visible = True
   mthsCB.Visible = True
   mthsCB.Height = 133.9

End Sub

Private Sub eCalToday_Click()
   Call LoadDates(Month(Date), Year(Date))
End Sub



Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
txtDate_Assigned.SetFocus
GetDate
End Sub

Private Sub mthsCB_Click()

   Me.eCalTitle.Caption = Me.mthsCB.Value & " " & Me.tbYear.Value
   tbYear.Visible = False
   mthsCB.Visible = False
   
   Call LoadDates(mthnobytext(mthsCB.Value), tbYear.Value)

End Sub

Private Sub tbYear_Change()
   If Len(tbYear.Value) >= 4 Then
      tbYear.Value = Left(tbYear.Value, 4)
   End If
End Sub

Private Sub tbYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   Select Case KeyAscii
      Case 48 To 57
         If Len(tbYear) >= 4 Then
            KeyAscii = 0
         End If
      Case Else
          KeyAscii = 0
   End Select
End Sub


Sub DatePickerX_Ini()
Dim Obj As Object, CtrlPointer As Long
'------------------------------------------------------
Me.tbYear.Value = Year(Date)
With mthsCB
   .Clear
   .AddItem "January"
   .AddItem "February"
   .AddItem "March"
   .AddItem "April"
   .AddItem "May"
   .AddItem "June"
   .AddItem "July"
   .AddItem "August"
   .AddItem "September"
   .AddItem "October"
   .AddItem "November"
   .AddItem "December"
End With
'------------------------------------------------------
Me.DatePickerX.Visible = False
ActiveUF = Me.Name
Call LoadDates(Month(Date), Year(Date))
DatePickerX.BackColor = DatePickerX_Back
Me.eCalTitle.ForeColor = DatePickerX_Title_Font
'------------------------------------------------------
ReDim DatePickerX_Ctrls(1 To Me.Controls.Count)
For Each Obj In Me.Controls
    If TypeName(Obj) = "Label" And (Obj.Tag = "daysbg" Or Obj.Tag = "days") Then
        CtrlPointer = CtrlPointer + 1
        Set DatePickerX_Ctrls(CtrlPointer) = New cDatePickerX
        Set DatePickerX_Ctrls(CtrlPointer).aMenu = Obj
    End If
Next Obj
ReDim Preserve DatePickerX_Ctrls(1 To CtrlPointer)
'------------------------------------------------------
End Sub
Function LoadDates(mth As Byte, yearX As Integer)
Dim nDate As Date, dayNo As String, lDate As Date, mthNo As Byte, yrNo As Byte, kDate As Date, i As Long, dayX As Long
'------------------------------------------------------
nDate = DateSerial(yearX, mth, 1)
dayNo = daybyNo(Format(nDate, "DDD"))

lDate = dhLastDayInMonth(nDate)
Me.eCalTitle.Caption = Format(nDate, "MMMM YYYY")

Me.mem_mth = Month(nDate)
Me.mem_year = Year(nDate)

Me.Controls("D" & 1).Caption = "S"
Me.Controls("D" & 2).Caption = "M"
Me.Controls("D" & 3).Caption = "T"
Me.Controls("D" & 4).Caption = "W"
Me.Controls("D" & 5).Caption = "T"
Me.Controls("D" & 6).Caption = "F"
Me.Controls("D" & 7).Caption = "S"

'------------------------------------------------------
dayX = 1

'reset
For i = 1 To 42
   Me.Controls("day" & i).ForeColor = Color_Dates_Font
   dayX = dayX + 1
Next i

dayX = 1

kDate = nDate

For i = dayNo To 42
   
   Me.Controls("day" & i).Caption = CInt(Format(kDate, "DD"))
   
   If kDate <> Date Then
      Me.Controls("s" & i).BackColor = Color_Dates_Back
   Else
      Me.Controls("s" & i).BackColor = Color_CDate_Backcolor
   End If
   
   Me.Controls("day" & i).ForeColor = Color_Dates_Font
   
   Me.Controls("s" & i).ControlTipText = kDate
   Me.Controls("day" & i).ControlTipText = kDate
   
   If kDate > lDate Then
      Me.Controls("day" & i).ForeColor = Color_ODates_Font
   End If
   
   dayX = dayX + 1
   kDate = kDate + 1
   
Next i
'------------------------------------------------------
'prior dates
kDate = nDate
If dayNo > 1 Then
For i = dayNo - 1 To 1 Step -1
   Me.Controls("day" & i).Caption = CInt(Format(kDate - 1, "DD"))
   
   Me.Controls("s" & i).ControlTipText = kDate - 1
   Me.Controls("day" & i).ControlTipText = kDate - 1
   Me.Controls("day" & i).ForeColor = Color_ODates_Font
   dayX = dayX + 1
   kDate = kDate - 1
Next i
End If

Dim m_d1 As Byte
m_d1 = Day(nDate)
'------------------------------------------------------
End Function


Private Sub txtDate_Assigned_Enter()
GetDate
End Sub

Private Sub txtEmail_Enter()
PX_hide
End Sub

Private Sub txtUserID_Enter()
PX_hide
End Sub

Private Sub txtUsername_Change()
If Me.txtUsername = "" Then
    Me.txtUserID = ""
Else
    Me.txtUserID = Me.txtUsername & "_01"
End If
End Sub

Private Sub txtUsername_Enter()
PX_hide
End Sub

Private Sub txtUsername_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtPWD.Value = mMain.Generate_Password(10)
End Sub

Private Sub UserForm_Activate()
    Me.cmbRole.AddItem "Analyst"
    Me.cmbRole.AddItem "Manager"
    Me.cmbRole.AddItem "Team Lead"
End Sub

Private Sub UserForm_Initialize()
DatePickerX_Ini
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
      If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

