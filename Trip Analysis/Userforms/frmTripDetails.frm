VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTripDetails 
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17205
   OleObjectBlob   =   "frmTripDetails.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTripDetails"
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

#If Win64 Then
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

#End If

Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16)
Private Const WS_SYSMENU As Long = &H80000
Private Const SW_SHOWMAXIMIZED = 3

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

Private Sub cmbDestination_AfterUpdate()
    On Error GoTo errH
    Me.cmbRoute.Value = Application.WorksheetFunction.VLookup(Me.cmbOrigin.Value & "|" & Me.cmbDestination.Value, Sheet3.Range("H2:N55"), 5, 0)
    Exit Sub
errH:
    MsgBox "You have entered a wrong entry"
    Exit Sub
End Sub

Private Sub cmbOrigin_AfterUpdate()
    On Error GoTo errhandl
    If Sheet3.Range("O2").Value <> "" Then
        Sheet3.Range("Destination_List").Clear
    End If
    Me.cmbDestination.Clear
    Sheet3.Range("O2").Value = Me.cmbOrigin.Value
    Call mTest.ReturnMultiples
    '=============
    'Load Destination List
    '==============
    Dim x2 As Range
    On Error GoTo errhandl
    For Each x2 In Sheet3.Range("Destination_List")
        Me.cmbDestination.AddItem x2
    Next x2
    Exit Sub
errhandl:
Sheet3.Range("O2").Value = ""
    Exit Sub
End Sub

Private Sub CommandButton1_Click()

Call mMain.Create_User_Report

End Sub


Private Sub cmbArrivedFrom_Enter()
Me.cmbArrivedFrom.BackColor = vbWhite
End Sub

Private Sub cmbBusCode_Change()
PX_hide
End Sub

Private Sub cmbBusCode_Enter()
    Me.cmbBusCode.BackColor = vbWhite
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub

Private Sub cmbCaptainsName_Enter()
Me.cmbCaptainsName.BackColor = vbWhite
End Sub

Private Sub cmbDestination_Enter()
Me.cmbDestination.BackColor = vbWhite
Call mTest.ReturnMultiples
End Sub

Private Sub cmbOrigin_Change()
    On Error GoTo errhandl
    If Sheet3.Range("O2").Value <> "" Then
        Sheet3.Range("Destination_List").Clear
    End If
    Me.cmbDestination.Clear
    Sheet3.Range("O2").Value = Me.cmbOrigin.Value
    Call mTest.ReturnMultiples
    '=============
    'Load Destination List
    '==============
    Dim x2 As Range
    On Error GoTo errhandl
    For Each x2 In Sheet3.Range("Destination_List")
        Me.cmbDestination.AddItem x2
    Next x2
    Exit Sub
errhandl:
Sheet3.Range("O2").Value = ""
Exit Sub
End Sub

Private Sub cmbOrigin_Enter()
Me.cmbOrigin.BackColor = vbWhite
End Sub

Private Sub cmbRoute_Enter()
Me.cmbRoute.BackColor = vbWhite
End Sub

Private Sub cmbSave_Click()
Call mMain.Data_Entry
End Sub

Private Sub cmbShift_Enter()
Me.cmbShift.BackColor = vbWhite
End Sub

Private Sub ComboBoxTripType_Enter()
Me.ComboBoxTripType.BackColor = vbWhite
End Sub

Private Sub CommandButtonViewEntries_Click()
    frmUserSummary.Show
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

Private Sub Image1_click()
txtDate.SetFocus
GetDate
End Sub

Private Sub Image2_click()
TimePicker.Show
Call TimePicker.Selected_Time(Me.txtArrivalTime)
End Sub

Private Sub Image3_click()
TimePicker.Show
Call TimePicker.Selected_Time(Me.txtDepartedTime)
End Sub


Private Sub Image4_Click()
TimePicker.Show
Call TimePicker.Selected_Time(Me.txtFirstTicket)
End Sub

Private Sub Image5_click()
TimePicker.Show
Call TimePicker.Selected_Time(Me.txtLastTicket)
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

'Private Sub TextBox1_Enter()
'GetDate
'End Sub

'Private Sub TextBox2_Enter()
'PX_hide
'End Sub

'Private Sub TextBox3_Enter()
'GetDate
'End Sub

Private Sub TextBox4_Enter()
'    PX_hide
'    TimePicker.Show
'    Call TimePicker.Selected_Time(Me.TextBox4)
End Sub

Private Sub txtArrivalTime_Enter()
    PX_hide
    Me.txtArrivalTime.BackColor = vbWhite
End Sub


Private Sub txtDate_Enter()
Me.txtDate.BackColor = vbWhite
End Sub

Private Sub txtDepartedTime_Enter()
PX_hide
Me.txtDepartedTime.BackColor = vbWhite
End Sub

Private Sub txtFirstTicket_Enter()
Me.txtFirstTicket.BackColor = vbWhite
End Sub

Private Sub txtLastTicket_Enter()
Me.txtLastTicket.BackColor = vbWhite
End Sub

Private Sub txtPassengers_Enter()
Me.txtPassengers.BackColor = vbWhite
End Sub

Private Sub UserForm_Activate()
    '=============
    'Load Origin List
    '==============
    Dim x1 As Range
    For Each x1 In Sheet3.Range("Origin_List")
        Me.cmbOrigin.AddItem x1
    Next x1
    
    '=============
    'Load Arrivedfrom List
    '==============
    Dim x6 As Range
    For Each x6 In Sheet3.Range("Origin_List")
        Me.cmbArrivedFrom.AddItem x6
    Next x6
    

    '=============
    'Load Route List
    '==============
    Dim x3 As Range
    For Each x3 In Sheet3.Range("Route_List")
        Me.cmbRoute.AddItem x3
    Next x3
    
    
    '=============
    'Load BusCode List
    '==============
    Dim x4 As Range
    For Each x4 In Sheet3.Range("BusCode_List")
        Me.cmbBusCode.AddItem x4
    Next x4
    

    '=============
    'Load Captain List
    '==============
    Dim x5 As Range
    For Each x5 In Sheet3.Range("Captains_List")
        Me.cmbCaptainsName.AddItem x5
    Next x5
    
    '=============
    'Load Shift List
    '==============
    
    Me.cmbShift.AddItem "AM"
    Me.cmbShift.AddItem "PM"
    
    Me.ComboBoxTripType.AddItem "Dead Trip"
    Me.ComboBoxTripType.AddItem "Main Trip"
    Me.ComboBoxTripType.AddItem "Sub Trip"
    
    Me.lblUsername.Caption = Sheet2.Range("B3").Value

    Sheet2.Range("A30").Value = 0
    
    
    Dim Ret As Long, styl As Long
    Ret = FindWindow("ThunderDFrame", Me.Caption)

    styl = GetWindowLong(Ret, GWL_STYLE)
    styl = styl Or WS_SYSMENU
    styl = styl Or WS_MINIMIZEBOX
    styl = styl Or WS_MAXIMIZEBOX
    SetWindowLong Ret, GWL_STYLE, (styl)

    DrawMenuBar Ret
End Sub

Private Sub UserForm_Initialize()
DatePickerX_Ini
Sheet3.Range("S2").Value = "Select Destination"
frmTripDetails.Label27.Caption = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
      If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


