VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatePickerX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public WithEvents aMenu As MSForms.Label
Attribute aMenu.VB_VarHelpID = -1

Private Sub aMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
If aMenu.Name = ActiveCtr1 Or aMenu.Name = ActiveCtr2 Then Exit Sub
eCalCtrlRESET

Dim tempUFXi As Object
Set tempUFXi = tempUF

If InStr(1, aMenu.Name, "day") > 0 Then
      tempUFXi.Controls(Replace(aMenu.Name, "day", "s")).BackColor = Color_HoverColor_Back
      ActiveCtr1 = aMenu.Name
      ActiveCtr2 = Replace(aMenu.Name, "day", "s")
Else
      aMenu.BackColor = Color_HoverColor_Back
      ActiveCtr2 = aMenu.Name
      ActiveCtr1 = Replace(aMenu.Name, "s", "day")
End If
Set tempUFXi = Nothing
End Sub
Private Sub aMenu_Click()
   sDate = aMenu.ControlTipText: SetDate
End Sub
