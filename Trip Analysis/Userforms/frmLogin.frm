VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "UserForm1"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'--------------------------------------------------------------
Private Declare PtrSafe Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
               (ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'--------------------------------------------------------------

Dim mdOriginX As Double
Dim mdOriginY As Double

Dim hWndForm As Long

Private Sub Label8_Click()
    Unload Me
    ThisWorkbook.Save
End Sub

Private Sub Label3_Click()

Unload Me

End Sub

Private Sub Image3_click()
  If Me.txtUsername.Value = "" Or Me.txtPassword.Value = "" Then
    Me.Label9.Visible = True
  Else
    Call mMain.Login
    frmLogin.Hide
  End If

End Sub

Private Sub Label10_Click()

'Me.BackColor = x1 '; vbYellow
frmLogin.Hide

End Sub

Private Sub Label11_Click()

Me.BackColor = x2 'vbGreen

End Sub

Private Sub Label12_Click()

Me.BackColor = x3 'vbBlue

End Sub

Private Sub Label13_Click()

Me.BackColor = x4 'vbGreen

End Sub

Private Sub Label6_Click()

Me.Label6.Visible = False
Me.txtUsername.SetFocus

End Sub

Private Sub Label7_Click()
Me.Label7.Visible = False
Me.txtPassword.SetFocus
End Sub

Private Sub txtUsername_Enter()

Me.Label6.Visible = False

End Sub

Private Sub txtUsername_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(txtUsername.Value & vbNullString) = 0 Then
  Me.Label6.Visible = True
End If
End Sub

Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Me.Label6.Visible = False
End Sub


Private Sub txtPassword_Enter()
Me.Label7.Visible = False
End Sub

Private Sub txtPassword_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Len(txtPassword.Value & vbNullString) = 0 Then
  Me.Label7.Visible = True
End If

End Sub

Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Me.Label7.Visible = False
End Sub

Private Sub UserForm_Initialize()
Dim fsdf As Variant, kstr As Variant


x1 = RGB(107, 79, 107)
x2 = RGB(217, 110, 227)
x3 = RGB(217, 90, 227)
x4 = RGB(217, 50, 227)


    Dim lngWinState As XlWindowState
    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, Me.Caption) ' The UserForm must have a caption
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)

    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
    
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    
    With Application
        .screenUpdating = False
        lngWinState = .WindowState
        .WindowState = xlMaximized
            Me.Move 0, 0, .Width, .Height
        .WindowState = lngWinState
        .screenUpdating = True
    End With
    
With Me
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With


Me.BackColor = fsdf

Me.txtPassword.Value = ""
Me.TextBox3.SetFocus
Label9.Visible = False

End Sub
