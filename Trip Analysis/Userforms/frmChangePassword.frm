VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangePassword 
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "frmChangePassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgSave_click()
    Call mMain.Change_Password_DB
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
      If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub
