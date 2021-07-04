VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmRollBack 
   Caption         =   "Roll Back Data"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "FrmRollBack.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmRollBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonCancel_Click()
    Unload FrmRollBack
End Sub

Private Sub CommandButtonOk_Click()
    Call Rollback_Data_From_Staging
    Unload Me
End Sub

Private Sub ListBoxUsers_Click()

End Sub


Private Sub UserForm_Initialize()
    Call mMain.Load_Users_listbox
End Sub
