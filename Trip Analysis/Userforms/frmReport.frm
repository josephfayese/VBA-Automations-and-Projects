VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Generate Report"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonCancel_Click()
    Unload frmReport
End Sub

Private Sub CommandButtonGenerateReport_Click()
    
    If Me.ComboBoxMonth.Value = "" Or Me.ComboBoxYear.Value = "" Then
        MsgBox "Either the Month or Year is empty"
        Exit Sub
    End If
    
    If frmReport.OptionButtonBusPerformance.Value = True Then
        Call mMain.Generate_Bus_Performance
    ElseIf frmReport.OptionButtonRouteSummary.Value = True Then
        Call mMain.Route_Summary
    ElseIf frmReport.OptionButtonStation.Value = True Then
        Call mMain.Loading_Station_Summary
    Else
        MsgBox "Kindly Select the report you wish to generate"
    End If
End Sub

Private Sub UserForm_Activate()
    Dim x As Long
    For x = 2010 To 2030
        Me.ComboBoxYear.AddItem x
    Next x
    
    Me.ComboBoxMonth.AddItem "Jan"
    Me.ComboBoxMonth.AddItem "Feb"
    Me.ComboBoxMonth.AddItem "Mar"
    Me.ComboBoxMonth.AddItem "Apr"
    Me.ComboBoxMonth.AddItem "May"
    Me.ComboBoxMonth.AddItem "Jun"
    Me.ComboBoxMonth.AddItem "Jul"
    Me.ComboBoxMonth.AddItem "Aug"
    Me.ComboBoxMonth.AddItem "Sep"
    Me.ComboBoxMonth.AddItem "Oct"
    Me.ComboBoxMonth.AddItem "Nov"
    Me.ComboBoxMonth.AddItem "Dec"
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
      If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

