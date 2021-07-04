VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Search_List 
   Caption         =   "Serchable Dropdown list"
   ClientHeight    =   2685
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3225
   OleObjectBlob   =   "Frm_Search_List.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Search_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  
Private Sub CommandButton1_Click()

Me.CommandButton2.Visible = True
Me.CommandButton1.Visible = False
Me.Height = 214


End Sub

Private Sub CommandButton2_Click()
Me.CommandButton2.Visible = False
Me.CommandButton1.Visible = True
Me.Height = 158
End Sub
 
Private Sub Label1_Click()
'ActiveWorkbook.FollowHyperlink Address:="https://PK-AnExcelExpert.Com"
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
       Update_Value
       Unload Me
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 13 Then Update_Value
 
End Sub

Private Sub TextBox1_Change()
    Refresh_List
    
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = 13 Then Update_Value
End Sub

Private Sub UserForm_Activate()
 
 Call Refresh_List
  
End Sub

Sub Refresh_List()

Dim arr() As String
Dim rng As Range
Dim cel As Range

Dim i As Integer
Me.ListBox1.Clear

If Is_validation(ActiveCell) Then

    If Validate_Range(ActiveCell.Validation.Formula1) Then
       Set rng = Range(ActiveCell.Validation.Formula1)
       
       For Each cel In rng
            If Me.TextBox1.Value = "" Then
                Me.ListBox1.AddItem cel.Value
            Else
                If VBA.InStr(UCase(cel.Value), UCase(Me.TextBox1.Value)) > 0 Then
                    Me.ListBox1.AddItem cel.Value
                End If
            End If
       Next
    Else
    
    arr() = VBA.Split(ActiveCell.Validation.Formula1, ",")
        For i = LBound(arr) To UBound(arr)
            If Me.TextBox1.Value = "" Then
                Me.ListBox1.AddItem arr(i)
            Else
                If VBA.InStr(UCase(arr(i)), UCase(Me.TextBox1.Value)) > 0 Then
                    Me.ListBox1.AddItem arr(i)
                End If
            End If
        Next i
    End If

End If


On Error Resume Next
Me.ListBox1.ListIndex = 0
 


End Sub

Function Validate_Range(rng_str As String) As Boolean

rng_str = Replace(rng_str, "=", "")

Dim rng As Range
On Error Resume Next
    Set rng = Range(rng_str)
On Error GoTo 0
 
If rng Is Nothing Then
    Validate_Range = False
Else
    Validate_Range = True
End If

End Function

Function Is_validation(rng As Range) As Boolean
    Dim dvtype As Integer
    
    On Error Resume Next
        dvtype = rng.Validation.Type
    On Error Resume Next
    
    If dvtype = 3 Then
        Is_validation = True
    Else
        Is_validation = False
    End If

End Function

Sub Update_Value()

    ActiveCell.Value = Me.ListBox1.Value
    
    If Me.OptionButton1.Value Then
        ActiveCell.Offset(0, 1).Select
        
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            Call Refresh_List
        End If
        
        If Me.ListBox1.ListCount = 0 Then Me.Hide
        
    ElseIf Me.OptionButton2.Value Then
        ActiveCell.Offset(0, 1).Select
        
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            Call Refresh_List
        End If
    
        If Me.ListBox1.ListCount = 0 Then Me.Hide
    
    ElseIf Me.OptionButton3.Value Then
        Me.TextBox1.Value = ""
    Else
        Me.TextBox1.Value = ""
        Me.Hide
    End If
        
End Sub

