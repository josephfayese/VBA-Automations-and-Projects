Attribute VB_Name = "mMain"
Option Explicit
Public x1, x2, x3, x4
Const Connection_MYSQL = "DRIVER={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Port=8080;Database=tripanalysis;User=root;Password=Joe@adekemi12;"
Const Connection_MSSQL = "Provider=MSOLEDBSQL;Server=GOP1096502LT\DATAMINE;Database=Trip_Analytics_DB;Trusted_Connection=yes;"
Const Connection_MSSQL_Production = "Provider=MSOLEDBSQL;Server=192.168.100.6\TRIP_ANALYTICS,1434;Database=Trip_Analytics_DB;UID=Admin_02;PWD=Admin@lbsl02;"


Function Generate_Password(num As Integer)
'PURPOSE: Create a Randomized Password
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'MODIFIED: Joseph Fayese

Dim CharacterBank As Variant
Dim x As Long
Dim str As String

'Test Length Input
  If num < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
  "#", "$", "%", "^", "&", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")
  
'Randomly Select Characters One-by-One
  For x = 1 To num
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
  Next x

'Output Randomly Generated String
  Generate_Password = str
End Function

Sub Connect_to_Database()
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim connection_string As String
Dim sql_query As String
Dim cell As Range
Set conn = New ADODB.Connection
Set rst = New ADODB.Recordset

On Error GoTo Err1
'server string
  connection_string = Connection_MSSQL_Production

conn.Open connection_string
'create table sample
'sql_query = "Create Table sample_table(ID integer not null, Customer_ID bigint, Status text, Sales float);"

For Each cell In Sheet3.Range("I2:N30").Rows
    sql_query = "Insert into services (Origin, Destination, Estimated_Distance, Route_Num, Service_Num, Fare) values('" & cell.Cells(1).Value & "', '" & cell.Cells(2).Value & "','" & cell.Cells(3).Value & "','" & cell.Cells(4).Value & "','" & cell.Cells(5).Value & "','" & cell.Cells(6).Value & "')"
    conn.Execute sql_query
Next cell

MsgBox "Done"
'Drop Table
'sql_query = "Drop Table sample_table"

conn.Close
Exit Sub
Err1:
MsgBox Err.Description
End Sub

'============================================================
'************* Load Services ********************************
'============================================================
Sub Pull_Services()
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim connection_string As String
Dim sql_query As String
Dim cell As Range
Set conn = New ADODB.Connection
Set rst = New ADODB.Recordset

Dim settings As New clsExcelSettings

settings.TurnOff

Sheet2.Range("S1").ClearContents

On Error GoTo Err1
'server string
  connection_string = Connection_MSSQL_Production

conn.Open connection_string

sql_query = "Select Fare from ServiceTbl Where Origin = '" & frmFareRate.ComboBoxOrigin.Value & "' and Destination = '" & frmFareRate.ComboBoxDestination.Value & "'"

With rst
    .ActiveConnection = conn
    .Source = sql_query
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

Sheet2.Range("S1").CopyFromRecordset rst

rst.Close
conn.Close

settings.TurnOn

Exit Sub

Err1:
MsgBox Err.Description
settings.Restore
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Load Services ********************************
'============================================================
Sub Update_to_Services()
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim connection_string As String
Dim sql_query As String
Dim cell As Range
Set conn = New ADODB.Connection
Set rst = New ADODB.Recordset

On Error GoTo Err1
  connection_string = Connection_MSSQL_Production

conn.Open connection_string

    sql_query = "Update ServiceTbl SET Fare = " & frmFareRate.txtNewFare.Value & " Where Origin = '" & frmFareRate.ComboBoxOrigin.Value & "' and Destination = '" & frmFareRate.ComboBoxDestination.Value & "'"
    conn.Execute sql_query


MsgBox "Fare Updated Successfully!!!"

conn.Close
Exit Sub
Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Get email Function ***********************
'============================================================
Function Get_Email()
Dim outlookapp As Outlook.Application
Dim oMail As Outlook.MailItem
Dim outSession As Outlook.Account
Dim currentUserEmailAddress As String

Set outlookapp = New Outlook.Application
Set oMail = outlookapp.CreateItem(olMailItem)

'Get sender's email Address
 currentUserEmailAddress = outlookapp.Session.Accounts(1)
 Get_Email = currentUserEmailAddress

End Function
'==================================================
' **************** End of Code *******************
'==================================================


'============================================================
'************* Send email Subroutine ************************
'============================================================
Sub Send_Email()
Dim outlookapp As Outlook.Application
Dim oMail As Outlook.MailItem

Set outlookapp = New Outlook.Application
Set oMail = outlookapp.CreateItem(olMailItem)

With oMail
    .Subject = "Login Credentials for Trip Analysis App"
    .To = frmUsers_Assign.txtEmail
'    .SenderEmailAddress = Get_Email
    .BodyFormat = olFormatHTML
    .HTMLBody = "Dear " & frmUsers_Assign.txtUserID & ", <br><br>Please see below your login credentials for the Trip Analysis Excel Application<br><br>" & _
            "<strong>Username:</strong> " & frmUsers_Assign.txtUsername & "<br><br>" & _
            "<strong>Password:</strong> " & frmUsers_Assign.txtPWD & "<br><br><br>" & _
            "<strong>NB:</strong> You will be required to change your password upon first login <br><br>Kind Regards,<br>Admin"
    .Send
End With


End Sub
'==================================================
' **************** End of Code *******************
'==================================================


'============================================================
'************* Trip Details Subroutine **********************
'============================================================
Sub Data_Entry()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim i, j As Long
    Dim ID As Long
    Dim cell As Range
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo Err1
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    i = Sheet2.Range("A30").Value
    j = Sheet2.Range("B30").Value
    
    sql_query = "SELECT Trip_Number FROM TripSummaryTbl WHERE Trip_Number = '" & frmTripDetails.ComboBoxTripNumber.Value & "'"
    
    rst.Open sql_query, conn, adOpenKeyset, adLockOptimistic
    
    Sheet2.Range("C30").CopyFromRecordset rst
    
    rst.Close
    
    i = i + 1
        If j > 20 Then
            MsgBox "Please change TripID before to enter new data"
            j = 0
            Sheet2.Range("B30").Value = j
            Exit Sub
        'ElseIf j < 20 Then
         '   j = 0
        End If
    
    'check if there are missing fields
    Dim ctrl1 As Control
    For Each ctrl1 In frmTripDetails.Frame1.Controls
        If TypeName(ctrl1) = "TextBox" Or TypeName(ctrl1) = "ComboBox" Then
            If ctrl1.Value = "" Then
                ctrl1.BackColor = RGB(255, 217, 217)
            End If
        End If
    Next ctrl1
    
    If frmTripDetails.txtDate.Value = "" Then
        frmTripDetails.txtDate.BackColor = RGB(255, 217, 217)
    End If
    
    'input into the users table
    sql_query = "SELECT * FROM TripSummaryTbl"
    
        With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        .AddNew
        .Fields(1).Value = mTest.CreateGUID() & "|" & frmTripDetails.ComboBoxTripNumber.Value
        .Fields(2).Value = frmTripDetails.lblUsername.Caption
        .Fields(3).Value = frmTripDetails.cmbBusCode.Value
        .Fields(4).Value = FormatDate(frmTripDetails.txtDate.Value)
        .Fields(5).Value = frmTripDetails.cmbRoute.Value
        .Fields(6).Value = frmTripDetails.cmbArrivedFrom.Value
        .Fields(7).Value = frmTripDetails.txtArrivalTime.Value
        .Fields(8).Value = frmTripDetails.txtDepartedTime.Value
        .Fields(9).Value = frmTripDetails.cmbOrigin.Value
        .Fields(10).Value = frmTripDetails.cmbDestination.Value
        .Fields(11).Value = frmTripDetails.cmbShift.Value
        .Fields(12).Value = frmTripDetails.cmbCaptainsName.Value
        .Fields(13).Value = frmTripDetails.txtPassengers.Value
        If frmTripDetails.ComboBoxTripType.Value = "Main Trip" Then
            .Fields(14).Value = 1
        ElseIf frmTripDetails.ComboBoxTripType.Value = "Sub Trip" Then
            .Fields(14).Value = 2
        Else
            .Fields(14).Value = 3
        End If
        .Fields(15).Value = Application.WorksheetFunction.Index(Sheet3.Range("Q2:Q25"), Application.WorksheetFunction.Match(frmTripDetails.cmbOrigin.Value, Sheet3.Range("R2:R25"), 0))
        .Fields(16).Value = frmTripDetails.txtFirstTicket.Value
        .Fields(17).Value = frmTripDetails.txtLastTicket.Value
        .Fields(18).Value = Application.WorksheetFunction.VLookup(frmTripDetails.cmbOrigin.Value & "|" & frmTripDetails.cmbDestination.Value, Sheet3.Range("H2:N53"), 7, 0)
        .Fields(19).Value = FormatDate(VBA.Date())
        
        Sheet2.Range("A24").Value = frmTripDetails.cmbBusCode.Value
        Sheet2.Range("A25").Value = frmTripDetails.cmbRoute.Value
        
        If TimeValue(frmTripDetails.txtArrivalTime.Value) > TimeValue(frmTripDetails.txtDepartedTime.Value) Then
            MsgBox "Arrival time can't be more than departed time"
            Exit Sub
        End If
        
        If TimeValue(frmTripDetails.txtLastTicket.Value) < TimeValue(frmTripDetails.txtFirstTicket.Value) Then
            MsgBox "Last ticket time can't be less than first ticket time"
            Exit Sub
        End If
        
        
        Dim message As String
        message = MsgBox("Kindly check if all fields are imputed correctly" & vbCrLf & vbCrLf & "Do you wish to submit", vbYesNo + vbInformation, "Status")
        If message = vbYes Then
            .Update
            MsgBox "Data Sucessfully imputed!!!"
            'clear all inputs in form
            Dim ctrl As Control
            For Each ctrl In mApp.Controls
            Debug.Print TypeName(ctrl)
            If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
                ctrl.Value = ""
            End If
            Next ctrl
        ElseIf message = vbNo Then
            .Close
            Exit Sub
        End If
        .Close
    End With
    
    conn.Close
    
    'clear all info
        For Each ctrl1 In frmTripDetails.Frame1.Controls
            If TypeName(ctrl1) = "TextBox" Or TypeName(ctrl1) = "ComboBox" Then
                ctrl1.Value = ""
            End If
        Next ctrl1
    
Call mMain.Audit_Trail
j = j + 1
frmTripDetails.Label27.Caption = i
Sheet2.Range("A30").Value = i
Sheet2.Range("B30").Value = j
Exit Sub

Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Users Subroutine ********************************
'============================================================
Sub Create_Users()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim cell As Range
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo Err1
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    'input into the users table
    sql_query = "SELECT * FROM UsersTbl"
    
    'check if there are missing fields
    Dim ctrl1 As MSForms.Control
    For Each ctrl1 In frmUsers_Assign.Controls
        If TypeName(ctrl1) = "TextBox" Or TypeName(ctrl1) = "ComboBox" Then
            If ctrl1.Value = "" Then
                ctrl1.BackColor = RGB(255, 217, 217)
            End If
        End If
    Next ctrl1
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        Do Until .EOF
            If .Fields(1).Value = frmUsers_Assign.txtUserID.Value Then
                MsgBox "User Already Created!!!"
                Dim ctrl2 As MSForms.Control
                For Each ctrl2 In frmUsers_Assign.Controls
                    If TypeName(ctrl2) = "TextBox" Or TypeName(ctrl2) = "ComboBox" Then
                        ctrl2.Value = ""
                    End If
                Next ctrl2
                rst.Close
                conn.Close
                Exit Sub
            End If
            .MoveNext
        Loop
        .AddNew
        .Fields(1).Value = frmUsers_Assign.txtUserID.Value
        .Fields(2).Value = FormatDate(frmUsers_Assign.txtDate_Assigned.Value)
        .Fields(3).Value = frmUsers_Assign.txtUsername.Value
        .Fields(4).Value = frmUsers_Assign.txtPWD.Value
        .Fields(5).Value = frmUsers_Assign.txtEmail.Value
        .Fields(6).Value = frmUsers_Assign.cmbRole.Value
        .Update
        .Close
    End With
    
    'input into the login table
    new_query = "INSERT INTO LoginTbl (Username, Password) VALUES ('" & frmUsers_Assign.txtUsername.Value & "', '" & frmUsers_Assign.txtPWD.Value & "')"
    
    conn.Execute new_query
    
    conn.Close
    Call mMain.Send_Email
    Call mMain.Audit_Trail
    MsgBox "User Assigned Successfully"
    Dim ctrl3 As MSForms.Control
    For Each ctrl3 In frmUsers_Assign.Controls
        If TypeName(ctrl3) = "TextBox" Or TypeName(ctrl3) = "ComboBox" Then
            ctrl3.Value = ""
        End If
    Next ctrl3
Exit Sub

Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Login Subroutine *****************************
'============================================================
Sub Login()

    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim cell As Range
    Dim i As Integer
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    Application.screenUpdating = False
    
    On Error GoTo ErrorHandler_1
    i = 0
    
    'server string
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT * FROM LoginTbl"
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    If frmLogin.txtUsername.Value = "Admin01" Then
        If frmLogin.txtUsername.Value <> "Admin01" Or frmLogin.txtPassword.Value <> "Adminlbsl@01" Then
            MsgBox "Either Username or Password is incorrect"
            frmLogin.txtUsername.Value = ""
            frmLogin.txtPassword.Value = ""
            Exit Sub
        End If
        If rst.Fields(1).Value = "Admin01" And rst.Fields(2).Value = "Adminlbsl@01" Then
            MsgBox "Logged in Successfully!!!"
            Sheet2.Range("A20").Value = frmLogin.txtUsername.Value
            Sheet2.Range("A22").Value = frmLogin.txtPassword.Value
            frmLogin.Hide
            frmUsers_Assign.Show
            rst.Close
            conn.Close
            Exit Sub
        End If
    Else
        Do Until rst.EOF
            If rst.Fields(1).Value = frmLogin.txtUsername.Value And rst.Fields(2).Value = frmLogin.txtPassword.Value Then
                Sheet2.Range("B3").Value = frmLogin.txtUsername.Value & "_01"
                Sheet2.Range("A20").Value = frmLogin.txtUsername.Value
                Sheet2.Range("A22").Value = frmLogin.txtPassword.Value
                MsgBox "Logged in Successfully!!!"
              Exit Do
            End If
            rst.MoveNext
            If rst.EOF Then
                MsgBox "User hasn't been created yet. Kindly contact admin"
                rst.Close
                conn.Close
                Application.screenUpdating = True
                Exit Sub
            End If
        Loop
    End If
    Call mMain.Audit_Trail
    If Sheet2.Range("P1").Value <= 1 Then
        MsgBox "Please Kindly change your password"
        frmChangePassword.Show
        MsgBox "Password changed successfully!!!"
        Unload frmLogin
    End If
    If Sheet2.Range("A20").Value = "taiwo.olusola" Or Sheet2.Range("A20").Value = "babatunde.obadina" Then
        Unload frmLogin
        Sheet5.Visible = xlSheetVisible
        Sheet4.Visible = xlSheetVeryHidden
        Sheet6.Visible = xlSheetVeryHidden
        Sheet7.Visible = xlSheetVeryHidden
        Dim lrow As Long
        lrow = Sheet5.Range("A" & Application.Rows.Count).End(xlUp).Row
        'If lrow > 5 Then
            'Sheet5.Range("A6:C" & lrow).ClearContents
            'Sheet5.Range("E6:M" & lrow).ClearContents
            'Sheet5.Range("O6:P" & lrow).ClearContents
        'End If
        Sheet5.Range("A6").Select
    ElseIf Sheet2.Range("A20").Value = "olatunji.olopade" Then
        Unload frmLogin
        Sheet7.Visible = xlSheetVisible
        Sheet4.Visible = xlSheetVeryHidden
        Sheet6.Visible = xlSheetVeryHidden
        Sheet5.Visible = xlSheetVeryHidden
        Sheet7.Range("B3").Select
    Else
        Unload frmLogin
        Sheet6.Visible = xlSheetVisible
        Sheet4.Visible = xlSheetVeryHidden
        Sheet5.Visible = xlSheetVeryHidden
        Sheet7.Visible = xlSheetVeryHidden
        lrow = Sheet6.Range("A" & Application.Rows.Count).End(xlUp).Row
        If lrow > 5 Then
            Sheet6.Range("A6:C" & lrow).ClearContents
            Sheet6.Range("E6:M" & lrow).ClearContents
            Sheet6.Range("O6:P" & lrow).ClearContents
        End If
        Sheet6.Range("A6").Select
    End If
    Application.screenUpdating = True
    Exit Sub
    
ErrorHandler_1:
i = i + 1
If i = 3 Then
    MsgBox "Kindly check your internet settings or login cretentials"
    Exit Sub
    Application.screenUpdating = True
End If
Resume

End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Audit Trail Subroutine ***********************
'============================================================
Sub Audit_Trail()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim cell As Range
    Dim CurrentTime As String
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT * FROM AuditTrailTbl"
    
    CurrentTime = VBA.Time()
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        .AddNew
        .Fields(1).Value = Sheet2.Range("A20").Value & "_01"
        .Fields(2).Value = Sheet2.Range("A20").Value
        .Fields(3).Value = VBA.Date()
        .Fields(4).Value = TimeValue(CurrentTime)
'        .Fields(5).Value = "Logged in Successfully"
        If frmLogin.Visible = True Then
            .Fields(5).Value = "User logged in Successfully"
        ElseIf frmUsers_Assign.Visible = True Then
            .Fields(5).Value = "User Assigned Successfully"
        ElseIf frmTripDetails.Visible = True Then
            .Fields(5).Value = "Created a Trip with BusCode - " & Sheet2.Range("A24").Value & "|Route Number - " & Sheet2.Range("A25").Value & " Successfully"
        Else
            .Fields(5).Value = "User logged in Successfully"
        End If
        .Update
    End With
    
    rst.Close
    
    'loging in for the first time
    new_query = "SELECT COUNT(Username) FROM AuditTrailTbl WHERE UserID = " & "'" & Sheet2.Range("A20").Value & "_01" & "'"
    With rst
        .ActiveConnection = conn
        .Source = new_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    Sheet2.Range("P1").CopyFromRecordset rst
    
    rst.Close
    conn.Close
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* User Activity check Subroutine ***************
'============================================================
Sub User_Activity()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim ext As String
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo Err1
    
    frmTripSummary.Show
    
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT User_No, Bus_Code, Route_No, Time_logged_on as Time_of_Entry, Date_Capture, Action_taken FROM TripSummaryTbl " & _
                "JOIN AuditTrailTbl ON TripSummaryTbl.User_No = AuditTrailTbl.UserID " & _
                "WHERE Date_Capture BETWEEN " & "'" & Format(Sheet1.Range("A2").Value, "YYYY-MM-DD") & "'" & " AND " & "'" & Format(Sheet1.Range("A3").Value, "YYYY-MM-DD") & "'"
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    'open a new workbook
    Dim xpath As String
    Dim strDir As String
    Dim strPath As String
    Application.screenUpdating = False
    Application.DisplayAlerts = False
    strDir = "User_Activity"
    strPath = "C:\Users\" & Environ("username") & "\Documents\"
    
    Workbooks.Add
    
    xpath = MkPath(strDir, strPath)
    
    ActiveWorkbook.SaveAs xpath & "\User_Activity_Report" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
    Dim workbookname As String
    workbookname = "User_Activity_Report" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "UserID"
            .Range("B1").Value = "Bus_Code"
            .Range("C1").Value = "Route Number"
            .Range("D1").Value = "Date_of_Entry"
            .Range("E1").Value = "Action_Taken"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Actvity report created successfully, Please check the path " & xpath
    
        
    rst.Close
    conn.Close
    
Exit Sub

Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================


'============================================================
'************* Trip Summary Report Subroutine ***************
'============================================================
Sub Create_Trip_Report()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim ext As String
    Dim i As Long
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo Err1
    
    frmTripSummary.Show
    
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT Bus_Code, Date_capture, Route_No, Arrived_from, XXX.Origin, XXX.Destination, Arrival_Time, Depart_Time, Shift_time, StationID, Estimated_Distance, Passenger,Captain_Name, Fare_Rate, " & _
                "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END Trip_Type, SUM(Passenger * Fare_Rate) Revenue " & _
                "FROM TripSummaryTbl XXX " & _
                "JOIN ServiceTbl YYY ON XXX.Origin + '-' + XXX.Destination = YYY.Origin + '-' + YYY.Destination " & _
                "WHERE Date_Capture BETWEEN " & "'" & Format(Sheet1.Range("A2").Value, "YYYY-MM-DD") & "'" & " AND " & "'" & Format(Sheet1.Range("A3").Value, "YYYY-MM-DD") & "'" & _
                "GROUP BY Bus_Code, Date_capture, Route_No, Arrived_from, XXX.Origin, XXX.Destination, Arrival_Time, Depart_Time, Shift_time, StationID, Estimated_Distance, Passenger, Captain_Name, Fare_Rate," & _
                "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END"
                
   
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

'open a new workbook

Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
Application.DisplayAlerts = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

    ActiveWorkbook.SaveAs xpath & "\Trip_Summary_Report_" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
    Dim workbookname As String
    workbookname = "Trip_Summary_Report_" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "Bus_Code"
            .Range("B1").Value = "Date_capture"
            .Range("C1").Value = "Route_No"
            .Range("D1").Value = "Arrived_from"
            .Range("E1").Value = "Origin"
            .Range("F1").Value = "Destination"
            .Range("G1").Value = "Arrival_Time"
            .Range("H1").Value = "Depart_Time"
            .Range("I1").Value = "Shift_time"
            .Range("J1").Value = "StationID"
            .Range("K1").Value = "Estimated Distance"
            .Range("L1").Value = "No_of_Passenger"
            .Range("M1").Value = "Captain_Name"
            .Range("N1").Value = "Fare"
            .Range("O1").Value = "Trip_Type"
            .Range("P1").Value = "Revenue Per Trip"
            .Range("P1").Value = "Estimated Distance"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Trip Summary report created successfully, Please check the path " & xpath
    rst.Close
    conn.Close
Exit Sub

Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

Function MkPath(strDir As String, strPath As String)

Dim fso As Object
Dim path As String

Set fso = CreateObject("Scripting.FileSystemObject")
'strDir = "User_Activity"
'strPath = "C:\Users\" & Environ("username") & "\Documents"

path = strPath & strDir

If Not fso.FolderExists(path) Then

' doesn't exist, so create the folder
        fso.CreateFolder path
End If

MkPath = path

End Function


Sub check()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim i As Integer
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    'server string
    On Error GoTo ErrorHandler_1
    i = 0
    conn.ConnectionString = MyConnectionString
                       
    sql_query = "CREATE TABLE TestTbl_2 ( User_ID varchar(20) not null, First_Name varchar(50), Last_Name varchar(50), Email_Address varchar(100))"
                        
    conn.Open connection_string
    
    conn.Execute sql_query
    
    conn.Close
    Exit Sub
ErrorHandler_1:
i = i + 1
If i = 3 Then
    MsgBox "Kindly check your internet settings or login cretentials"
    Exit Sub
End If
Resume

End Sub

Sub launch_app()
    frmLogin.Show
End Sub

Sub Change_Password_DB()
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim i As Integer
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    On Error GoTo Err1
    
    sql_query = "SELECT Password FROM LoginTbl WHERE Username = " & "'" & Sheet2.Range("A20").Value & "'"
    
    con.ConnectionString = Connection_MSSQL_Production
    
    If frmChangePassword.txtOldPassword.Value = "" Or frmChangePassword.txtNewPassword.Value = "" Then
        MsgBox "Passwords can't be empty"
    End If
    
'TryAgain:
'        If Sheet2.Range("A21").Value = frmChangePassword.txtOldPassword.Value Then
'            .Fields("Password").Value = frmChangePassword.txtNewPassword.Value
'        Else
'            MsgBox "Your Old Password didn't match"
'            GoTo TryAgain
'        End If
    
    con.Open
    
'    con.Execute sql_query
    
    
    With rs
        .ActiveConnection = con
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        .Fields("Password").Value = frmChangePassword.txtNewPassword.Value
        .Update
    End With

    Sheet2.Range("A22").Value = frmChangePassword.txtNewPassword.Value
    rs.Close
    con.Close
    
Exit Sub

Err1:
MsgBox Err.Description
End Sub

Public Function IsLoaded(formName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
End Function

'============================================================
'************* User Data Summary Report Subroutine **********
'============================================================
Sub Create_User_Report()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim ext As String
    Dim i As Long
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo Err1
    
    frmTripSummary.Show
    
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT Bus_Code, Date_capture, Route_No, Arrived_from, Origin, Destination, Arrival_Time, Depart_Time, Shift_time, StationID, Passenger, Captain_Name, Fare_Rate, " & _
                "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END Trip_Type, SUM(Passenger * Fare_Rate) Revenue " & _
                "FROM TripSummaryTbl " & _
                "WHERE Date_Capture BETWEEN " & "'" & Format(Sheet1.Range("A2").Value, "YYYY-MM-DD") & "'" & " AND " & "'" & Format(Sheet1.Range("A3").Value, "YYYY-MM-DD") & "'" & _
                "AND User_No = '" & Sheet2.Range("B3").Value & "'" & _
                "GROUP BY Bus_Code, Date_capture, Route_No, Arrived_from, Origin, Destination, Arrival_Time, Depart_Time, Shift_time, StationID,Passenger, Captain_Name, Fare_Rate," & _
                "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END"
   
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

'open a new workbook

Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
Application.DisplayAlerts = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

    ActiveWorkbook.SaveAs xpath & "\User_Summary_Report_" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
    Dim workbookname As String
    workbookname = "User_Summary_Report_" & Format(VBA.Date(), "YYYYMMDD") & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "Bus_Code"
            .Range("B1").Value = "Date_capture"
            .Range("C1").Value = "Route_No"
            .Range("D1").Value = "Arrived_from"
            .Range("E1").Value = "Origin"
            .Range("F1").Value = "Destination"
            .Range("G1").Value = "Arrival_Time"
            .Range("H1").Value = "Depart_Time"
            .Range("I1").Value = "Shift_time"
            .Range("J1").Value = "StationID"
            .Range("K1").Value = "No_of_Passenger"
            .Range("L1").Value = "Captain_Name"
            .Range("M1").Value = "Fare"
            .Range("N1").Value = "Trip_Type"
            .Range("O1").Value = "Revenue Per Trip"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Trip Summary report created successfully, Please check the path " & xpath
    rst.Close
    conn.Close
Exit Sub

Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* User Data Summary Report Subroutine **********
'============================================================
Sub View_Entries()
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim connection_string As String
Dim sql_query As String
Dim lw_rec As ListItem
Dim i As Long, x As Long
Set conn = New ADODB.Connection
Set rst = New ADODB.Recordset

On Error GoTo Err1
'server string
  connection_string = Connection_MSSQL_Production

conn.Open connection_string


sql_query = "SELECT * FROM TripSummaryTbl WHERE User_No = '" & Sheet2.Range("A20").Value & "_01' AND Date_Entered = '" & VBA.Date() & "'"

With rst
    .ActiveConnection = conn
    .Source = sql_query
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

With frmUserSummary.ListViewSummary
    For i = 0 To rst.Fields.Count - 1
        .ColumnHeaders.Add , , rst.Fields(i).Name
    Next i
End With

With frmUserSummary.ListViewSummary
    .ListItems.Clear
    
    Do While Not rst.EOF
    
        Set lw_rec = .ListItems.Add(, , rst.Fields(0).Value)
        
        For x = 1 To rst.Fields.Count - 1
            lw_rec.SubItems(x) = IIf(IsNull(rst.Fields(x).Value), "-", rst.Fields(x).Value)
        Next x
        
        rst.MoveNext
    
    Loop
    .FullRowSelect = True
    .Gridlines = True
    .View = lvwReport
End With

rst.Close
conn.Close

Exit Sub
Err1:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

Sub LogOut()
If Sheet2.Range("A20").Value = "taiwo.olusola" Or Sheet2.Range("A20").Value = "babatunde.obadina" Then
    Sheet4.Visible = xlSheetVisible
    Sheet5.Visible = xlSheetVeryHidden
    Sheet4.Range("A1").Select
Else
    Sheet4.Visible = xlSheetVisible
    Sheet6.Visible = xlSheetVeryHidden
    Sheet4.Range("A1").Select
End If
End Sub


Sub Data_Entry_2()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim lrow, i, j As Long
    Dim message As Variant
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim TripID As String
    
    Call Data_Entry_4
    
    Sheet6.Range("A5").Select
    
    j = 0
    
    lrow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    
    'checks
    message = MsgBox("Do you wish to load all your entries now?", vbYesNo)
    If message = vbYes Then
        'Do Nothing
    Else
        Exit Sub
    End If
    For i = 6 To lrow
        If Sheet6.Range("A" & i).Offset(0, 5).Value > Sheet6.Range("A" & i).Offset(0, 6).Value Then
            MsgBox "Arrival time can't be more than departed time" & vbCrLf & "Check Row A" & i
            Exit Sub
        End If
        
        If Sheet6.Range("A" & i).Offset(0, 14).Value > Sheet6.Range("A" & i).Offset(0, 15).Value Then
            MsgBox "Last ticket time can't be less than first ticket time" & vbCrLf & "Check Row A" & i
            Exit Sub
    End If
    
    Next i
    
    On Error GoTo Err1
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = "SELECT * FROM Staging_table"
    
        With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
        For i = 6 To lrow
            TripID = mTest.CreateGUID()
            .AddNew
            .Fields(0).Value = TripID & "|" & Sheet6.Range("A" & i).Value
            .Fields(1).Value = Sheet2.Range("B3").Value
            .Fields(2).Value = Sheet6.Range("A" & i).Offset(0, 1).Value
            .Fields(3).Value = FormatDate(Sheet6.Range("A" & i).Offset(0, 2).Value)
            .Fields(4).Value = Sheet6.Range("A" & i).Offset(0, 3).Value
            .Fields(5).Value = Sheet6.Range("A" & i).Offset(0, 4).Value
            .Fields(6).Value = Format(Sheet6.Range("A" & i).Offset(0, 5).Value, "hh:mm:ss AM/PM")
            .Fields(7).Value = Format(Sheet6.Range("A" & i).Offset(0, 6).Value, "hh:mm:ss AM/PM")
            .Fields(8).Value = Sheet6.Range("A" & i).Offset(0, 7).Value
            .Fields(9).Value = Sheet6.Range("A" & i).Offset(0, 8).Value
            .Fields(10).Value = Sheet6.Range("A" & i).Offset(0, 9).Value
            .Fields(11).Value = Sheet6.Range("A" & i).Offset(0, 10).Value
            .Fields(12).Value = Sheet6.Range("A" & i).Offset(0, 13).Value
            If Sheet6.Range("A" & i).Offset(0, 11).Value = "Main Trip" Then
                .Fields(13).Value = 1
            ElseIf Sheet6.Range("A" & i).Offset(0, 11).Value = "Sub Trip" Then
                .Fields(13).Value = 2
            Else
                .Fields(13).Value = 3
            End If
            .Fields(14).Value = Sheet6.Range("A" & i).Offset(0, 12).Value
            .Fields(15).Value = Format(Sheet6.Range("A" & i).Offset(0, 14).Value, "hh:mm:ss AM/PM")
            .Fields(16).Value = Format(Sheet6.Range("A" & i).Offset(0, 15).Value, "hh:mm:ss AM/PM")
            .Fields(17).Value = Application.WorksheetFunction.VLookup(Sheet6.Range("A" & i).Offset(0, 7).Value & "|" & Sheet6.Range("A" & i).Offset(0, 8).Value, Sheet3.Range("H2:N150"), 7, 0)
            .Fields(18).Value = FormatDate(VBA.Date())
            .Update
            j = j + 1
            Application.StatusBar = j & " records out of " & lrow - 5 & " loaded successfully"
        Next i
    End With
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing
    
    Call Audit_Trail_2
    If lrow > 5 Then
        Sheet6.Range("A6:C" & lrow).ClearContents
        Sheet6.Range("E6:L" & lrow).ClearContents
        Sheet6.Range("N6:P" & lrow).ClearContents
    End If
    
    MsgBox "Data Successfully loaded"
    Application.StatusBar = ""
    'Sheet5.Range("M6").Formula = "=IFERROR(INDEX(Sheet3!$Q$2:$Q$23,MATCH(H6," & "Origin_List,0)),"""")"
    'Sheet5.Range("M6").AutoFill Sheet5.Range("M6:M10000")
Exit Sub

Err1:
MsgBox Err.Description
rst.Close
conn.Close
Set rst = Nothing
Set conn = Nothing

End Sub
'==================================================
' **************** End of Code *******************
'==================================================

'============================================================
'************* Audit Trail Subroutine Two *******************
'============================================================
Sub Audit_Trail_2()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim lrow, i As Long
    Dim CurrentTime As String
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    lrow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
    On Error GoTo ErrorHandler_2
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "SELECT * FROM AuditTrailTbl"
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        For i = 6 To lrow
            .AddNew
            CurrentTime = VBA.Time()
            .Fields(1).Value = Sheet2.Range("A20").Value & "_01"
            .Fields(2).Value = Sheet2.Range("A20").Value
            .Fields(3).Value = VBA.Date()
            .Fields(4).Value = TimeValue(CurrentTime)
            .Fields(5).Value = "Created a Trip with BusCode - " & Sheet5.Range("B" & i).Value & "|Route Number - " & Sheet5.Range("D" & i).Value & " Successfully"
            .Update
        Next i
    End With
    
    rst.Close
    
    'loging in for the first time
    new_query = "SELECT COUNT(Username) FROM AuditTrailTbl WHERE UserID = " & "'" & Sheet2.Range("A20").Value & "_01" & "'"
    With rst
        .ActiveConnection = conn
        .Source = new_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    Sheet2.Range("P1").CopyFromRecordset rst
    
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

Public Function MyConnectionString() As String
 
    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myIPAddress     As String
    Dim conn_string     As String
    Dim GetMyLocalIP    As String
    
    'Set the computer.
    strComputer = "."
 
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 
    'A select query is used to get a collection of IP addresses from the network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
    'Loop through all the objects of the collection and return the first non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then myIPAddress = Trim(objItem.IPAddress(0))
        Exit For
    Next
 
    'Return the IP string.
    GetMyLocalIP = myIPAddress

    
    If GetMyLocalIP = "192.168.100.154" Then
    
        MyConnectionString = Connection_MSSQL_Production
    Else
        MyConnectionString = Connection_MSSQL
    End If
 
End Function


Sub Generate_Bus_Performance()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim mth As String
    Dim CurrentTime As String
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
    mth = frmReport.ComboBoxMonth & "-" & frmReport.ComboBoxYear
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "select Bus_code, count(UniqueID) as [Total Trips], sum(Passenger) as [Total Passenger], sum(Passenger * Fare_Rate) as [Total Revenue], " & _
                "round(sum(Estimated_Distance),2) as [Total Mileage (KM)], rank() over(order by sum(Passenger * Fare_Rate) desc ) as [Revenue Rank] from " & _
                 "(select Origin + ' | ' + Destination as UniqueID, Date_capture, Bus_code, Passenger, Fare_Rate " & _
                 "from TripSummaryTbl)a " & _
                "Left Join " & _
                "(select Origin + ' | ' + Destination as UniqueID_1, Estimated_Distance from ServiceTbl WHERE Route_Num <> 'XXXXX')b " & _
                "on a.UniqueID = b.UniqueID_1 " & _
                "where CONCAT(FORMAT(Date_capture,'MMM'), '-', datepart(yy,Date_capture))=" & "'" & mth & "'" & _
                "group by Bus_code"
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

'open a new workbook

Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
Application.DisplayAlerts = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

    ActiveWorkbook.SaveAs xpath & "\Bus_Performance_Report_" & mth & ".xlsx"
    Dim workbookname As String
    workbookname = "Bus_Performance_Report_" & mth & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "Bus Code"
            .Range("B1").Value = "Total Trips"
            .Range("C1").Value = "Total Passenger"
            .Range("D1").Value = "Total Revenue"
            .Range("E1").Value = "Total Mileage (KM)"
            .Range("F1").Value = "REVENUE RANK"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Bus Performance report created successfully, Please check the path " & xpath
    rst.Close
    conn.Close
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub


Sub Route_Summary()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim mth As String
    Dim CurrentTime As String
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
    mth = frmReport.ComboBoxMonth & "-" & frmReport.ComboBoxYear
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "select Route_Name +' ' + Route_No as [Route Matrix], count(Trip_Type) as [No of Trips], Trip_Type, sum(Passenger) as [Total Passenger], sum(Passenger * Fare_Rate) as [Total Revenue], " & _
                "round(sum(Estimated_Distance),2) as [Total Mileage (KM)], round(sum(Passenger * Fare_Rate)/sum(Estimated_Distance),2) as [Rev/Mileage] ," & _
                "rank() over(order by sum(Passenger * Fare_Rate) desc ) as [Revenue Rank] from " & _
                 "(select Origin + ' | ' + Destination as UniqueID, Route_No, " & _
                 "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END Trip_Type, " & _
                 "Date_capture , Bus_code, Passenger, Fare_Rate " & _
                 "from TripSummaryTbl)a " & _
                "Left Join " & _
                "(select Origin + ' | ' + Destination as UniqueID_1, Estimated_Distance from ServiceTbl WHERE Route_Num <> 'XXXXX')b " & _
                "on a.UniqueID = b.UniqueID_1 " & _
                "inner join RouteTbl c " & _
                "on a.Route_No = c.Route_ID " & _
                "where CONCAT(FORMAT(Date_capture,'MMM'), '-',datepart(yy,Date_capture) ) =" & "'" & mth & "'" & _
                "group by  Route_Name +' ' + Route_No, Trip_Type " & _
                "order by Trip_Type"
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

'open a new workbook

Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
Application.DisplayAlerts = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

    ActiveWorkbook.SaveAs xpath & "\Route_Summary_Report_" & mth & ".xlsx"
    Dim workbookname As String
    workbookname = "Route_Summary_Report_" & mth & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "ROUTE MATRIX"
            .Range("B1").Value = "NO OF TRIPS"
            .Range("C1").Value = "TRIP TYPE"
            .Range("D1").Value = "PASSENGERS"
            .Range("E1").Value = "REVENUE"
            .Range("F1").Value = "MAIN MILEAGE"
            .Range("G1").Value = "REV/MILEAGE"""
            .Range("H1").Value = "REVENUE RANK"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Route Summary created successfully, Please check the path " & xpath
    rst.Close
    conn.Close
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub


Sub Loading_Station_Summary()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim mth As String
    Dim CurrentTime As String
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
    mth = frmReport.ComboBoxMonth & "-" & frmReport.ComboBoxYear
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string

    sql_query = "select Origin, Destination, count(Trip_Type) as [No Of Trips],Trip_Type, sum(Passenger) as [Total Passenger], sum(Passenger * Fare_Rate) as [Total Revenue], " & _
                "round(sum(Estimated_Distance),2) as [Total Mileage (KM)], rank() over(order by sum(Passenger * Fare_Rate) desc ) as [Revenue Rank] from " & _
                 "(select Origin + ' | ' + Destination as UniqueID, Origin, Destination, " & _
                 "CASE " & _
                    "WHEN Trip_ID = 1 THEN 'Main Trip' " & _
                    "WHEN Trip_ID = 2 THEN 'Sub Trip' " & _
                    "Else 'Dead Trip' " & _
                "END Trip_Type, " & _
                 "Date_capture , Bus_code, Passenger, Fare_Rate " & _
                 "from TripSummaryTbl)a " & _
                "Left Join " & _
                "(select Origin + ' | ' + Destination as UniqueID_1, Estimated_Distance from ServiceTbl WHERE Route_Num <> 'XXXXX')b " & _
                "on a.UniqueID = b.UniqueID_1 " & _
                "where CONCAT(FORMAT(Date_capture,'MMM'), '-',datepart(yy,Date_capture) ) = " & "'" & mth & "'" & _
                "group by  Origin, Destination, Trip_Type " & _
                "order by Trip_Type"
    
    
    With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

'open a new workbook

Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
Application.DisplayAlerts = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

    ActiveWorkbook.SaveAs xpath & "\Loading_Station_Report_" & mth & ".xlsx"
    Dim workbookname As String
    workbookname = "Loading_Station_Report_" & mth & ".xlsx"
        With Application.Workbooks(workbookname).Worksheets("Sheet1")
            .Range("A1").Value = "Origin"
            .Range("B1").Value = "Destination"
            .Range("C1").Value = "NO OF TRIPS"
            .Range("D1").Value = "TRIP TYPE"
            .Range("E1").Value = "PASSENGERS"
            .Range("F1").Value = "REVENUE"
            .Range("G1").Value = "MAIN MILEAGE"
            .Range("H1").Value = "REVENUE RANK"
            .Range("A2").CopyFromRecordset rst
        End With
        Workbooks(workbookname).Sheets(1).Columns.AutoFit
        
        Workbooks(workbookname).Save
        Workbooks(workbookname).Close
    
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Loading Station Summary report created successfully, Please check the path " & xpath
    rst.Close
    conn.Close
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub


Public Function FormatDate(DatePassed As Variant) As String
'  0 = month-day-year;   1 = day-month-year;   2 = year-month-day
If Application.International(xlDateOrder) = 0 Then
    FormatDate = CDate(VBA.DateSerial(Left(VBA.Format(DatePassed, "YYYY-MM-DD"), 4), Mid(VBA.Format(DatePassed, "YYYY-MM-DD"), 6, 2), Right(VBA.Format(DatePassed, "YYYY-MM-DD"), 2)))
    FormatDate = VBA.Format(FormatDate, "YYYY-MM-DD")
ElseIf Application.International(xlDateOrder) = 1 Then
    FormatDate = CDate(VBA.DateSerial(Left(VBA.Format(DatePassed, "YYYY-MM-DD"), 4), Mid(VBA.Format(DatePassed, "YYYY-MM-DD"), 6, 2), Right(VBA.Format(DatePassed, "YYYY-MM-DD"), 2)))
    FormatDate = VBA.Format(FormatDate, "YYYY-MM-DD")
Else
    FormatDate = CDate(VBA.DateSerial(Right(DatePassed, 4), Mid(DatePassed, 6, 2), Left(DatePassed, 2)))
    FormatDate = VBA.Format(FormatDate, "YYYY-MM-DD")
End If

End Function

Sub Multisearch()
Attribute Multisearch.VB_ProcData.VB_Invoke_Func = "m\n14"
    Frm_Search_List.Show
End Sub


Sub Load_Data_Staging()
    Sheet5.Shapes("Btn_Load_Data").Visible = msoFalse
    Sheet5.Shapes("Btn_Rollback_Data").Visible = msoCTrue
    Sheet5.Shapes("Btn_Load").Visible = msoCTrue
    Sheet5.Shapes("Btn_Rollback").Visible = msoFalse
End Sub

Sub Rollback_Data_Staging()
    Sheet5.Shapes("Btn_Load_Data").Visible = msoCTrue
    Sheet5.Shapes("Btn_Rollback_Data").Visible = msoFalse
    Sheet5.Shapes("Btn_Rollback").Visible = msoCTrue
    Sheet5.Shapes("Btn_Load").Visible = msoFalse
End Sub



Sub Rollback_Data_From_Staging()
    Dim lrow, i As Long, val As Long
    Dim mystring As String
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query_1, new_query_2, new_query_3, new_query_4, new_query_5, new_query_6 As String
    Dim new_query_1_1, new_query_2_2, new_query_3_3, new_query_4_4, new_query_5_5, new_query_6_6 As String
    Dim PrevDay As String
    
    lrow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
    If lrow > 5 Then
        Sheet5.Range("A6:C" & lrow).ClearContents
        Sheet5.Range("E6:L" & lrow).ClearContents
        Sheet5.Range("N6:R" & lrow).ClearContents
    End If

    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    If FrmRollBack.OptionButtonBatch.Value = True Then
        val = Application.InputBox("Please enter the batch to pull", "Enter Batch Number", Type:=1)
    End If
    
    Call RemoveDuplicates
    
    ' -------Insert Query--------------
    new_query_1 = "SELECT DISTINCT * INTO TempTbl " & _
                "FROM " & _
                "(SELECT Top(" & val & ") " & "Trip_Number, User_No, Bus_code, Date_capture, Route_No, Arrived_from, Arrival_Time, " & _
                "Depart_Time , Origin, Destination, Shift_time, Captain_Name, Passenger, Trip_ID, StationID, Time_of_First_Ticket, Time_of_Last_Ticket, Fare_Rate, Date_Entered " & _
                "From " & _
                "(" & _
                    "SELECT ROW_NUMBER() OVER (Order by Date_Entered) AS RowID,* " & _
                    "From Staging_table " & _
                    "WHERE User_No in ('" & FrmRollBack.ComboBox_Users.Value & "_01')" & _
                ") as LastEnteredTbl order by Arrival_Time, Origin, Destination asc) as LonTbl"
    

                
    ' ---------------delete query---------------------
    
    new_query_1_1 = "DELETE LonTbl FROM " & _
                "(SELECT Top(" & val & ") " & "Trip_Number, User_No, Bus_code, Date_capture, Route_No, Arrived_from, Arrival_Time, " & _
                "Depart_Time , Origin, Destination, Shift_time, Captain_Name, Passenger, Trip_ID, StationID, Time_of_First_Ticket, Time_of_Last_Ticket, Fare_Rate, Date_Entered " & _
                "From " & _
                "(" & _
                    "SELECT ROW_NUMBER() OVER (Order by Date_Entered) AS RowID,* " & _
                    "From Staging_table " & _
                    "WHERE User_No in ('" & FrmRollBack.ComboBox_Users.Value & "_01')" & _
                ") as LastEnteredTbl order by Arrival_Time asc) as LonTbl"
        
    
        conn.Execute new_query_1
        conn.Execute new_query_1_1
   
     sql_query = "Select * From TempTbl"
     
    rst.Open sql_query, conn, adOpenKeyset, adLockOptimistic
    
    If rst.RecordCount < 1 Then
        MsgBox "No records available for retrieval"
    End If
    
    For i = 0 To rst.RecordCount - 1
        Sheet5.Range("A6").Offset(i, 0).Value = Right(rst.Fields(0).Value, VBA.Len(rst.Fields(0).Value) - VBA.InStr(1, rst.Fields(0).Value, "|"))
        Sheet5.Range("A6").Offset(i, 1).Value = rst.Fields(2).Value
        Sheet5.Range("A6").Offset(i, 2).Value = rst.Fields(3).Value
        Sheet5.Range("A6").Offset(i, 4).Value = rst.Fields(5).Value
        Sheet5.Range("A6").Offset(i, 5).Value = rst.Fields(6).Value
        Sheet5.Range("A6").Offset(i, 6).Value = rst.Fields(7).Value
        Sheet5.Range("A6").Offset(i, 7).Value = rst.Fields(8).Value
        Sheet5.Range("A6").Offset(i, 8).Value = rst.Fields(9).Value
        Sheet5.Range("A6").Offset(i, 9).Value = rst.Fields(10).Value
        Sheet5.Range("A6").Offset(i, 10).Value = rst.Fields(11).Value
        If rst.Fields(13).Value = 1 Then
            Sheet5.Range("A6").Offset(i, 11).Value = "Main Trip"
        ElseIf rst.Fields(13).Value = 2 Then
            Sheet5.Range("A6").Offset(i, 11).Value = "Sub Trip"
        Else
            Sheet5.Range("A6").Offset(i, 11).Value = "Dead Trip"
        End If
        Sheet5.Range("A6").Offset(i, 13).Value = rst.Fields(12).Value
        Sheet5.Range("A6").Offset(i, 14).Value = rst.Fields(15).Value
        Sheet5.Range("A6").Offset(i, 15).Value = rst.Fields(16).Value
        Sheet5.Range("A6").Offset(i, 16).Value = rst.Fields(1).Value
        Sheet5.Range("A6").Offset(i, 17).Value = Left(rst.Fields(0).Value, VBA.InStr(1, rst.Fields(0).Value, "|") - 1)
        rst.MoveNext
    Next i
    MsgBox rst.RecordCount & " records retrieved successfully"
    conn.Execute "Drop Table TempTbl"
    rst.Close
    conn.Close
    
    Call Sort_Macro
    
    Exit Sub
    
ErrorHandler_2:
MsgBox Err.Description


End Sub


Sub Load_Users_listbox()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim mth As String
    Dim i As Long

    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    On Error GoTo ErrorHandler_2
    
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    
    sql_query = "Select Username from LoginTbl Where Username NOT IN ('Watchman', 'Admin01', 'Kolawole.Alakija')"
    
    rst.Open sql_query, conn, adOpenKeyset, adLockPessimistic
    
    For i = 1 To rst.RecordCount
        FrmRollBack.ComboBox_Users.AddItem rst.Fields(0).Value
        rst.MoveNext
    Next i
    rst.Close
    conn.Close
    Exit Sub
ErrorHandler_2:
MsgBox Err.Description
End Sub

Sub Load_Rollback_form()
    FrmRollBack.Show
End Sub

Sub Data_Entry_3()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim lrow, i, j As Long
    Dim message As Variant
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim trip_id As String
    
    Sheet5.Range("A5").Select
    
    j = 0
    
    lrow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
    'checks
    message = MsgBox("Do you wish to load all your entries now?", vbYesNo)
    If message = vbYes Then
        'Do Nothing
    Else
        Exit Sub
    End If
    For i = 6 To lrow
        If Sheet5.Range("A" & i).Offset(0, 5).Value > Sheet5.Range("A" & i).Offset(0, 6).Value Then
            MsgBox "Arrival time can't be more than departed time" & vbCrLf & "Check Row A" & i
            Exit Sub
        End If
        
        If Sheet5.Range("A" & i).Offset(0, 14).Value > Sheet5.Range("A" & i).Offset(0, 15).Value Then
            MsgBox "Last ticket time can't be less than first ticket time" & vbCrLf & "Check Row A" & i
            Exit Sub
    End If
    
    Next i
    
    On Error GoTo Err1
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = "SELECT * FROM TripSummaryTbl"
    
        With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
        For i = 6 To lrow
            .AddNew
            trip_id = Sheet5.Range("A" & i).Offset(0, 17).Value & "|" & Format(Sheet5.Range("A" & i).Value, "0000000")
            .Fields(1).Value = trip_id
            .Fields(2).Value = Sheet5.Range("A" & i).Offset(0, 16).Value
            .Fields(3).Value = Sheet5.Range("A" & i).Offset(0, 1).Value
            .Fields(4).Value = FormatDate(Sheet5.Range("A" & i).Offset(0, 2).Value)
            .Fields(5).Value = Sheet5.Range("A" & i).Offset(0, 3).Value
            .Fields(6).Value = Sheet5.Range("A" & i).Offset(0, 4).Value
            .Fields(7).Value = Format(Sheet5.Range("A" & i).Offset(0, 5).Value, "hh:mm:ss AM/PM")
            .Fields(8).Value = Format(Sheet5.Range("A" & i).Offset(0, 6).Value, "hh:mm:ss AM/PM")
            .Fields(9).Value = Sheet5.Range("A" & i).Offset(0, 7).Value
            .Fields(10).Value = Sheet5.Range("A" & i).Offset(0, 8).Value
            .Fields(11).Value = Sheet5.Range("A" & i).Offset(0, 9).Value
            .Fields(12).Value = Sheet5.Range("A" & i).Offset(0, 10).Value
            .Fields(13).Value = Sheet5.Range("A" & i).Offset(0, 13).Value
            If Sheet5.Range("A" & i).Offset(0, 11).Value = "Main Trip" Then
                .Fields(14).Value = 1
            ElseIf Sheet5.Range("A" & i).Offset(0, 11).Value = "Sub Trip" Then
                .Fields(14).Value = 2
            Else
                .Fields(14).Value = 3
            End If
            .Fields(15).Value = Sheet5.Range("A" & i).Offset(0, 12).Value
            .Fields(16).Value = Format(Sheet5.Range("A" & i).Offset(0, 14).Value, "hh:mm:ss AM/PM")
            .Fields(17).Value = Format(Sheet5.Range("A" & i).Offset(0, 15).Value, "hh:mm:ss AM/PM")
            .Fields(18).Value = Application.WorksheetFunction.VLookup(Sheet5.Range("A" & i).Offset(0, 7).Value & "|" & Sheet5.Range("A" & i).Offset(0, 8).Value, Sheet3.Range("H2:N150"), 7, 0)
            .Fields(19).Value = FormatDate(VBA.Date())
            .Update
            j = j + 1
            Application.StatusBar = j & " records out of " & lrow - 5 & " loaded successfully"
        Next i
    End With
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing
    
    Call Audit_Trail_2
    If lrow > 5 Then
        Sheet5.Range("A6:C" & lrow).ClearContents
        Sheet5.Range("E6:L" & lrow).ClearContents
        Sheet5.Range("N6:R" & lrow).ClearContents
    End If
    
    MsgBox "Data Successfully loaded"
    Application.StatusBar = ""
    'Sheet5.Range("M6").Formula = "=IFERROR(INDEX(Sheet3!$Q$2:$Q$23,MATCH(H6," & "Origin_List,0)),"""")"
    'Sheet5.Range("M6").AutoFill Sheet5.Range("M6:M10000")
Exit Sub

Err1:
MsgBox Err.Description
rst.Close
conn.Close
Set rst = Nothing
Set conn = Nothing
    
End Sub
'==================================================
' **************** End of Code *******************
'==================================================

Sub RemoveDuplicates()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim PrevDay As String
    
    
    Set conn = New ADODB.Connection
    
    'On Error GoTo ErrorHandler_2
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = Sheet1.Range("L22").Value
    
    conn.Execute sql_query
    
    Exit Sub

ErrorHandler_2:
MsgBox Err.Description
conn.Close
End Sub

Sub LoadCaptain()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim PrevDay As String
    
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    Sheet3.Range("Captains_List").ClearContents
    
    On Error GoTo ErrorHandler_2
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = "SELECT Captains_Name FROM CaptainsTbl ORDER BY 1 ASC"
    
    rst.Open sql_query, conn, adOpenKeyset, adLockOptimistic
    
    Sheet3.Range("U2").CopyFromRecordset rst
    
    Exit Sub

ErrorHandler_2:
MsgBox Err.Description
conn.Close
End Sub

Sub LoadUsers()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim PrevDay As String
    
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Not IsEmpty(Sheet3.Range("AB2").Value) Then
        Sheet3.Range("Users").ClearContents
        Sheet3.Range("Usernames").ClearContents
    End If
    
    Sheet3.Range("AB2").Value = "All"
    
    On Error GoTo ErrorHandler_2
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = "SELECT UserID, Username FROM UsersTbl WHERE UserID Not in ('Admin01_01', 'Watchman_01','kolawole.alakija_01', 'taiwo.olusola_01', 'babatunde.obadina_01') ORDER BY 1 ASC"
    
    rst.Open sql_query, conn, adOpenKeyset, adLockOptimistic
    
    Sheet3.Range("AB3").CopyFromRecordset rst
    
    Exit Sub

ErrorHandler_2:
MsgBox Err.Description
conn.Close
End Sub

Sub Data_Entry_4()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim lrow, i, j As Long
    Dim message As Variant
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim TripID As String
    
    Sheet6.Range("A5").Select
    
    j = 0
    
    lrow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    
    'checks
    message = MsgBox("Do you wish to load all your entries now?", vbYesNo)
    If message = vbYes Then
        'Do Nothing
    Else
        Exit Sub
    End If
    For i = 6 To lrow
        If Sheet6.Range("A" & i).Offset(0, 5).Value > Sheet6.Range("A" & i).Offset(0, 6).Value Then
            MsgBox "Arrival time can't be more than departed time" & vbCrLf & "Check Row A" & i
            Exit Sub
        End If
        
        If Sheet6.Range("A" & i).Offset(0, 14).Value > Sheet6.Range("A" & i).Offset(0, 15).Value Then
            MsgBox "Last ticket time can't be less than first ticket time" & vbCrLf & "Check Row A" & i
            Exit Sub
    End If
    
    Next i
    
    On Error GoTo Err1
      connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    sql_query = "SELECT * FROM Staging_table_V2"
    
        With rst
        .ActiveConnection = conn
        .Source = sql_query
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
        For i = 6 To lrow
            TripID = mTest.CreateGUID()
            .AddNew
            .Fields(0).Value = TripID & "|" & Sheet6.Range("A" & i).Value
            .Fields(1).Value = Sheet2.Range("B3").Value
            .Fields(2).Value = Sheet6.Range("A" & i).Offset(0, 1).Value
            .Fields(3).Value = FormatDate(Sheet6.Range("A" & i).Offset(0, 2).Value)
            .Fields(4).Value = Sheet6.Range("A" & i).Offset(0, 3).Value
            .Fields(5).Value = Sheet6.Range("A" & i).Offset(0, 4).Value
            .Fields(6).Value = Format(Sheet6.Range("A" & i).Offset(0, 5).Value, "hh:mm:ss AM/PM")
            .Fields(7).Value = Format(Sheet6.Range("A" & i).Offset(0, 6).Value, "hh:mm:ss AM/PM")
            .Fields(8).Value = Sheet6.Range("A" & i).Offset(0, 7).Value
            .Fields(9).Value = Sheet6.Range("A" & i).Offset(0, 8).Value
            .Fields(10).Value = Sheet6.Range("A" & i).Offset(0, 9).Value
            .Fields(11).Value = Sheet6.Range("A" & i).Offset(0, 10).Value
            .Fields(12).Value = Sheet6.Range("A" & i).Offset(0, 13).Value
            If Sheet6.Range("A" & i).Offset(0, 11).Value = "Main Trip" Then
                .Fields(13).Value = 1
            ElseIf Sheet6.Range("A" & i).Offset(0, 11).Value = "Sub Trip" Then
                .Fields(13).Value = 2
            Else
                .Fields(13).Value = 3
            End If
            .Fields(14).Value = Sheet6.Range("A" & i).Offset(0, 12).Value
            .Fields(15).Value = Format(Sheet6.Range("A" & i).Offset(0, 14).Value, "hh:mm:ss AM/PM")
            .Fields(16).Value = Format(Sheet6.Range("A" & i).Offset(0, 15).Value, "hh:mm:ss AM/PM")
            .Fields(17).Value = Application.WorksheetFunction.VLookup(Sheet6.Range("A" & i).Offset(0, 7).Value & "|" & Sheet6.Range("A" & i).Offset(0, 8).Value, Sheet3.Range("H2:N150"), 7, 0)
            .Fields(18).Value = FormatDate(VBA.Date())
            .Update
            j = j + 1
            Application.StatusBar = j & " records out of " & lrow - 5 & " loaded successfully"
        Next i
    End With
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing
    Application.StatusBar = ""
Exit Sub

Err1:
MsgBox Err.Description
rst.Close
conn.Close
Set rst = Nothing
Set conn = Nothing

End Sub
'==================================================
' **************** End of Code *******************
'==================================================


Sub Time_Sheet()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query As String
    Dim i As Long
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    'On Error GoTo ErrorHandler_2
    If Not IsEmpty(Sheet7.Range("B10")) Then
        Sheet7.Range("B10:B" & Cells(Rows.Count, 2).End(xlUp).Row).EntireRow.Delete
    End If
    
    connection_string = Connection_MSSQL_Production
    
    conn.Open connection_string
    
    conn.Execute "Drop Table TempTbl_2"
    
    If Sheet7.Range("C3").Value = "All" Then

        sql_query = "select * into TempTbl_2 from" & _
                    "(select Date_Entered, User_No, SUM(Passenger)[# of PAX], COUNT(User_No) [# of entries], FORMAT(SUM(Passenger * Fare_Rate),'#,##0.00') Revenue " & _
                    "From Staging_table_V2 " & _
                    "where Date_Entered between '" & Format(Sheet7.Range("C5").Value, "YYYY-MM-DD") & "' and '" & Format(Sheet7.Range("C6").Value, "YYYY-MM-DD") & "' " & _
                    "group by Date_Entered, User_No) AAA"
    Else
        sql_query = "select * into TempTbl_2 from" & _
                    "(select Date_Entered, User_No, SUM(Passenger)[# of PAX], COUNT(User_No) [# of entries], FORMAT(SUM(Passenger * Fare_Rate),'#,##0.00') Revenue " & _
                    "From Staging_table_V2 " & _
                    "where Date_Entered between '" & Format(Sheet7.Range("C5").Value, "YYYY-MM-DD") & "' and '" & Format(Sheet7.Range("C6").Value, "YYYY-MM-DD") & "' and User_No = '" & Sheet7.Range("C3").Value & "' " & _
                    "group by Date_Entered, User_No) AAA"
    End If
                
    conn.Execute sql_query
    
    With rst
        .ActiveConnection = conn
        .Source = "select * from TempTbl_2"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

    For i = 0 To rst.RecordCount - 1
        Sheet7.Range("B10").Offset(i, 0).Value = rst.Fields(0).Value
        Sheet7.Range("B10").Offset(i, 1).Value = rst.Fields(1).Value
        Sheet7.Range("B10").Offset(i, 4).Value = rst.Fields(2).Value
        Sheet7.Range("B10").Offset(i, 5).Value = rst.Fields(3).Value
        Sheet7.Range("B10").Offset(i, 6).Value = rst.Fields(4).Value
        rst.MoveNext
    Next i
    rst.Close
    conn.Close
    
Exit Sub

ErrorHandler_2:
MsgBox Err.Description
End Sub


