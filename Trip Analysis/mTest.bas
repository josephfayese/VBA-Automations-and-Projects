Attribute VB_Name = "mTest"
Option Explicit

Sub check()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim connection_string As String
    Dim sql_query, new_query As String
    Dim cell As Range
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    'server string
    conn.ConnectionString = "Provider=MSOLEDBSQL;" & _
                       "Server=GOP1096502LT\DATAMINE;" & _
                       "Database=Trip_Analytics_DB;" & _
                       "Trusted_Connection=yes;"
    'conn.ConnectionString = "Provider=MSOLEDBSQL;" & _
                                        "Server=192.168.100.6\TRIP_ANALYTICS,1434;" & _
                                        "Database=Trip_Analytics_DB;" & _
                                        "UID=Admin_02;PWD=Admin@lbsl02;"
                       
    'sql_query = "CREATE TABLE TestTbl ( User_ID varchar(20) not null, First_Name varchar(50), Last_Name varchar(50), Email_Address varchar(100))"
    
    '=====================Load Route Table===================
    'Sheet3.Range("B2:B18")
    'sql_query = "Insert into RouteTbl (Route_ID, Route_Name) values('" & cell.Value & "', '" & cell.Offset(0, 1).Value & "')"
    
    '=====================Load Station Table===================
    'Sheet3.Range ("Y2:Y22")
    'sql_query = "Insert into StationTbl (Station) values('" & cell.Value & "')"

    '=====================Load Service Table===================
    'Sheet3.Range("I2:I53")
    'sql_query = "Insert into ServiceTbl (Origin, Destination, Estimated_Distance, Route_Num, Service_Num, Fare)" & _
    "values('" & cell.Value & "', '" & cell.Offset(0, 1).Value & "','" & cell.Offset(0, 2).Value & "','" & _
    cell.Offset(0, 3).Value & "','" & cell.Offset(0, 4).Value & "','" & cell.Offset(0, 5).Value & "')"
                        
    conn.Open connection_string
    
    For Each cell In Sheet3.Range("I2:I53")
        sql_query = "Insert into ServiceTbl (Origin, Destination, Estimated_Distance, Route_Num, Service_Num, Fare)" & _
        "values('" & cell.Value & "', '" & cell.Offset(0, 1).Value & "','" & cell.Offset(0, 2).Value & "','" & _
        cell.Offset(0, 3).Value & "','" & cell.Offset(0, 4).Value & "','" & cell.Offset(0, 5).Value & "')"
        conn.Execute sql_query
    Next cell
    
'    conn.Execute sql_query
    MsgBox "loaded Successfully!!!"
    conn.Close
    
End Sub

Sub test()
Dim xpath As String
Dim strDir As String
Dim strPath As String

Application.screenUpdating = False
strDir = "User_Activity"
strPath = "C:\Users\" & Environ("username") & "\Documents\"

Workbooks.Add

xpath = MkPath(strDir, strPath)

ActiveWorkbook.SaveAs xpath & "\User_Activity_Report" & VBA.Date() & ".xlsx"

ActiveWorkbook.Close

Application.screenUpdating = True

MsgBox "Actvity report created successfully, Please check the path" & xpath

End Sub

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

Sub ReturnMultiples()
Dim c As Range
Dim rng As Range
Dim i As Long

Set rng = Sheet3.Range("Origin_Check")
i = 0
For Each c In rng
    If c.Value = Sheet3.Range("O2").Value Then
        Sheet3.Range("S2").Offset(i, 0).Value = c.Offset(0, 1).Value
        i = i + 1
    End If
Next c
Sheet3.Range("Destination_list").RemoveDuplicates Columns:=1, Header:=xlNo
'If Sheet3.Range("Destination_list").Rows.Count = 1 Then Exit Sub
'Sheet3.Range("Destination_list").Sort Key1:=Range("S2"), Order1:=xlAscending, Header:=xlNo
End Sub


Sub CCCC()

'Dim settings As New clsExcelSettings
'settings.TurnOn
Sheet3.Range("P5").Formula = "=IFERROR(INDEX(Sheet3!Q2:Q23,MATCH(O5,Sheet3!R2:R23,0)),"""")"
End Sub

Public Function CreateGUID() As String
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
End Function


Function GetMyLocalIP() As String
    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myIPAddress     As String
 
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
 
End Function
