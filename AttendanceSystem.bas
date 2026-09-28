Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' EXCEL VBA ATTENDANCE MANAGEMENT SYSTEM
' Version 1.0
'
' Main features:
'   - Employee Master
'   - Employee ID dropdown
'   - XLOOKUP-based employee information
'   - Clock In
'   - Clock Out
'   - Duplicate Clock In prevention
'   - Clock Out validation
'   - Late detection
'   - Working hours calculation
'   - Overtime calculation
'   - Attendance Log
'   - Dashboard KPIs
' ============================================================


' ============================================================
' 1. MAIN SETUP
' ============================================================

Public Sub SetupAttendanceSystem()

    Dim wsDash As Worksheet
    Dim wsEmp As Worksheet
    Dim wsLog As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo SetupError

    ' Get or create required sheets
    Set wsDash = GetOrCreateSheet("Dashboard", "Sheet1")
    Set wsLog = GetOrCreateSheet("Attendance_Log", "Sheet2")
    Set wsEmp = GetOrCreateSheet("Employees", "")

    ' Clean existing content
    PrepareDashboard wsDash
    PrepareEmployees wsEmp
    PrepareAttendanceLog wsLog

    ' Create employee dropdown
    CreateEmployeeValidation wsDash

    ' Build dashboard
    BuildDashboard wsDash

    ' Create buttons
    CreateDashboardButtons wsDash

    ' Create formulas / calculations
    CreateDashboardKPIs wsDash

    ' Final formatting
    FormatDashboard wsDash
    FormatEmployees wsEmp
    FormatAttendanceLog wsLog

    ' Refresh formulas
    Application.CalculateFull

    ' Activate Dashboard
    wsDash.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Attendance Management System has been created successfully!" & vbCrLf & vbCrLf & _
           "You can now:" & vbCrLf & _
           "1. Select an Employee ID" & vbCrLf & _
           "2. Click Clock In" & vbCrLf & _
           "3. Click Clock Out when leaving" & vbCrLf & _
           "4. View records in Attendance_Log", _
           vbInformation, "Setup Complete"

    Exit Sub

SetupError:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Setup failed." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Setup Error"

End Sub


' ============================================================
' 2. GET OR CREATE WORKSHEET
' ============================================================

Private Function GetOrCreateSheet(ByVal targetName As String, _
                                  ByVal legacyName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set GetOrCreateSheet = ws
        Exit Function
    End If

    ' Try to rename an old sheet
    If legacyName <> "" Then

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(legacyName)
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Name = targetName
            Set GetOrCreateSheet = ws
            Exit Function
        End If

    End If

    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = targetName

    Set GetOrCreateSheet = ws

End Function


' ============================================================
' 3. PREPARE DASHBOARD
' ============================================================

Private Sub PrepareDashboard(ByVal ws As Worksheet)

    ws.Cells.Clear

    On Error Resume Next
    ws.Shapes.Delete
    On Error GoTo 0

    ws.Range("A1:H25").ClearFormats

End Sub


' ============================================================
' 4. PREPARE EMPLOYEES SHEET
' ============================================================

Private Sub PrepareEmployees(ByVal ws As Worksheet)

    Dim tbl As ListObject

    ws.Cells.Clear

    ' Remove existing tables
    On Error Resume Next

    For Each tbl In ws.ListObjects
        tbl.Delete
    Next tbl

    On Error GoTo 0

    ' Headers
    ws.Range("A1:E1").Value = Array( _
        "Employee ID", _
        "Employee Name", _
        "Department", _
        "Shift Start", _
        "Shift End")

    ' Demo employees
    ws.Range("A2:E2").Value = Array( _
        "EMP001", _
        "Employee 001", _
        "Finance", _
        TimeValue("08:30"), _
        TimeValue("17:00"))

    ws.Range("A3:E3").Value = Array( _
        "EMP002", _
        "Employee 002", _
        "IT", _
        TimeValue("09:00"), _
        TimeValue("17:30"))

    ws.Range("A4:E4").Value = Array( _
        "EMP003", _
        "Employee 003", _
        "HR", _
        TimeValue("08:30"), _
        TimeValue("17:00"))

    ws.Range("A5:E5").Value = Array( _
        "EMP004", _
        "Employee 004", _
        "Finance", _
        TimeValue("09:00"), _
        TimeValue("17:30"))

    ' Create table
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("A1:E5"), _
        XlListObjectHasHeaders:=xlYes)

    tbl.Name = "tblEmployees"
    tbl.TableStyle = "TableStyleMedium2"

    ' Formats
    ws.Range("D2:E1000").NumberFormat = "hh:mm AM/PM"

End Sub


' ============================================================
' 5. PREPARE ATTENDANCE LOG
' ============================================================

Private Sub PrepareAttendanceLog(ByVal ws As Worksheet)

    Dim tbl As ListObject

    ws.Cells.Clear

    ' Remove old tables
    On Error Resume Next

    For Each tbl In ws.ListObjects
        tbl.Delete
    Next tbl

    On Error GoTo 0

    ' Headers
    ws.Range("A1:I1").Value = Array( _
        "Employee ID", _
        "Employee Name", _
        "Department", _
        "Date", _
        "Time In", _
        "Time Out", _
        "Hours", _
        "Status", _
        "Overtime")

    ' Create empty table
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("A1:I2"), _
        XlListObjectHasHeaders:=xlYes)

    tbl.Name = "tblAttendance"
    tbl.TableStyle = "TableStyleMedium4"

    ' Clear the first empty row
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
    End If

End Sub


' ============================================================
' 6. BUILD DASHBOARD
' ============================================================

Private Sub BuildDashboard(ByVal ws As Worksheet)

    With ws

        ' ----------------------------------------------------
        ' Title
        ' ----------------------------------------------------

        .Range("A1:H1").Merge
        .Range("A1").Value = "ATTENDANCE MANAGEMENT SYSTEM"

        .Range("A2:H2").Merge
        .Range("A2").Value = "Employee Attendance & Working Hours"

        ' ----------------------------------------------------
        ' Employee section
        ' ----------------------------------------------------

        .Range("A4").Value = "Employee ID"
        .Range("A5").Value = "Employee Name"
        .Range("A6").Value = "Department"
        .Range("A7").Value = "Current Date"
        .Range("A8").Value = "Current Time"

        ' Employee ID input
        .Range("B4").Value = ""

        ' Employee Name
        .Range("B5").Formula = _
            "=IFERROR(XLOOKUP(B4,Employees!A:A,Employees!B:B,""""),"""")"

        ' Department
        .Range("B6").Formula = _
            "=IFERROR(XLOOKUP(B4,Employees!A:A,Employees!C:C,""""),"""")"

        ' Current date
        .Range("B7").Formula = "=TODAY()"

        ' Current time
        .Range("B8").Formula = "=NOW()"

        ' ----------------------------------------------------
        ' System status
        ' ----------------------------------------------------

        .Range("A12").Value = "System Status"
        .Range("B12").Value = "Ready"

        ' ----------------------------------------------------
        ' KPI section
        ' ----------------------------------------------------

        .Range("D4").Value = "TOTAL EMPLOYEES"
        .Range("D5").Value = "PRESENT TODAY"
        .Range("D6").Value = "LATE TODAY"
        .Range("D7").Value = "MISSING CLOCK OUT"
        .Range("D8").Value = "TOTAL HOURS TODAY"

        .Range("E4").Formula = "=COUNTA(Employees!A2:A1000)"

        .Range("E5").Formula = _
            "=COUNTIFS(Attendance_Log!D:D,TODAY(),Attendance_Log!E:E,"">0"")"

        .Range("E6").Formula = _
            "=COUNTIFS(Attendance_Log!D:D,TODAY(),Attendance_Log!H:H,""*Late*"")"

        .Range("E7").Formula = _
            "=COUNTIFS(Attendance_Log!D:D,TODAY(),Attendance_Log!E:E,"">0"",Attendance_Log!F:F,"""")"

        .Range("E8").Formula = _
            "=SUMIFS(Attendance_Log!G:G,Attendance_Log!D:D,TODAY())"

        .Range("E8").NumberFormat = "[h]:mm"

        ' ----------------------------------------------------
        ' Instructions
        ' ----------------------------------------------------

        .Range("G4:H4").Merge
        .Range("G4").Value = "QUICK GUIDE"

        .Range("G5:H5").Merge
        .Range("G5").Value = "1. Select Employee ID"

        .Range("G6:H6").Merge
        .Range("G6").Value = "2. Click Clock In"

        .Range("G7:H7").Merge
        .Range("G7").Value = "3. Click Clock Out"

        .Range("G8:H8").Merge
        .Range("G8").Value = "4. Check Attendance Log"

        ' ----------------------------------------------------
        ' Footer
        ' ----------------------------------------------------

        .Range("A21:H21").Merge
        .Range("A21").Value = _
            "Excel VBA Attendance Management System | Portfolio Project"

    End With

End Sub


' ============================================================
' 7. CREATE EMPLOYEE DROPDOWN
' ============================================================

Private Sub CreateEmployeeValidation(ByVal ws As Worksheet)

    Dim formulaText As String

    ' Remove old named range if it exists
    On Error Resume Next
    ThisWorkbook.Names("EmployeeIDList").Delete
    On Error GoTo 0

    ' Dynamic employee ID list
    formulaText = _
        "=Employees!$A$2:INDEX(Employees!$A:$A,MAX(2,COUNTA(Employees!$A:$A)))"

    ThisWorkbook.Names.Add _
        Name:="EmployeeIDList", _
        RefersTo:="=" & Mid$(formulaText, 2)

    ' Remove existing validation
    On Error Resume Next
    ws.Range("B4").Validation.Delete
    On Error GoTo 0

    ' Add dropdown
    With ws.Range("B4").Validation

        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=EmployeeIDList"

        .IgnoreBlank = True
        .InCellDropdown = True

        .InputTitle = "Employee ID"
        .InputMessage = "Select an Employee ID."

        .ErrorTitle = "Invalid Employee ID"
        .ErrorMessage = "Please select an Employee ID from the dropdown."

        .ShowInput = True
        .ShowError = True

    End With

End Sub


' ============================================================
' 8. CREATE DASHBOARD BUTTONS
' ============================================================

Private Sub CreateDashboardButtons(ByVal ws As Worksheet)

    Dim btnClockIn As Shape
    Dim btnClockOut As Shape
    Dim btnRefresh As Shape

    ' --------------------------------------------------------
    ' CLOCK IN BUTTON
    ' --------------------------------------------------------

    Set btnClockIn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("A14").Left, _
        ws.Range("A14").Top, _
        150, _
        45)

    With btnClockIn

        .Name = "btnClockIn"
        .TextFrame2.TextRange.Text = "CLOCK IN"
        .OnAction = "ClockIn"

        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 13

        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Line.ForeColor.RGB = RGB(0, 100, 40)

    End With

    ' --------------------------------------------------------
    ' CLOCK OUT BUTTON
    ' --------------------------------------------------------

    Set btnClockOut = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("C14").Left, _
        ws.Range("C14").Top, _
        150, _
        45)

    With btnClockOut

        .Name = "btnClockOut"
        .TextFrame2.TextRange.Text = "CLOCK OUT"
        .OnAction = "ClockOut"

        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 13

        .Fill.ForeColor.RGB = RGB(192, 0, 0)
        .Line.ForeColor.RGB = RGB(120, 0, 0)

    End With

    ' --------------------------------------------------------
    ' REFRESH BUTTON
    ' --------------------------------------------------------

    Set btnRefresh = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("E14").Left, _
        ws.Range("E14").Top, _
        150, _
        45)

    With btnRefresh

        .Name = "btnRefresh"
        .TextFrame2.TextRange.Text = "REFRESH"
        .OnAction = "RefreshSystem"

        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 13

        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Line.ForeColor.RGB = RGB(35, 75, 140)

    End With

End Sub


' ============================================================
' 9. DASHBOARD KPI FORMULAS
' ============================================================

Private Sub CreateDashboardKPIs(ByVal ws As Worksheet)

    ws.Range("E4:E8").Calculate

End Sub


' ============================================================
' 10. FORMAT DASHBOARD
' ============================================================

Private Sub FormatDashboard(ByVal ws As Worksheet)

    With ws

        ' Font
        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 11

        ' Title
        With .Range("A1:H1")

            .Font.Size = 20
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Interior.Color = RGB(31, 78, 121)
            .Font.Color = RGB(255, 255, 255)

        End With

        .Rows(1).RowHeight = 35

        ' Subtitle
        With .Range("A2:H2")

            .Font.Size = 11
            .Font.Italic = True
            .HorizontalAlignment = xlCenter

        End With

        ' Labels
        With .Range("A4:A8")

            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)

        End With

        ' Input cell
        With .Range("B4")

            .Interior.Color = RGB(255, 242, 204)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter

        End With

        ' Lookup cells
        With .Range("B5:B8")

            .Interior.Color = RGB(242, 242, 242)
            .HorizontalAlignment = xlCenter

        End With

        .Range("B7").NumberFormat = "dd-mmm-yyyy"
        .Range("B8").NumberFormat = "hh:mm:ss AM/PM"

        ' KPI labels
        With .Range("D4:D8")

            .Font.Bold = True
            .Interior.Color = RGB(226, 239, 218)

        End With

        ' KPI values
        With .Range("E4:E8")

            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(242, 242, 242)

        End With

        ' Quick guide
        With .Range("G4:H4")

            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(217, 225, 242)

        End With

        With .Range("G5:H8")

            .HorizontalAlignment = xlLeft
            .Interior.Color = RGB(242, 242, 242)

        End With

        ' Status
        .Range("A12").Font.Bold = True
        .Range("A12").Interior.Color = RGB(221, 235, 247)

        With .Range("B12")

            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(226, 239, 218)

        End With

        ' Footer
        With .Range("A21:H21")

            .Font.Italic = True
            .HorizontalAlignment = xlCenter

        End With

        ' Borders
        With .Range("A4:B8").Borders

            .LineStyle = xlContinuous
            .Weight = xlThin

        End With

        With .Range("D4:E8").Borders

            .LineStyle = xlContinuous
            .Weight = xlThin

        End With

        ' Column widths
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 4
        .Columns("D").ColumnWidth = 22
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 4
        .Columns("G").ColumnWidth = 22
        .Columns("H").ColumnWidth = 18

        ' Hide gridlines
        On Error Resume Next
        ActiveWindow.DisplayGridlines = False
        On Error GoTo 0

    End With

End Sub


' ============================================================
' 11. FORMAT EMPLOYEES SHEET
' ============================================================

Private Sub FormatEmployees(ByVal ws As Worksheet)

    With ws

        .Cells.Font.Name = "Aptos"

        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 18
        .Columns("D:E").ColumnWidth = 15

        .Rows(1).Font.Bold = True

        .Range("D2:E1000").NumberFormat = "hh:mm AM/PM"

    End With

End Sub


' ============================================================
' 12. FORMAT ATTENDANCE LOG
' ============================================================

Private Sub FormatAttendanceLog(ByVal ws As Worksheet)

    With ws

        .Cells.Font.Name = "Aptos"

        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 14
        .Columns("E:F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 20
        .Columns("I").ColumnWidth = 12

        .Range("D2:D1000").NumberFormat = "dd-mmm-yyyy"
        .Range("E2:F1000").NumberFormat = "hh:mm:ss AM/PM"
        .Range("G2:G1000").NumberFormat = "[h]:mm"
        .Range("I2:I1000").NumberFormat = "[h]:mm"

        .Rows(1).Font.Bold = True

        ' Freeze header
        .Activate

        On Error Resume Next
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
        On Error GoTo 0

    End With

End Sub


' ============================================================
' 13. CLOCK IN
' ============================================================

Public Sub ClockIn()

    Dim wsDash As Worksheet
    Dim wsLog As Worksheet
    Dim wsEmp As Worksheet

    Dim employeeID As String
    Dim employeeName As String
    Dim department As String

    Dim shiftStart As Date
    Dim shiftEnd As Date

    Dim nextRow As Long
    Dim openRow As Long

    Dim currentDate As Date
    Dim currentTime As Date

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsLog = ThisWorkbook.Worksheets("Attendance_Log")
    Set wsEmp = ThisWorkbook.Worksheets("Employees")

    ' --------------------------------------------------------
    ' Get Employee ID
    ' --------------------------------------------------------

    employeeID = Trim(CStr(wsDash.Range("B4").Value))

    If employeeID = "" Then

        MsgBox "Please select an Employee ID first.", _
               vbExclamation, "Employee Required"

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' Find employee
    ' --------------------------------------------------------

    If Not GetEmployeeInfo(employeeID, employeeName, department, shiftStart, shiftEnd) Then

        MsgBox "Employee ID not found." & vbCrLf & vbCrLf & _
               "Please select a valid Employee ID.", _
               vbExclamation, "Invalid Employee"

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' Check duplicate open Clock In
    ' --------------------------------------------------------

    openRow = FindOpenAttendanceRow(employeeID, Date)

    If openRow > 0 Then

        MsgBox employeeName & " is already clocked in." & vbCrLf & vbCrLf & _
               "Please use Clock Out before clocking in again.", _
               vbExclamation, "Already Clocked In"

        Exit Sub

    End If

    currentDate = Date
    currentTime = Time

    ' --------------------------------------------------------
    ' Find next row
    ' --------------------------------------------------------

    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    If nextRow < 2 Then nextRow = 2

    ' --------------------------------------------------------
    ' Record attendance
    ' --------------------------------------------------------

    wsLog.Cells(nextRow, "A").Value = employeeID
    wsLog.Cells(nextRow, "B").Value = employeeName
    wsLog.Cells(nextRow, "C").Value = department
    wsLog.Cells(nextRow, "D").Value = currentDate
    wsLog.Cells(nextRow, "E").Value = currentTime
    wsLog.Cells(nextRow, "F").Value = ""
    wsLog.Cells(nextRow, "G").Value = ""
    wsLog.Cells(nextRow, "I").Value = ""

    ' --------------------------------------------------------
    ' Late detection
    ' --------------------------------------------------------

    If currentTime > shiftStart Then

        wsLog.Cells(nextRow, "H").Value = "Late - Checked In"

    Else

        wsLog.Cells(nextRow, "H").Value = "Checked In"

    End If

    ' --------------------------------------------------------
    ' Formatting
    ' --------------------------------------------------------

    wsLog.Cells(nextRow, "D").NumberFormat = "dd-mmm-yyyy"
    wsLog.Cells(nextRow, "E").NumberFormat = "hh:mm:ss AM/PM"

    ' --------------------------------------------------------
    ' Dashboard update
    ' --------------------------------------------------------

    wsDash.Range("B12").Value = "Clock In recorded for " & employeeName

    Application.CalculateFull

    MsgBox "Clock In recorded successfully!" & vbCrLf & vbCrLf & _
           "Employee: " & employeeName & vbCrLf & _
           "Date: " & Format(currentDate, "dd-mmm-yyyy") & vbCrLf & _
           "Time: " & Format(currentTime, "hh:mm:ss AM/PM"), _
           vbInformation, "Clock In"

End Sub


' ============================================================
' 14. CLOCK OUT
' ============================================================

Public Sub ClockOut()

    Dim wsDash As Worksheet
    Dim wsLog As Worksheet

    Dim employeeID As String
    Dim employeeName As String
    Dim department As String

    Dim shiftStart As Date
    Dim shiftEnd As Date

    Dim openRow As Long

    Dim timeIn As Date
    Dim timeOut As Date

    Dim workingHours As Double
    Dim scheduledHours As Double
    Dim overtimeHours As Double

    Dim finalStatus As String

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsLog = ThisWorkbook.Worksheets("Attendance_Log")

    ' --------------------------------------------------------
    ' Get Employee ID
    ' --------------------------------------------------------

    employeeID = Trim(CStr(wsDash.Range("B4").Value))

    If employeeID = "" Then

        MsgBox "Please select an Employee ID first.", _
               vbExclamation, "Employee Required"

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' Find employee
    ' --------------------------------------------------------

    If Not GetEmployeeInfo(employeeID, employeeName, department, shiftStart, shiftEnd) Then

        MsgBox "Employee ID not found.", _
               vbExclamation, "Invalid Employee"

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' Find open record
    ' --------------------------------------------------------

    openRow = FindOpenAttendanceRow(employeeID, Date)

    If openRow = 0 Then

        MsgBox "No active Clock In was found for " & employeeName & "." & vbCrLf & vbCrLf & _
               "Please Clock In before using Clock Out.", _
               vbExclamation, "Clock In Required"

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' Time values
    ' --------------------------------------------------------

    timeIn = wsLog.Cells(openRow, "E").Value
    timeOut = Time

    ' --------------------------------------------------------
    ' Working hours
    ' --------------------------------------------------------

    workingHours = timeOut - timeIn

    ' Handle overnight shifts
    If workingHours < 0 Then
        workingHours = workingHours + 1
    End If

    ' --------------------------------------------------------
    ' Scheduled hours
    ' --------------------------------------------------------

    scheduledHours = shiftEnd - shiftStart

    ' Handle overnight shift
    If scheduledHours < 0 Then
        scheduledHours = scheduledHours + 1
    End If

    ' --------------------------------------------------------
    ' Overtime
    ' --------------------------------------------------------

    overtimeHours = workingHours - scheduledHours

    If overtimeHours < 0 Then
        overtimeHours = 0
    End If

    ' --------------------------------------------------------
    ' Status
    ' --------------------------------------------------------

    If timeIn > shiftStart Then

        finalStatus = "Late - Completed"

    Else

        finalStatus = "Completed"

    End If

    ' --------------------------------------------------------
    ' Update record
    ' --------------------------------------------------------

    wsLog.Cells(openRow, "F").Value = timeOut
    wsLog.Cells(openRow, "G").Value = workingHours
    wsLog.Cells(openRow, "H").Value = finalStatus
    wsLog.Cells(openRow, "I").Value = overtimeHours

    ' --------------------------------------------------------
    ' Formatting
    ' --------------------------------------------------------

    wsLog.Cells(openRow, "F").NumberFormat = "hh:mm:ss AM/PM"
    wsLog.Cells(openRow, "G").NumberFormat = "[h]:mm"
    wsLog.Cells(openRow, "I").NumberFormat = "[h]:mm"

    ' --------------------------------------------------------
    ' Dashboard status
    ' --------------------------------------------------------

    wsDash.Range("B12").Value = "Clock Out recorded for " & employeeName

    Application.CalculateFull

    MsgBox "Clock Out recorded successfully!" & vbCrLf & vbCrLf & _
           "Employee: " & employeeName & vbCrLf & _
           "Time Out: " & Format(timeOut, "hh:mm:ss AM/PM") & vbCrLf & _
           "Working Hours: " & Format(workingHours, "[h]:mm") & vbCrLf & _
           "Overtime: " & Format(overtimeHours, "[h]:mm"), _
           vbInformation, "Clock Out"

End Sub


' ============================================================
' 15. FIND EMPLOYEE INFORMATION
' ============================================================

Private Function GetEmployeeInfo(ByVal employeeID As String, _
                                 ByRef employeeName As String, _
                                 ByRef department As String, _
                                 ByRef shiftStart As Date, _
                                 ByRef shiftEnd As Date) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Employees")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow

        If UCase(Trim(CStr(ws.Cells(i, "A").Value))) = _
           UCase(Trim(employeeID)) Then

            employeeName = CStr(ws.Cells(i, "B").Value)
            department = CStr(ws.Cells(i, "C").Value)
            shiftStart = ws.Cells(i, "D").Value
            shiftEnd = ws.Cells(i, "E").Value

            GetEmployeeInfo = True
            Exit Function

        End If

    Next i

    GetEmployeeInfo = False

End Function


' ============================================================
' 16. FIND OPEN ATTENDANCE ROW
' ============================================================

Private Function FindOpenAttendanceRow(ByVal employeeID As String, _
                                       ByVal attendanceDate As Date) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Attendance_Log")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        FindOpenAttendanceRow = 0
        Exit Function
    End If

    ' Search from bottom to top
    For i = lastRow To 2 Step -1

        If UCase(Trim(CStr(ws.Cells(i, "A").Value))) = _
           UCase(Trim(employeeID)) Then

            If IsDate(ws.Cells(i, "D").Value) Then

                If DateValue(ws.Cells(i, "D").Value) = _
                   DateValue(attendanceDate) Then

                    If Trim(CStr(ws.Cells(i, "F").Value)) = "" Then

                        FindOpenAttendanceRow = i
                        Exit Function

                    End If

                End If

            End If

        End If

    Next i

    FindOpenAttendanceRow = 0

End Function


' ============================================================
' 17. REFRESH SYSTEM
' ============================================================

Public Sub RefreshSystem()

    Dim wsDash As Worksheet

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    Application.CalculateFull

    wsDash.Range("B7").Formula = "=TODAY()"
    wsDash.Range("B8").Formula = "=NOW()"

    wsDash.Range("B12").Value = "System refreshed at " & _
                                Format(Now, "hh:mm:ss AM/PM")

    MsgBox "Dashboard refreshed successfully.", _
           vbInformation, "Refresh"

End Sub


' ============================================================
' 18. CLEAR EMPLOYEE SELECTION
' ============================================================

Public Sub ClearEmployeeSelection()

    With ThisWorkbook.Worksheets("Dashboard")

        .Range("B4").ClearContents
        .Range("B12").Value = "Ready"

    End With

End Sub


' ============================================================
' 19. DEMO DATA RESET
' ============================================================

Public Sub ClearAttendanceLog()

    Dim ws As Worksheet
    Dim lastRow As Long

    If MsgBox("Are you sure you want to clear all attendance records?", _
              vbYesNo + vbQuestion, _
              "Clear Attendance Log") <> vbYes Then
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("Attendance_Log")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow >= 2 Then
        ws.Range("A2:I" & lastRow).ClearContents
    End If

    MsgBox "Attendance log cleared.", _
           vbInformation, "Completed"

End Sub


' ============================================================
' 20. TEST SYSTEM
' ============================================================

Public Sub TestSystem()

    Dim wsDash As Worksheet

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    wsDash.Range("B4").Value = "EMP001"

    Application.CalculateFull

    MsgBox "EMP001 has been selected for testing." & vbCrLf & vbCrLf & _
           "Now click Clock In.", _
           vbInformation, "Test Mode"

End Sub


