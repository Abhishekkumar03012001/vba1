Option Explicit

' ===== CONFIG =====
Private Const PL_YEARLY As Double = 18
Private Const SL_YEARLY As Double = 7
Private Const CL_YEARLY As Double = 7
Private Const PL_CAP As Double = 30

' Table names
Private Const EMP_TABLE As String = "tblEmp"
Private Const LV_TABLE As String = "tblLeave"

' Column header names
Private Const EMP_ID As String = "Employee ID"
Private Const EMP_NAME As String = "Name"
Private Const EMP_TEAM As String = "Team"
Private Const EMP_JOIN As String = "Join Date"
Private Const EMP_EXIT As String = "Exit Date"   ' optional column

' ===== RUN REPORT =====
Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As ListObject, wsLv As ListObject
    Dim wsOut As Worksheet
    Dim lastRow As Long, outRow As Long
    Dim r As ListRow
    
    ' Get tables
    Set wsEmp = ThisWorkbook.Sheets(EMP_TABLE).ListObjects(EMP_TABLE)
    Set wsLv = ThisWorkbook.Sheets(LV_TABLE).ListObjects(LV_TABLE)
    
    ' Create output sheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("FilteredData")
    If Not wsOut Is Nothing Then wsOut.Delete
    On Error GoTo 0
    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "FilteredData"
    
    ' Headers
    wsOut.Range("A1:H1").Value = Array("Employee ID", "Name", "Team", "Join Year", _
                                       "PL_Balance", "SL_Balance", "CL_Balance", "Total_Balance")
    outRow = 2
    
    ' Loop employees
    For Each r In wsEmp.ListRows
        Dim empId As Variant, empName As String, empTeam As String
        Dim joinDate As Date, exitDate As Variant
        Dim plBal As Double, slBal As Double, clBal As Double, totalBal As Double
        Dim joinYr As Long
        
        empId = Nz(r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value)
        empName = Nz(r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value)
        empTeam = Nz(r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value)
        joinDate = NzDate(r.Range.Columns(wsEmp.ListColumns(EMP_JOIN).Index).Value)
        joinYr = Year(joinDate)
        
        ' Exit date (optional)
        exitDate = GetExitDate(r, wsEmp)
        
        ' Calculate balances
        Call CalcBalances(joinDate, exitDate, plBal, slBal, clBal)
        totalBal = plBal + slBal + clBal
        
        ' Write output
        With wsOut
            .Cells(outRow, 1).Value = empId
            .Cells(outRow, 2).Value = empName
            .Cells(outRow, 3).Value = empTeam
            .Cells(outRow, 4).Value = joinYr
            .Cells(outRow, 5).Value = plBal
            .Cells(outRow, 6).Value = slBal
            .Cells(outRow, 7).Value = clBal
            .Cells(outRow, 8).Value = totalBal
        End With
        outRow = outRow + 1
    Next r
    
    ' Highlight max/min balances
    lastRow = wsOut.Cells(wsOut.Rows.Count, "H").End(xlUp).Row
    Call HighlightBalances(wsOut, "H2:H" & lastRow)
    
    MsgBox "Report generated successfully!", vbInformation
    Application.ScreenUpdating = True
    Exit Sub
    
ErrH:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ===== HELPER: Calculate Balances =====
Private Sub CalcBalances(joinDate As Date, exitDate As Variant, _
                         ByRef plBal As Double, ByRef slBal As Double, ByRef clBal As Double)
    Dim yr As Long, thisYr As Long
    Dim yearsWorked As Long
    
    thisYr = Year(Date)
    
    If IsDate(joinDate) = False Then Exit Sub
    
    ' If exit date before today, stop at exit year
    Dim endYr As Long
    If IsDate(exitDate) Then
        endYr = Year(exitDate)
    Else
        endYr = thisYr
    End If
    
    yearsWorked = endYr - Year(joinDate) + 1
    
    Dim i As Long
    plBal = 0: slBal = 0: clBal = 0
    
    For i = Year(joinDate) To endYr
        ' Each year accrual
        plBal = WorksheetFunction.Min(PL_CAP, plBal + PL_YEARLY)
        slBal = SL_YEARLY    ' resets each year
        clBal = CL_YEARLY    ' resets each year
    Next i
End Sub

' ===== HELPER: Exit Date (optional column) =====
Private Function GetExitDate(r As ListRow, wsEmp As ListObject) As Variant
    Dim exitDateCol As Long
    On Error Resume Next
    exitDateCol = wsEmp.ListColumns(EMP_EXIT).Index
    On Error GoTo 0
    
    If exitDateCol > 0 Then
        GetExitDate = NzDate(r.Range.Columns(exitDateCol).Value)
    Else
        GetExitDate = Empty
    End If
End Function

' ===== HELPER: Highlight Min/Max =====
Private Sub HighlightBalances(ws As Worksheet, rngAddress As String)
    Dim rng As Range
    Set rng = ws.Range(rngAddress)
    
    Dim maxVal As Double, minVal As Double
    maxVal = Application.Max(rng)
    minVal = Application.Min(rng)
    
    Dim c As Range
    For Each c In rng
        If c.Value = maxVal Then
            c.Interior.Color = vbGreen
        ElseIf c.Value = minVal Then
            c.Interior.Color = vbRed
        Else
            c.Interior.Color = xlNone
        End If
    Next c
End Sub

' ===== HELPERS =====
Private Function Nz(v As Variant, Optional def As Variant = "") As Variant
    If IsError(v) Or IsEmpty(v) Then
        Nz = def
    Else
        Nz = v
    End If
End Function

Private Function NzDate(v As Variant) As Date
    If IsDate(v) Then
        NzDate = CDate(v)
    Else
        NzDate = 0
    End If
End Function

