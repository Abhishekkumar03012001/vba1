Option Explicit

' ====== CONFIG ======
Private Const PL_YEARLY As Double = 18
Private Const SL_YEARLY As Double = 7
Private Const CL_YEARLY As Double = 7
Private Const PL_CAP As Double = 30

' Table names
Private Const EMP_TABLE As String = "tblEmp"
Private Const LV_TABLE As String = "tblLeave"

' Column header names (must match exactly in Excel)
Private Const EMP_ID As String = "Employee ID"
Private Const EMP_NAME As String = "Name"
Private Const EMP_TEAM As String = "Team"
Private Const EMP_JOIN As String = "Join Date"
Private Const EMP_EXIT As String = "Exit Date"   ' update if your header differs

' ====== RUN REPORT ======
Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As ListObject, wsLv As ListObject
    Set wsEmp = ThisWorkbook.Sheets(EMP_TABLE).ListObjects(1)
    Set wsLv = ThisWorkbook.Sheets(LV_TABLE).ListObjects(1)
    
    ' Delete old FilteredData if exists
    Dim wsFD As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsFD = ThisWorkbook.Sheets.Add
    wsFD.Name = "FilteredData"
    
    ' Headers
    Dim headers As Variant
    headers = Array("Emp_ID", "Name", "Team", "JoinYear", _
                    "PL_Accrued", "SL_Accrued", "CL_Accrued", _
                    "PL_Balance", "SL_Balance", "CL_Balance", _
                    "Net_Leave_Balance")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        wsFD.Cells(1, i + 1).Value = headers(i)
    Next
    
    ' Loop Employees
    Dim r As ListRow, outRow As Long
    outRow = 2
    
    Dim joinDate As Date, exitDate As Variant
    Dim yearsWorked As Double
    
    For Each r In wsEmp.ListRows
        joinDate = r.Range.Columns(wsEmp.ListColumns(EMP_JOIN).Index).Value
        
        ' Exit Date check
        If ColumnExists(wsEmp, EMP_EXIT) Then
            exitDate = r.Range.Columns(wsEmp.ListColumns(EMP_EXIT).Index).Value
        Else
            exitDate = ""
        End If
        
        If IsDate(exitDate) Then
            yearsWorked = DateDiff("yyyy", joinDate, exitDate)
        Else
            yearsWorked = DateDiff("yyyy", joinDate, Date)
        End If
        If yearsWorked < 0 Then yearsWorked = 0
        
        ' Accrued
        Dim plAcc As Double, slAcc As Double, clAcc As Double
        plAcc = Application.Min(PL_YEARLY * yearsWorked, PL_CAP)
        slAcc = SL_YEARLY * yearsWorked
        clAcc = CL_YEARLY * yearsWorked
        
        ' TODO: subtract taken leave from tblLeave if required
        
        ' Output row
        wsFD.Cells(outRow, 1).Value = r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value
        wsFD.Cells(outRow, 2).Value = r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value
        wsFD.Cells(outRow, 3).Value = r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value
        wsFD.Cells(outRow, 4).Value = Year(joinDate)
        
        wsFD.Cells(outRow, 5).Value = plAcc
        wsFD.Cells(outRow, 6).Value = slAcc
        wsFD.Cells(outRow, 7).Value = clAcc
        
        wsFD.Cells(outRow, 8).Value = plAcc
        wsFD.Cells(outRow, 9).Value = slAcc
        wsFD.Cells(outRow, 10).Value = clAcc
        
        wsFD.Cells(outRow, 11).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"
        
        outRow = outRow + 1
    Next r
    
    ' Color formatting (max green, min red)
    Dim lastRow As Long
    lastRow = wsFD.Cells(wsFD.Rows.Count, "A").End(xlUp).Row
    
    Dim maxBal As Double, minBal As Double
    maxBal = Application.Max(wsFD.Range("K2:K" & lastRow))
    minBal = Application.Min(wsFD.Range("K2:K" & lastRow))
    
    Dim cell As Range
    For Each cell In wsFD.Range("K2:K" & lastRow)
        If cell.Value = maxBal Then cell.Interior.Color = vbGreen
        If cell.Value = minBal Then cell.Interior.Color = vbRed
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "FilteredData created successfully!"
    Exit Sub
    
ErrH:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description
End Sub

' ====== Helper Function ======
Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    On Error Resume Next
    ColumnExists = Not lo.ListColumns(colName) Is Nothing
    On Error GoTo 0
End Function
