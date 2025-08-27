Option Explicit

Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As Worksheet, wsLv As Worksheet, wsFD As Worksheet
    Dim loEmp As ListObject, loLv As ListObject
    Dim rngFD As Range
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, netLeaveCol As Long
    Dim plCol As Long, slCol As Long, clCol As Long
    Dim plTaken As Double, slTaken As Double, clTaken As Double
    
    ' Source tables
    Set wsEmp = ThisWorkbook.Worksheets("Employees")
    Set wsLv = ThisWorkbook.Worksheets("Leaves")
    Set loEmp = wsEmp.ListObjects("tblEmp")
    Set loLv = wsLv.ListObjects("tblLeave")
    
    ' Create / reset FilteredData sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrH
    
    Set wsFD = ThisWorkbook.Worksheets.Add
    wsFD.Name = "FilteredData"
    
    ' Headers
    wsFD.Range("A1").Resize(1, 10).Value = Array("emp_id", "name", "team", "join_year", _
        "pl_balance", "sl_balance", "cl_balance", _
        "pl_taken", "sl_taken", "cl_taken", "net_leave")
    
    ' Populate FilteredData
    Dim r As Long: r = 2
    Dim empRow As ListRow
    For Each empRow In loEmp.ListRows
        Dim empID As Variant
        empID = empRow.Range.Columns(1).Value
        
        ' balances
        Dim plBal As Double, slBal As Double, clBal As Double
        plBal = empRow.Range.Columns(2).Value
        slBal = empRow.Range.Columns(3).Value
        clBal = empRow.Range.Columns(4).Value
        
        ' taken (aggregate from LV table)
        Dim lvRow As ListRow
        Dim plTakenSum As Double, slTakenSum As Double, clTakenSum As Double
        plTakenSum = 0: slTakenSum = 0: clTakenSum = 0
        For Each lvRow In loLv.ListRows
            If lvRow.Range.Columns(1).Value = empID Then
                plTakenSum = plTakenSum + lvRow.Range.Columns(2).Value
                slTakenSum = slTakenSum + lvRow.Range.Columns(3).Value
                clTakenSum = clTakenSum + lvRow.Range.Columns(4).Value
            End If
        Next lvRow
        
        ' net leave
        Dim netLeave As Double
        netLeave = plTakenSum + slTakenSum + clTakenSum
        
        ' write row
        wsFD.Cells(r, 1).Value = empID
        wsFD.Cells(r, 2).Value = empRow.Range.Columns(2).Value ' name
        wsFD.Cells(r, 3).Value = empRow.Range.Columns(3).Value ' team
        wsFD.Cells(r, 4).Value = Year(empRow.Range.Columns(4).Value) ' join_year
        wsFD.Cells(r, 5).Value = plBal
        wsFD.Cells(r, 6).Value = slBal
        wsFD.Cells(r, 7).Value = clBal
        wsFD.Cells(r, 8).Value = plTakenSum
        wsFD.Cells(r, 9).Value = slTakenSum
        wsFD.Cells(r, 10).Value = clTakenSum
        wsFD.Cells(r, 11).Value = netLeave
        
        r = r + 1
    Next empRow
    
    ' Sort by net_leave DESC
    lastRow = wsFD.Cells(wsFD.Rows.Count, "A").End(xlUp).Row
    lastCol = wsFD.Cells(1, wsFD.Columns.Count).End(xlToLeft).Column
    Set rngFD = wsFD.Range("A1").Resize(lastRow, lastCol)
    rngFD.Sort Key1:=wsFD.Range("K2"), Order1:=xlDescending, Header:=xlYes
    
    ' Color formatting
    netLeaveCol = 11
    plCol = 8: slCol = 9: clCol = 10
    
    Dim maxNet As Double
    maxNet = Application.WorksheetFunction.Max(wsFD.Range("K2:K" & lastRow))
    
    For i = 2 To lastRow
        If wsFD.Cells(i, netLeaveCol).Value = maxNet Then
            wsFD.Cells(i, netLeaveCol).Interior.Color = vbGreen
        End If
        
        ' If pl_taken+sl_taken+cl_taken < 0 → red
        plTaken = wsFD.Cells(i, plCol).Value
        slTaken = wsFD.Cells(i, slCol).Value
        clTaken = wsFD.Cells(i, clCol).Value
        If plTaken + slTaken + clTaken < 0 Then
            wsFD.Cells(i, netLeaveCol).Interior.Color = vbRed
        End If
    Next i
    
ExitPoint:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrH:
    MsgBox "Error: " & Err.Description
    Resume ExitPoint
End Sub
