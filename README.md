Option Explicit

Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As Worksheet, wsLeave As Worksheet, wsFD As Worksheet
    Dim lastRowEmp As Long, lastRowLeave As Long, lastRowFD As Long
    Dim rng As Range, maxLeave As Double
    Dim i As Long, val As Double
    
    ' Source sheets
    Set wsEmp = ThisWorkbook.Sheets("tblEmp")
    Set wsLeave = ThisWorkbook.Sheets("tblLeave")
    
    ' Create/clear FilteredData
    On Error Resume Next
    Set wsFD = ThisWorkbook.Sheets("FilteredData")
    On Error GoTo 0
    If wsFD Is Nothing Then
        Set wsFD = ThisWorkbook.Sheets.Add
        wsFD.Name = "FilteredData"
    Else
        wsFD.Cells.Clear
    End If
    
    ' Headers
    wsFD.Range("A1:H1").Value = Array("emp_id", "name", "team", "join_year", _
                                      "pl_taken", "sl_taken", "cl_taken", "net_leave")
    
    ' Copy data
    lastRowLeave = wsLeave.Cells(wsLeave.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRowLeave
        wsFD.Cells(i, 1).Value = wsLeave.Cells(i, 1).Value ' emp_id
        wsFD.Cells(i, 2).Value = wsEmp.Cells(i, 2).Value   ' name
        wsFD.Cells(i, 3).Value = wsEmp.Cells(i, 3).Value   ' team
        wsFD.Cells(i, 4).Value = Year(wsEmp.Cells(i, 4).Value) ' join_year
        
        wsFD.Cells(i, 5).Value = wsLeave.Cells(i, 2).Value ' pl_taken
        wsFD.Cells(i, 6).Value = wsLeave.Cells(i, 3).Value ' sl_taken
        wsFD.Cells(i, 7).Value = wsLeave.Cells(i, 4).Value ' cl_taken
        
        ' Net Leave
        wsFD.Cells(i, 8).FormulaR1C1 = "=RC5+RC6+RC7"
    Next i
    
    ' Sort by net_leave descending
    lastRowFD = wsFD.Cells(wsFD.Rows.Count, 1).End(xlUp).Row
    Set rng = wsFD.Range("A1:H" & lastRowFD)
    rng.Sort Key1:=wsFD.Range("H2"), Order1:=xlDescending, Header:=xlYes
    
    ' Find max net_leave
    maxLeave = Application.Max(wsFD.Range("H2:H" & lastRowFD))
    
    ' Apply formatting
    For i = 2 To lastRowFD
        val = wsFD.Cells(i, 8).Value
        
        ' Highlight only the value cells
        If val = maxLeave Then
            wsFD.Range("A" & i & ":H" & i).SpecialCells(xlCellTypeConstants).Interior.Color = vbGreen
        End If
        
        If wsFD.Cells(i, 5).Value + wsFD.Cells(i, 6).Value + wsFD.Cells(i, 7).Value < 0 Then
            wsFD.Range("E" & i & ":G" & i).SpecialCells(xlCellTypeConstants).Interior.Color = vbRed
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrH:
    MsgBox "Error: " & Err.Description
    Resume CleanExit
End Sub
