Option Explicit

Sub GenerateDATA2()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As ListObject, wsLv As ListObject
    Set wsEmp = ThisWorkbook.Sheets("Employee Data").ListObjects("EMP_TABLE")
    Set wsLv = ThisWorkbook.Sheets("Leave Status").ListObjects("LV_TABLE")
    
    ' Delete old DATA2 sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("DATA2").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new DATA2 sheet
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "DATA2"
    
    ' Headers
    wsOut.Range("A1:I1").Value = Array("Employee ID", "Name", "Team", "JoinYear", _
                                       "PL_Balance", "SL_Balance", _
                                       "MetricValue", "MetricDetail", "TotalBalance")
    
    ' Build Leave lookup dictionary
    Dim lvDict As Object: Set lvDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, key As String
    For Each lr In wsLv.ListRows
        key = CStr(lr.Range.Cells(1, wsLv.ListColumns("Employee ID").Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Cells(1, wsLv.ListColumns("PL").Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Cells(1, wsLv.ListColumns("SL").Index).Value)
    Next lr
    
    ' Process employees
    Dim r As ListRow, empId, nm, tm, jy
    Dim plBal#, slBal#, totalBal#, metric#, detail As String
    Dim rowOut&: row
