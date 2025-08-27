Option Explicit

Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As ListObject, wsLv As ListObject
    Set wsEmp = ThisWorkbook.Sheets("Employee Data").ListObjects(EMP_TABLE)
    Set wsLv = ThisWorkbook.Sheets("Leave Status").ListObjects(LV_TABLE)
    
    ' Delete old sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "FilteredData"
    
    wsOut.Range("A1:J1").Value = Array("Employee ID","Name","Designation","Team","JoinYear", _
                                       "PL_Balance","SL_Balance","CL_Balance","MetricValue","MetricDetail")
    
    ' build leave lookup dictionary
    Dim lvDict As Object: Set lvDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, key As String
    For Each lr In wsLv.ListRows
        key = CStr(lr.Range.Cells(1, wsLv.ListColumns("Employee ID").Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Cells(1, wsLv.ListColumns("PL").Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Cells(1, wsLv.ListColumns("SL").Index).Value)
        lvDict(key & "|CL") = NzD(lr.Range.Cells(1, wsLv.ListColumns("CL").Index).Value)
    Next lr
    
    Dim r As ListRow, empId, nm, des, tm, jy, doj
    Dim plBal#, slBal#, clBal#
    Dim rowOut&: rowOut = 2
    Dim asOf As Date: asOf = AsOfDateVal()
    
    For Each r In wsEmp.ListRows
        empId = r.Range.Cells(1, wsEmp.ListColumns("Employee ID").Index).Value
        nm = r.Range.Cells(1, wsEmp.ListColumns("Name").Index).Value
        des = r.Range.Cells(1, wsEmp.ListColumns("Designation").Index).Value
        tm = r.Range.Cells(1, wsEmp.ListColumns("Team").Index).Value
        
        ' --- FIXED DOJ / JoinYear extraction ---
        On Error Resume Next
        doj = r.Range.Cells(1, wsEmp.ListColumns("DOJ").Index).Value
        On Error GoTo 0
        
        If IsDate(doj) Then
            jy = Year(doj)
        Else
            jy = NzD(r.Range.Cells(1, wsEmp.ListColumns("JoinYear").Index).Value)
        End If
        ' --- end fix ---
        
        ' filter
        If JoinYearSel <> "All" Then If CStr(jy) <> CStr(JoinYearSel) Then GoTo NextEmp
        If DesigSel <> "All" Then If CStr(des) <> CStr(DesigSel) Then GoTo NextEmp
        If TeamSel <> "All" Then If CStr(tm) <> CStr(TeamSel) Then GoTo NextEmp
        
        plBal = NzD(lvDict(CStr(empId) & "|PL"))
        slBal = NzD(lvDict(CStr(empId) & "|SL"))
        clBal = NzD(lvDict(CStr(empId) & "|CL"))
        
        Dim plAcc#, slAcc#, clAcc#, metric#, detail As String
        plAcc = AccruedPL(asOf, doj)
        slAcc = AccruedSL(asOf, doj)
        clAcc = AccruedCL(asOf, doj)
        
        Dim totalBal#, totalAcc#
        totalBal = plBal + slBal + clBal
        totalAcc = plAcc + slAcc + clAcc
        
        Select Case UCase(LeaveTypeSel())
            Case "PL"
                metric = PickMetric(plBal, plAcc)
                detail = "PL-" & MetricSel()
            Case "SL"
                metric = PickMetric(slBal, slAcc)
                detail = "SL-" & MetricSel()
            Case "CL"
                metric = PickMetric(clBal, clAcc)
                detail = "CL-" & MetricSel()
            Case Else
                metric = PickMetric(totalBal, totalAcc)
                detail = "Total-" & MetricSel()
        End Select
        
        wsOut.Cells(rowOut, 1).Resize(1, 10).Value = _
            Array(empId, nm, des, tm, jy, plBal, slBal, clBal, metric, detail)
        rowOut = rowOut + 1
NextEmp:
    Next r
    
    If rowOut > 2 Then
        wsOut.Range("A1:J" & rowOut - 1).Sort Key1:=wsOut.Range("I2"), Order1:=xlDescending, Header:=xlYes
        Dim keepRows As Long: keepRows = Application.Min(TopNSel, rowOut - 2)
        If rowOut - 2 > keepRows Then wsOut.Rows(keepRows + 2 & ":" & rowOut - 1).Delete
    End If
    
    ApplyBonusFormatting wsOut
    With wsOut.UsedRange
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    
    MsgBox "Report ready: sheet 'FilteredData' created.", vbInformation
    Application.ScreenUpdating = True
    Exit Sub
    
ErrH:
    Application.ScreenUpdating = True
    MsgBox "RunReport error: " & Err.Description, vbCritical
End Sub
