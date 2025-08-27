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
    Dim rowOut&: rowOut = 2
    
    For Each r In wsEmp.ListRows
        empId = r.Range.Cells(1, wsEmp.ListColumns("Employee ID").Index).Value
        nm = r.Range.Cells(1, wsEmp.ListColumns("Name").Index).Value
        tm = r.Range.Cells(1, wsEmp.ListColumns("Team").Index).Value
        jy = NzD(r.Range.Cells(1, wsEmp.ListColumns("JoinYear").Index).Value)
        
        ' Apply Dashboard Filters
        If JoinYearSel <> "All" Then If CStr(jy) <> CStr(JoinYearSel) Then GoTo NextEmp
        If TeamSel <> "All" Then If CStr(tm) <> CStr(TeamSel) Then GoTo NextEmp
        If DesigSel <> "All" Then
            If r.Range.Cells(1, wsEmp.ListColumns("Designation").Index).Value <> DesigSel Then GoTo NextEmp
        End If
        
        ' Balances
        plBal = NzD(lvDict(CStr(empId) & "|PL"))
        slBal = NzD(lvDict(CStr(empId) & "|SL"))
        totalBal = plBal + slBal
        
        ' Metric selection
        Select Case UCase(LeaveTypeSel())
            Case "PL"
                metric = PickMetric(plBal, plBal)   ' use PL only
                detail = "PL-" & MetricSel()
            Case "SL"
                metric = PickMetric(slBal, slBal)   ' use SL only
                detail = "SL-" & MetricSel()
            Case Else
                metric = PickMetric(totalBal, totalBal)   ' use total
                detail = "Total-" & MetricSel()
        End Select
        
        ' Output row
        wsOut.Cells(rowOut, 1).Resize(1, 9).Value = _
            Array(empId, nm, tm, jy, plBal, slBal, metric, detail, totalBal)
        rowOut = rowOut + 1
NextEmp:
    Next r
    
    ' Sort by MetricValue and keep Top N
    If rowOut > 2 Then
        wsOut.Range("A1:I" & rowOut - 1).Sort Key1:=wsOut.Range("G2"), Order1:=xlDescending, Header:=xlYes
        Dim keepRows As Long: keepRows = Application.Min(TopNSel, rowOut - 2)
        If rowOut - 2 > keepRows Then wsOut.Rows(keepRows + 2 & ":" & rowOut - 1).Delete
    End If
    
    ' Formatting
    With wsOut.UsedRange
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    
    MsgBox "Report ready: sheet 'DATA2' created.", vbInformation
    Application.ScreenUpdating = True
    Exit Sub
    
ErrH:
    Application.ScreenUpdating = True
    MsgBox "Error in GenerateDATA2: " & Err.Description, vbCritical
End Sub

' --- Support Functions ---
Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or Trim(v & "") = "" Then
        NzD = 0
    Else
        NzD = CDbl(v)
    End If
End Function

Private Function PickMetric(balance As Double, accrued As Double) As Double
    Select Case UCase(MetricSel())
        Case "BALANCE"
            PickMetric = balance
        Case "ACCRUED"
            PickMetric = accrued
        Case "DIFF"
            PickMetric = balance - accrued
        Case Else
            PickMetric = balance
    End Select
End Function
