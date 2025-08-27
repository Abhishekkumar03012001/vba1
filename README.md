Option Explicit

' === Safe Null/Empty to Double converter ===
Private Function NzD(val As Variant) As Double
    If IsError(val) Then
        NzD = 0
    ElseIf IsEmpty(val) Or IsNull(val) Or Trim(CStr(val)) = "" Then
        NzD = 0
    ElseIf IsNumeric(val) Then
        NzD = CDbl(val)
    Else
        NzD = 0
    End If
End Function

' === Returns the As-Of date used by the report ===
Private Function AsOfDateVal() As Date
    Dim v As Variant
    On Error Resume Next

    ' 1) Named range AsOfDate
    v = vbNull
    If ThisWorkbook.Names.Count > 0 Then
        v = ThisWorkbook.Names("AsOfDate").RefersToRange.Value
        If Err.Number = 0 Then
            If Not IsEmpty(v) Then AsOfDateVal = CDate(v): Exit Function
        End If
        Err.Clear
    End If

    ' 2) Dashboard!B2
    v = ThisWorkbook.Worksheets("Dashboard").Range("B2").Value
    If Err.Number = 0 Then
        If Not IsEmpty(v) Then AsOfDateVal = CDate(v): Exit Function
    End If
    Err.Clear

    ' 3) Settings!B2
    v = ThisWorkbook.Worksheets("Settings").Range("B2").Value
    If Err.Number = 0 Then
        If Not IsEmpty(v) Then AsOfDateVal = CDate(v): Exit Function
    End If
    Err.Clear

    ' 4) Prompt user (default = today)
    Dim s As String
    s = InputBox("Enter As-Of date (yyyy-mm-dd). Leave blank for today:", _
                 "As-Of Date", Format(Date, "yyyy-mm-dd"))
    If Trim(s) <> "" Then
        If IsDate(s) Then AsOfDateVal = CDate(s): Exit Function
    End If

    ' 5) fallback to today
    AsOfDateVal = Date
End Function

' ====== MAIN REPORT MACRO ======
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
    
    wsOut.Range("A1:O1").Value = Array("Employee ID", "Name", "Designation", "Team", "JoinYear", _
                                       "PL_Balance", "SL_Balance", "CL_Balance", _
                                       "PL_Received", "SL_Received", "CL_Received", _
                                       "Total_Received", "Total_Taken", "Difference", _
                                       "MetricValue", "MetricDetail")
    
    ' Build leave lookup dictionary
    Dim lvDict As Object: Set lvDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, key As String
    For Each lr In wsLv.ListRows
        key = CStr(lr.Range.Columns(wsLv.ListColumns(EMP_ID).Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_PL).Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_SL).Index).Value)
        lvDict(key & "|CL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_CL).Index).Value)
    Next lr
    
    Dim r As ListRow, empId, nm, des, tm, jy, doj
    Dim plBal#, slBal#, clBal#
    Dim rowOut&: rowOut = 2
    Dim asOf As Date: asOf = AsOfDateVal()
    Dim baseDate As Date: baseDate = DateSerial(2010, 1, 1)
    
    Dim maxBal As Double: maxBal = -1
    
    For Each r In wsEmp.ListRows
        empId = r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value
        nm = r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value
        des = r.Range.Columns(wsEmp.ListColumns(EMP_DESIG).Index).Value
        tm = r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value
        jy = r.Range.Columns(wsEmp.ListColumns(EMP_JOINYEAR).Index).Value
        doj = r.Range.Columns(wsEmp.ListColumns(EMP_DOJ).Index).Value
        
        ' balances from Leave Status
        plBal = NzD(lvDict(CStr(empId) & "|PL"))
        slBal = NzD(lvDict(CStr(empId) & "|SL"))
        clBal = NzD(lvDict(CStr(empId) & "|CL"))
        
        ' ==== LEAVE ACCRUAL RULES ====
        Dim startDate As Date
        startDate = IIf(doj > baseDate, doj, baseDate)
        
        ' PL accrues since DOJ, capped at 30
        Dim monthsWorked As Long
        monthsWorked = DateDiff("m", startDate, asOf)
        If monthsWorked < 0 Then monthsWorked = 0
        Dim plReceived As Double
        plReceived = monthsWorked * (18 / 12)
        If plReceived > 30 Then plReceived = 30
        
        ' SL/CL accrue only in the current year, flush every Jan
        Dim yearStart As Date: yearStart = DateSerial(Year(asOf), 1, 1)
        If doj > yearStart Then yearStart = doj
        
        Dim monthsThisYear As Long
        monthsThisYear = DateDiff("m", yearStart, asOf)
        If monthsThisYear < 0 Then monthsThisYear = 0
        
        Dim slReceived As Double, clReceived As Double
        slReceived = monthsThisYear * (7 / 12)
        If slReceived > 7 Then slReceived = 7
        clReceived = monthsThisYear * (7 / 12)
        If clReceived > 7 Then clReceived = 7
        
        ' Totals
        Dim totalReceived#, totalTaken#, diff#
        totalReceived = plReceived + slReceived + clReceived
        totalTaken = totalReceived - (plBal + slBal + clBal)
        diff = (plBal + slBal + clBal) - totalTaken
        
        ' Metric (example)
        Dim metric#, detail As String
        Select Case UCase(LeaveTypeSel())
            Case "PL"
                metric = plBal
                detail = "PL-" & MetricSel()
            Case "SL"
                metric = slBal
                detail = "SL-" & MetricSel()
            Case "CL"
                metric = clBal
                detail = "CL-" & MetricSel()
            Case Else
                metric = plBal + slBal + clBal
                detail = "Total-" & MetricSel()
        End Select
        
        wsOut.Cells(rowOut, 1).Resize(1, 15).Value = _
            Array(empId, nm, des, tm, jy, _
                  plBal, slBal, clBal, _
                  plReceived, slReceived, clReceived, _
                  totalReceived, totalTaken, diff, _
                  metric, detail)
        
        If (plBal + slBal + clBal) > maxBal Then
            maxBal = (plBal + slBal + clBal)
        End If
        rowOut = rowOut + 1
    Next r
    
    ' === Highlighting ===
    Dim lastRow As Long: lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    Dim c As Range
    
    ' Negative difference = Red
    For Each c In wsOut.Range("N2:N" & lastRow)
        If c.Value < 0 Then c.Interior.Color = vbRed
    Next c
    
    ' Max balance = Green
    Dim rRow As Long
    For rRow = 2 To lastRow
        If (wsOut.Cells(rRow, 6).Value + wsOut.Cells(rRow, 7).Value + wsOut.Cells(rRow, 8).Value) = maxBal Then
            wsOut.Range("F" & rRow & ":H" & rRow).Interior.Color = vbGreen
        End If
    Next rRow
    
    ' formatting
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
