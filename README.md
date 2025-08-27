Option Explicit

' ====== CONFIG ======
Private Const PL_YEARLY As Double = 18
Private Const SL_YEARLY As Double = 7
Private Const CL_YEARLY As Double = 7
Private Const PL_CAP As Double = 30

' Table names - make sure these match your table names
Private Const EMP_TABLE As String = "tblEmp"
Private Const LV_TABLE As String = "tblLeave"

' Column header names used in those tables (must match your headers)
Private Const EMP_ID As String = "Employee ID"
Private Const EMP_NAME As String = "Name"
Private Const EMP_DESIG As String = "Designation"
Private Const EMP_TEAM As String = "Team"
Private Const EMP_DOJ As String = "Date of joining"
Private Const EMP_JOINYEAR As String = "JoinYear"

Private Const LV_PL As String = "PL"
Private Const LV_SL As String = "SL"
Private Const LV_CL As String = "CL"

' ====== DASHBOARD helpers ======
Private Function Dsh() As Worksheet: Set Dsh = ThisWorkbook.Sheets("Dashboard"): End Function
Private Function AsOfDateVal() As Date: AsOfDateVal = Dsh.Range("B2").Value: End Function
Private Function JoinYearSel() As String: JoinYearSel = CStr(Dsh.Range("B4").Value): End Function
Private Function DesigSel() As String: DesigSel = CStr(Dsh.Range("B5").Value): End Function
Private Function TeamSel() As String: TeamSel = CStr(Dsh.Range("B6").Value): End Function
Private Function LeaveTypeSel() As String: LeaveTypeSel = CStr(Dsh.Range("B8").Value): End Function
Private Function MetricSel() As String: MetricSel = CStr(Dsh.Range("B9").Value): End Function
Private Function TopNSel() As Long
    Dim v: v = Dsh.Range("B10").Value
    If v = "" Or Not IsNumeric(v) Then TopNSel = 10 Else TopNSel = CLng(v)
End Function

' ====== utility ======
Private Function LastDayOfMonth(ByVal d As Date) As Date
    LastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Private Function MonthsCreditedYTD(ByVal asOfDate As Date, ByVal doj As Date) As Long
    Dim y As Long: y = Year(asOfDate)
    If doj > asOfDate Then MonthsCreditedYTD = 0: Exit Function
    Dim startMonth As Long
    If Year(doj) < y Then
        startMonth = 1
    ElseIf Year(doj) = y Then
        startMonth = Month(doj)
    Else
        MonthsCreditedYTD = 0: Exit Function
    End If
    Dim m As Long, cnt As Long
    For m = startMonth To 12
        If LastDayOfMonth(DateSerial(y, m, 1)) <= asOfDate Then cnt = cnt + 1 Else Exit For
    Next m
    MonthsCreditedYTD = cnt
End Function

Private Function AccruedPL(ByVal asOfDate As Date, ByVal doj As Date) As Double
    Dim months As Long: months = MonthsCreditedYTD(asOfDate, doj)
    AccruedPL = Application.Min(PL_CAP, months * (PL_YEARLY / 12#))
End Function
Private Function AccruedSL(ByVal asOfDate As Date, ByVal doj As Date) As Double
    AccruedSL = MonthsCreditedYTD(asOfDate, doj) * (SL_YEARLY / 12#)
End Function
Private Function AccruedCL(ByVal asOfDate As Date, ByVal doj As Date) As Double
    AccruedCL = MonthsCreditedYTD(asOfDate, doj) * (CL_YEARLY / 12#)
End Function

Private Function NzD(v) As Double
    If IsError(v) Then NzD = 0: Exit Function
    If IsEmpty(v) Or IsNull(v) Or v = "" Then NzD = 0 Else NzD = CDbl(v)
End Function

Private Function PickMetric(balance As Double, accrued As Double) As Double
    Select Case UCase(MetricSel())
        Case "BALANCE": PickMetric = balance
        Case "ACCUMULATED": PickMetric = accrued
        Case "TAKEN": PickMetric = accrued - balance
        Case Else: PickMetric = balance
    End Select
End Function

' ====== INIT LISTS (fills Lists sheet and sets Dashboard validations) ======
Public Sub InitLists()
    On Error GoTo ErrH
    Dim wsE As ListObject, wsL As Worksheet
    Set wsE = ThisWorkbook.Sheets("Employee Data").ListObjects(EMP_TABLE)
    Set wsL = ThisWorkbook.Sheets("Lists")
    
    Dim d1 As Object, d2 As Object, d3 As Object
    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = CreateObject("Scripting.Dictionary")
    Set d3 = CreateObject("Scripting.Dictionary")
    
    Dim r As ListRow, y, des, t
    For Each r In wsE.ListRows
        y = r.Range.Columns(wsE.ListColumns(EMP_JOINYEAR).Index).Value
        des = r.Range.Columns(wsE.ListColumns(EMP_DESIG).Index).Value
        t = r.Range.Columns(wsE.ListColumns(EMP_TEAM).Index).Value
        If Not d1.exists(y) And Not IsEmpty(y) Then d1.Add y, y
        If Not d2.exists(des) And Not IsEmpty(des) Then d2.Add des, des
        If Not d3.exists(t) And Not IsEmpty(t) Then d3.Add t, t
    Next r
    
    wsL.Cells.Clear
    wsL.Range("A1").Value = "JoinYear": wsL.Range("A2").Value = "All"
    wsL.Range("B1").Value = "Designation": wsL.Range("B2").Value = "All"
    wsL.Range("C1").Value = "Team": wsL.Range("C2").Value = "All"
    
    If d1.Count > 0 Then wsL.Range("A3").Resize(d1.Count, 1).Value = Application.Transpose(d1.Items)
    If d2.Count > 0 Then wsL.Range("B3").Resize(d2.Count, 1).Value = Application.Transpose(d2.Items)
    If d3.Count > 0 Then wsL.Range("C3").Resize(d3.Count, 1).Value = Application.Transpose(d3.Items)
    
    ' --- Safe range definitions ---
    Dim lastRowA As Long, lastRowB As Long, lastRowC As Long
    lastRowA = Application.Max(2, wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row)
    lastRowB = Application.Max(2, wsL.Cells(wsL.Rows.Count, "B").End(xlUp).Row)
    lastRowC = Application.Max(2, wsL.Cells(wsL.Rows.Count, "C").End(xlUp).Row)
    
    ' Data validation on Dashboard cells (never fails, even if only "All")
    With Dsh.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=Lists!A2:A" & lastRowA
    End With
    With Dsh.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=Lists!B2:B" & lastRowB
    End With
    With Dsh.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=Lists!C2:C" & lastRowC
    End With
    
    MsgBox "Lists initialized and Dashboard dropdowns created.", vbInformation
    Exit Sub
ErrH:
    MsgBox "InitLists error: " & Err.Description, vbCritical
End Sub


' ====== RUN REPORT (creates FilteredData sheet) ======
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
        key = CStr(lr.Range.Columns(wsLv.ListColumns(EMP_ID).Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_PL).Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_SL).Index).Value)
        lvDict(key & "|CL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_CL).Index).Value)
    Next lr
    
    Dim r As ListRow, empId, nm, des, tm, jy, doj
    Dim plBal#, slBal#, clBal#
    Dim rowOut&: rowOut = 2
    Dim asOf As Date: asOf = AsOfDateVal()
    
    For Each r In wsEmp.ListRows
        empId = r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value
        nm = r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value
        des = r.Range.Columns(wsEmp.ListColumns(EMP_DESIG).Index).Value
        tm = r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value
        jy = r.Range.Columns(wsEmp.ListColumns(EMP_JOINYEAR).Index).Value
        doj = r.Range.Columns(wsEmp.ListColumns(EMP_DOJ).Index).Value
        
        ' filter
        If JoinYearSel <> "All" Then If CStr(jy) <> JoinYearSel Then GoTo NextEmp
        If DesigSel <> "All" Then If CStr(des) <> DesigSel Then GoTo NextEmp
        If TeamSel <> "All" Then If CStr(tm) <> TeamSel Then GoTo NextEmp
        
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
        
        wsOut.Cells(rowOut, 1).Resize(1, 10).Value = Array(empId, nm, des, tm, jy, plBal, slBal, clBal, metric, detail)
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

' ====== Formatting & highlight rules ======
Private Sub ApplyBonusFormatting(ws As Worksheet)
    On Error Resume Next
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ws.Range("K1").Value = "TotalBalance"
    ws.Range("K2:K" & lastRow).FormulaR1C1 = "=RC6+RC7+RC8"
    Dim maxVal As Double: maxVal = Application.Max(ws.Range("K2:K" & lastRow))
    
    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, 6).Value < 0 Or ws.Cells(r, 7).Value < 0 Or ws.Cells(r, 8).Value < 0 Then
            ws.Rows(r).Interior.Color = vbRed
            ws.Rows(r).Font.Color = vbWhite
        End If
        If NzD(ws.Cells(r, 11).Value) = maxVal Then
            ws.Rows(r).Interior.Color = vbGreen
            ws.Rows(r).Font.Color = vbBlack
        End If
    Next r
    ws.Columns("K").Hidden = True
End Sub

' ====== Update Leave via InputBoxes ======
Public Sub UpdateLeaveFromUI()
    On Error GoTo ErrH
    Dim empId As Variant: empId = InputBox("Enter Employee ID to update:", "Update Leave")
    If empId = "" Then Exit Sub
    
    Dim pl As Variant, sl As Variant, cl As Variant
    pl = InputBox("New PL balance (number):", "Update Leave")
    sl = InputBox("New SL balance (number):", "Update Leave")
    cl = InputBox("New CL balance (number):", "Update Leave")
    If pl = "" Or sl = "" Or cl = "" Then MsgBox "Update cancelled.", vbExclamation: Exit Sub
    
    Dim lo As ListObject: Set lo = ThisWorkbook.Sheets("Leave Status").ListObjects(LV_TABLE)
    Dim r As ListRow, found As Boolean
    For Each r In lo.ListRows
        If CStr(r.Range.Columns(lo.ListColumns(EMP_ID).Index).Value) = CStr(empId) Then
            r.Range.Columns(lo.ListColumns(LV_PL).Index).Value = CDbl(pl)
            r.Range.Columns(lo.ListColumns(LV_SL).Index).Value = CDbl(sl)
            r.Range.Columns(lo.ListColumns(LV_CL).Index).Value = CDbl(cl)
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        Set r = lo.ListRows.Add
        r.Range.Columns(lo.ListColumns(EMP_ID).Index).Value = empId
        r.Range.Columns(lo.ListColumns(EMP_NAME).Index).Value = ""
        r.Range.Columns(lo.ListColumns(LV_PL).Index).Value = CDbl(pl)
        r.Range.Columns(lo.ListColumns(LV_SL).Index).Value = CDbl(sl)
        r.Range.Columns(lo.ListColumns(LV_CL).Index).Value = CDbl(cl)
    End If
    
    MsgBox "Leave status updated.", vbInformation
    Exit Sub
ErrH:
    MsgBox "UpdateLeave error: " & Err.Description, vbCritical
End Sub

' ====== Export FilteredData to new workbook ======
Public Sub ExportReport()
    On Error GoTo ErrH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FilteredData")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Run the report first.", vbExclamation: Exit Sub
    
    ws.Copy
    Dim newWb As Workbook: Set newWb = ActiveWorkbook
    Dim path As String: path = ThisWorkbook.Path & "\Filtered_Results_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    newWb.SaveAs path
    newWb.Close False
    MsgBox "Exported to: " & path, vbInformation
    Exit Sub
ErrH:
    MsgBox "ExportReport error: " & Err.Description, vbCritical
End Sub

' ====== Quick Refresh that runs report & refreshes pivots ======
Public Sub RefreshDashboard()
    On Error Resume Next
    RunReport
    ThisWorkbook.RefreshAll
    MsgBox "Dashboard refreshed.", vbInformation
End Sub
