Option Explicit

' ====== CONFIG ======
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
Private Const EMP_DESIG As String = "Designation"
Private Const EMP_TEAM As String = "Team"
Private Const EMP_DOJ As String = "Date of joining"
Private Const EMP_JOINYEAR As String = "JoinYear"

Private Const LV_PL As String = "PL"
Private Const LV_SL As String = "SL"
Private Const LV_CL As String = "CL"

' Dashboard
Private Function Dsh() As Worksheet: Set Dsh = ThisWorkbook.Sheets("Dashboard"): End Function
Private Function AsOfDateVal() As Date: AsOfDateVal = Dsh.Range("B2").Value: End Function

' ====== UTILITY ======
Private Function NzD(v) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Or v = "" Then
        NzD = 0
    Else
        NzD = CDbl(v)
    End If
End Function

' ====== ACCRUAL CALC ======
Private Sub ComputeAccruals(ByVal doj As Date, ByVal asOf As Date, _
                            ByRef plAcc As Double, ByRef slAcc As Double, ByRef clAcc As Double)
    Dim startDate As Date
    startDate = Application.Max(doj, DateSerial(2010, 1, 1)) ' All start from 1-Jan-2010
    
    plAcc = 0: slAcc = 0: clAcc = 0
    Dim cur As Date
    cur = DateSerial(Year(startDate), Month(startDate), 1)
    
    Do While cur <= asOf
        ' add monthly accrual
        plAcc = Application.Min(PL_CAP, plAcc + (PL_YEARLY / 12#))
        slAcc = slAcc + (SL_YEARLY / 12#)
        clAcc = clAcc + (CL_YEARLY / 12#)
        
        ' flush SL and CL at year end (works on first day of December iteration)
        If Month(cur) = 12 And Day(cur) = 1 Then
            slAcc = 0
            clAcc = 0
        End If
        
        cur = DateAdd("m", 1, cur)
    Loop
End Sub

' ====== RUN REPORT ======
Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wsEmp As ListObject, wsLv As ListObject
    On Error Resume Next
    Set wsEmp = ThisWorkbook.Sheets("Employee Data").ListObjects(EMP_TABLE)
    Set wsLv = ThisWorkbook.Sheets("Leave Status").ListObjects(LV_TABLE)
    On Error GoTo ErrH
    
    If wsEmp Is Nothing Or wsLv Is Nothing Then
        MsgBox "Can't find required tables. Ensure ""Employee Data"" sheet has table '" & EMP_TABLE & "' and ""Leave Status"" has table '" & LV_TABLE & "'.", vbExclamation
        Exit Sub
    End If
    
    ' delete old sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrH
    
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "FilteredData"
    
    ' Headers A:Q (17 columns)
    wsOut.Range("A1:Q1").Value = Array("EmpID","Name","Designation","Team","JoinYear", _
        "PL_Balance","SL_Balance","CL_Balance", _
        "PL_Accrued","SL_Accrued","CL_Accrued", _
        "PL_Taken","SL_Taken","CL_Taken", _
        "Net_Leave_Taken","Net_Leave_Balance")
    
    ' load balances into dictionary (safe)
    Dim lvDict As Object: Set lvDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, key As String
    For Each lr In wsLv.ListRows
        key = CStr(lr.Range.Columns(wsLv.ListColumns(EMP_ID).Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_PL).Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_SL).Index).Value)
        lvDict(key & "|CL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_CL).Index).Value)
    Next lr
    
    Dim r As ListRow
    Dim empId, nm, des, tm, jy, doj
    Dim plBal As Double, slBal As Double, clBal As Double
    Dim plAcc As Double, slAcc As Double, clAcc As Double
    Dim plTaken As Double, slTaken As Double, clTaken As Double
    Dim netTaken As Double, netBalance As Double
    Dim rowOut As Long: rowOut = 2
    Dim asOf As Date: asOf = AsOfDateVal()
    
    For Each r In wsEmp.ListRows
        empId = r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value
        nm = r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value
        des = r.Range.Columns(wsEmp.ListColumns(EMP_DESIG).Index).Value
        tm = r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value
        jy = r.Range.Columns(wsEmp.ListColumns(EMP_JOINYEAR).Index).Value
        doj = r.Range.Columns(wsEmp.ListColumns(EMP_DOJ).Index).Value
        
        ' safe retrieval from dictionary (avoid missing-key error)
        If lvDict.Exists(CStr(empId) & "|PL") Then plBal = NzD(lvDict(CStr(empId) & "|PL")) Else plBal = 0
        If lvDict.Exists(CStr(empId) & "|SL") Then slBal = NzD(lvDict(CStr(empId) & "|SL")) Else slBal = 0
        If lvDict.Exists(CStr(empId) & "|CL") Then clBal = NzD(lvDict(CStr(empId) & "|CL")) Else clBal = 0
        
        Call ComputeAccruals(doj, asOf, plAcc, slAcc, clAcc)
        
        plTaken = plAcc - plBal
        slTaken = slAcc - slBal
        clTaken = clAcc - clBal
        
        netTaken = plTaken + slTaken + clTaken
        netBalance = plBal + slBal + clBal
        
        wsOut.Cells(rowOut, 1).Resize(1, 17).Value = Array( _
            empId, nm, des, tm, jy, _
            Round(plBal, 2), Round(slBal, 2), Round(clBal, 2), _
            Round(plAcc, 2), Round(slAcc, 2), Round(clAcc, 2), _
            Round(plTaken, 2), Round(slTaken, 2), Round(clTaken, 2), _
            Round(netTaken, 2), Round(netBalance, 2))
        rowOut = rowOut + 1
    Next r
    
    ' === Sort by Net_Leave_Balance (col Q) descending - only if there's at least 1 data row ===
    If rowOut > 2 Then
        With wsOut.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOut.Range("Q2:Q" & rowOut - 1), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange wsOut.Range("A1:Q" & rowOut - 1)
            .Header = xlYes
            .Apply
        End With
    End If
    
    ApplyFormatting wsOut
    wsOut.Columns.AutoFit
    MsgBox "Report ready (sorted by Net_Leave_Balance).", vbInformation
    Application.ScreenUpdating = True
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "RunReport error: " & Err.Description, vbCritical
End Sub

' ====== FORMATTING ======
Private Sub ApplyFormatting(ws As Worksheet)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim r As Long
    Dim val As Double
    Dim maxBal As Double, minBal As Double
    Dim foundAny As Boolean: foundAny = False
    
    ' Compute max/min safely by iterating numeric values only
    For r = 2 To lastRow
        If Not IsError(ws.Cells(r, 17).Value) Then
            If Trim(CStr(ws.Cells(r, 17).Value)) <> "" And IsNumeric(ws.Cells(r, 17).Value) Then
                val = CDbl(ws.Cells(r, 17).Value)
                If Not foundAny Then
                    maxBal = val
                    minBal = val
                    foundAny = True
                Else
                    If val > maxBal Then maxBal = val
                    If val < minBal Then minBal = val
                End If
            End If
        End If
    Next r
    
    If Not foundAny Then Exit Sub ' nothing numeric to compare
    
    ' clear previous formatting for data rows (optional)
    On Error Resume Next
    ws.Range("A2:Q" & lastRow).Interior.Pattern = xlNone
    ws.Range("A2:Q" & lastRow).Font.Color = vbBlack
    On Error GoTo 0
    
    ' Highlight rows: max -> green, min -> red (if equal, green takes precedence)
    For r = 2 To lastRow
        If Not IsError(ws.Cells(r, 17).Value) Then
            If Trim(CStr(ws.Cells(r, 17).Value)) <> "" And IsNumeric(ws.Cells(r, 17).Value) Then
                val = CDbl(ws.Cells(r, 17).Value)
                If val = maxBal Then
                    ws.Rows(r).Interior.Color = vbGreen
                    ws.Rows(r).Font.Color = vbBlack
                ElseIf val = minBal Then
                    ws.Rows(r).Interior.Color = vbRed
                    ws.Rows(r).Font.Color = vbWhite
                End If
            End If
        End If
    Next r
End Sub
