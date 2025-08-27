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

' Dashboard helpers
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
    startDate = Application.Max(doj, DateSerial(2010, 1, 1)) ' accruals start from 1-Jan-2010
    plAcc = 0: slAcc = 0: clAcc = 0

    Dim cur As Date
    cur = DateSerial(Year(startDate), Month(startDate), 1)

    Do While cur <= asOf
        ' monthly accrual
        plAcc = Application.Min(PL_CAP, plAcc + (PL_YEARLY / 12#))
        slAcc = slAcc + (SL_YEARLY / 12#)
        clAcc = clAcc + (CL_YEARLY / 12#)

        ' flush SL and CL at year end
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
    Set wsEmp = ThisWorkbook.Sheets("Employee Data").ListObjects(EMP_TABLE)
    Set wsLv = ThisWorkbook.Sheets("Leave Status").ListObjects(LV_TABLE)

    ' delete old report
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "FilteredData"

    ' header row (15 columns including Net_Leave)
    wsOut.Range("A1:O1").Value = Array("EmpID", "Name", "Designation", "Team", "JoinYear", _
        "PL_Balance", "SL_Balance", "CL_Balance", _
        "PL_Accrued", "SL_Accrued", "CL_Accrued", _
        "PL_Taken", "SL_Taken", "CL_Taken", "Net_Leave")

    ' load leave balances into dictionary
    Dim lvDict As Object: Set lvDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, key As String
    For Each lr In wsLv.ListRows
        key = CStr(lr.Range.Columns(wsLv.ListColumns(EMP_ID).Index).Value)
        lvDict(key & "|PL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_PL).Index).Value)
        lvDict(key & "|SL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_SL).Index).Value)
        lvDict(key & "|CL") = NzD(lr.Range.Columns(wsLv.ListColumns(LV_CL).Index).Value)
    Next lr

    ' build output rows
    Dim r As ListRow, empId, nm, des, tm, jy, doj
    Dim plBal#, slBal#, clBal#, plAcc#, slAcc#, clAcc#
    Dim rowOut&: rowOut = 2
    Dim asOf As Date: asOf = AsOfDateVal()

    For Each r In wsEmp.ListRows
        empId = r.Range.Columns(wsEmp.ListColumns(EMP_ID).Index).Value
        nm = r.Range.Columns(wsEmp.ListColumns(EMP_NAME).Index).Value
        des = r.Range.Columns(wsEmp.ListColumns(EMP_DESIG).Index).Value
        tm = r.Range.Columns(wsEmp.ListColumns(EMP_TEAM).Index).Value
        jy = r.Range.Columns(wsEmp.ListColumns(EMP_JOINYEAR).Index).Value
        doj = r.Range.Columns(wsEmp.ListColumns(EMP_DOJ).Index).Value

        plBal = NzD(lvDict(CStr(empId) & "|PL"))
        slBal = NzD(lvDict(CStr(empId) & "|SL"))
        clBal = NzD(lvDict(CStr(empId) & "|CL"))

        Call ComputeAccruals(doj, asOf, plAcc, slAcc, clAcc)

        Dim plTaken#, slTaken#, clTaken#
        plTaken = plAcc - plBal
        slTaken = slAcc - slBal
        clTaken = clAcc - clBal

        ' Net_Leave = PL_Balance + SL_Balance + CL_Balance
        Dim netLeave#
        netLeave = plBal + slBal + clBal

        wsOut.Cells(rowOut, 1).Resize(1, 15).Value = Array(empId, nm, des, tm, jy, _
            plBal, slBal, clBal, _
            Round(plAcc, 2), Round(slAcc, 2), Round(clAcc, 2), _
            Round(plTaken, 2), Round(slTaken, 2), Round(clTaken, 2), netLeave)

        rowOut = rowOut + 1
    Next r

    ' ===== SORT BY Net_Leave (Column O) DESC =====
    Dim lastRow As Long
    lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    wsOut.Sort.SortFields.Clear
    wsOut.Sort.SortFields.Add Key:=wsOut.Range("O2:O" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With wsOut.Sort
        .SetRange wsOut.Range("A1:O" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' ===== HIGHLIGHTING =====
    Dim maxLeave As Double, i As Long
    maxLeave = Application.WorksheetFunction.Max(wsOut.Range("O2:O" & lastRow))

    For i = 2 To lastRow
        ' Highlight only the Net_Leave cell if it's max
        If wsOut.Cells(i, 15).Value = maxLeave Then
            wsOut.Cells(i, 15).Interior.Color = vbGreen
        End If

        ' Highlight PL_Taken + SL_Taken + CL_Taken if negative
        If wsOut.Cells(i, 12).Value + wsOut.Cells(i, 13).Value + wsOut.Cells(i, 14).Value < 0 Then
            wsOut.Range(wsOut.Cells(i, 12), wsOut.Cells(i, 14)).Interior.Color = vbRed
        End If
    Next i

    wsOut.Columns.AutoFit
    MsgBox "Report ready with accruals, taken leaves, Net_Leave (sorted + highlighted).", vbInformation

    Application.ScreenUpdating = True
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "RunReport error: " & Err.Description, vbCritical
End Sub
