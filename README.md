Option Explicit

' ====== CONFIG ======
Private Const PL_YEARLY As Double = 18
Private Const SL_YEARLY As Double = 7
Private Const CL_YEARLY As Double = 7
Private Const PL_CAP    As Double = 30

' Sheet + Table names
Private Const EMP_SHEET As String = "Employee Data"
Private Const LV_SHEET  As String = "Leave Status"
Private Const EMP_TABLE As String = "tblEmp"
Private Const LV_TABLE  As String = "tblLeave"

' Column header names (must match your tables)
Private Const EMP_ID       As String = "Employee ID"
Private Const EMP_NAME     As String = "Name"
Private Const EMP_TEAM     As String = "Team"
Private Const EMP_DESIG    As String = "Designation"
Private Const EMP_DOJ      As String = "Join Date"          ' use your exact header
Private Const EMP_JOINYEAR As String = "JoinYear"           ' optional; derived if missing

' Leave columns in LV table (balances)
Private Const LV_PL As String = "PL"
Private Const LV_SL As String = "SL"
Private Const LV_CL As String = "CL"

' Possible exit date header names (first match wins)
Private ExitHeaderCandidates As Variant

' ====== DASHBOARD AS-OF ======
Private Function AsOfDateVal() As Date
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If IsDate(ws.Range("B2").Value) Then
        AsOfDateVal = CDate(ws.Range("B2").Value)
    Else
Fallback:
        AsOfDateVal = Date
    End If
End Function

' ====== MAIN ======
Public Sub RunReport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False

    ExitHeaderCandidates = Array("Exit Date", "Date Left", "Date of Leaving", "Relieving Date", "Last Working Day")

    Dim loEmp As ListObject, loLv As ListObject
    Set loEmp = GetListObject(EMP_SHEET, EMP_TABLE)
    Set loLv = GetListObject(LV_SHEET, LV_TABLE)

    ' Build balances dictionary from Leave Status (assumed balances by EmpID)
    Dim balDict As Object: Set balDict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, empKey As String
    Dim idColLv&, plCol&, slCol&, clCol&
    idColLv = SafeColIndex(loLv, EMP_ID, True)
    plCol = SafeColIndex(loLv, LV_PL, True)
    slCol = SafeColIndex(loLv, LV_SL, True)
    clCol = SafeColIndex(loLv, LV_CL, True)

    For Each lr In loLv.ListRows
        empKey = CStr(lr.Range.Columns(idColLv).Value)
        If Len(empKey) > 0 Then
            balDict(empKey & "|PL") = NzD(lr.Range.Columns(plCol).Value)
            balDict(empKey & "|SL") = NzD(lr.Range.Columns(slCol).Value)
            balDict(empKey & "|CL") = NzD(lr.Range.Columns(clCol).Value)
        End If
    Next lr

    ' Recreate output sheet
    Dim wsOut As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("FilteredData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOut = ThisWorkbook.Worksheets.Add
    wsOut.Name = "FilteredData"

    ' Headers
    Dim headers As Variant
    headers = Array( _
        "EmpID","Name","Designation","Team","JoinYear","ExitDate", _
        "PL_Accrued","SL_Accrued","CL_Accrued", _
        "PL_Balance","SL_Balance","CL_Balance", _
        "PL_Taken","SL_Taken","CL_Taken", _
        "Net_Leave_Balance" _
    )
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, UBound(headers) + 1)).Value = headers

    Dim asOf As Date: asOf = AsOfDateVal()

    ' Resolve key column indexes in EMP table (fail if critical missing)
    Dim idColEmp&, nameCol&, teamCol&, desigCol&, dojCol&, jyCol&, exitCol&
    idColEmp = SafeColIndex(loEmp, EMP_ID, True)
    nameCol = SafeColIndex(loEmp, EMP_NAME, True)
    teamCol = SafeColIndex(loEmp, EMP_TEAM, False)
    desigCol = SafeColIndex(loEmp, EMP_DESIG, False)
    dojCol = SafeColIndex(loEmp, EMP_DOJ, True)
    jyCol = SafeColIndex(loEmp, EMP_JOINYEAR, False)
    exitCol = FindFirstExistingCol(loEmp, ExitHeaderCandidates) ' 0 if none

    ' Process employees
    Dim outRow&: outRow = 2
    Dim r As ListRow

    For Each r In loEmp.ListRows
        Dim empId As String, nm As String, tm As Variant, desig As Variant
        Dim dojV As Variant, exitV As Variant, joinYr As Variant

        empId = CStr(r.Range.Columns(idColEmp).Value)
        If Len(empId) = 0 Then GoTo NextR

        nm = CStr(r.Range.Columns(nameCol).Value)
        tm = GetCellIfExists(r, teamCol)
        desig = GetCellIfExists(r, desigCol)

        dojV = r.Range.Columns(dojCol).Value
        If Not IsDate(dojV) Then GoTo NextR ' skip invalid DOJ

        If jyCol > 0 Then
            joinYr = r.Range.Columns(jyCol).Value
        Else
            joinYr = Year(CDate(dojV))
        End If

        If exitCol > 0 Then
            exitV = r.Range.Columns(exitCol).Value
        Else
            exitV = Empty
        End If

        ' ---- Accruals (robust; monthly; stop at exit/asOf) ----
        Dim plAcc As Double, slAcc As Double, clAcc As Double
        ComputeMonthlyAccruals dojV, exitV, asOf, plAcc, slAcc, clAcc

        ' Balances from LV sheet (if absent -> 0)
        Dim plBal As Double, slBal As Double, clBal As Double
        plBal = DictGetD(balDict, empId & "|PL")
        slBal = DictGetD(balDict, empId & "|SL")
        clBal = DictGetD(balDict, empId & "|CL")

        ' Taken = Accrued - Balance (if your LV holds balances; if it holds taken, flip it)
        Dim plTaken As Double, slTaken As Double, clTaken As Double
        plTaken = Round(plAcc - plBal, 2)
        slTaken = Round(slAcc - slBal, 2)
        clTaken = Round(clAcc - clBal, 2)

        Dim netBal As Double
        netBal = Round(plBal + slBal + clBal, 2)

        ' Output row
        With wsOut
            .Cells(outRow, 1).Value = empId
            .Cells(outRow, 2).Value = nm
            .Cells(outRow, 3).Value = desig
            .Cells(outRow, 4).Value = tm
            .Cells(outRow, 5).Value = joinYr
            .Cells(outRow, 6).Value = SafeDate(exitV)

            .Cells(outRow, 7).Value = Round(plAcc, 2)
            .Cells(outRow, 8).Value = Round(slAcc, 2)
            .Cells(outRow, 9).Value  = Round(clAcc, 2)

            .Cells(outRow,10).Value = Round(plBal, 2)
            .Cells(outRow,11).Value = Round(slBal, 2)
            .Cells(outRow,12).Value = Round(clBal, 2)

            .Cells(outRow,13).Value = plTaken
            .Cells(outRow,14).Value = slTaken
            .Cells(outRow,15).Value = clTaken

            .Cells(outRow,16).Value = netBal
        End With

        outRow = outRow + 1
NextR:
    Next r

    ' Sort by Net_Leave_Balance (col 16) desc
    Dim lastRow&: lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        wsOut.Sort.SortFields.Clear
        wsOut.Sort.SortFields.Add Key:=wsOut.Range(wsOut.Cells(2, 16), wsOut.Cells(lastRow, 16)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With wsOut.Sort
            .SetRange wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastRow, 16))
            .Header = xlYes
            .Apply
        End With
    End If

    ' Formatting: max row green, min row red (by total balance col 16)
    ApplyRowColorExtremes wsOut, 16

    wsOut.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "FilteredData ready (accrued, balance, taken, net; exit-aware).", vbInformation
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "RunReport error: " & Err.Description, vbCritical
End Sub

' ====== ACCRUALS (MONTHLY, EXIT-AWARE) ======
Private Sub ComputeMonthlyAccruals( _
    ByVal dojV As Variant, ByVal exitV As Variant, ByVal asOf As Date, _
    ByRef plAcc As Double, ByRef slAcc As Double, ByRef clAcc As Double)

    plAcc = 0: slAcc = 0: clAcc = 0
    If Not IsDate(dojV) Then Exit Sub

    Dim doj As Date: doj = CDate(dojV)
    Dim capStart As Date: capStart = DateSerial(2010, 1, 1)    ' business rule: start no earlier than 1-Jan-2010
    Dim startDate As Date: startDate = IIf(doj > capStart, doj, capStart)

    Dim endDate As Date
    If IsDate(exitV) Then
        endDate = CDate(exitV)
        If endDate > asOf Then endDate = asOf
    Else
        endDate = asOf
    End If
    If endDate < startDate Then Exit Sub

    ' iterate first of each month
    Dim cur As Date
    cur = DateSerial(Year(startDate), Month(startDate), 1)

    Do While cur <= endDate
        ' reset SL/CL on Jan (new year)
        If Month(cur) = 1 Then
            slAcc = 0
            clAcc = 0
        End If

        ' monthly accruals
        plAcc = WorksheetFunction.Min(PL_CAP, plAcc + (PL_YEARLY / 12#))
        slAcc = slAcc + (SL_YEARLY / 12#)
        clAcc = clAcc + (CL_YEARLY / 12#)

        cur = DateAdd("m", 1, cur)
    Loop
End Sub

' ====== HELPERS ======
Private Function GetListObject(sheetName As String, tableName As String) As ListObject
    On Error GoTo Fail
    Set GetListObject = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    Exit Function
Fail:
    Err.Raise 5, , "Table '" & tableName & "' not found on sheet '" & sheetName & "'."
End Function

Private Function SafeColIndex(lo As ListObject, colName As String, required As Boolean) As Long
    On Error Resume Next
    SafeColIndex = lo.ListColumns(colName).Index
    On Error GoTo 0
    If SafeColIndex = 0 And required Then
        Err.Raise 5, , "Column '" & colName & "' not found in table '" & lo.Name & "'."
    End If
End Function

Private Function FindFirstExistingCol(lo As ListObject, names As Variant) As Long
    Dim i As Long, idx As Long
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        idx = lo.ListColumns(CStr(names(i))).Index
        On Error GoTo 0
        If idx > 0 Then
            FindFirstExistingCol = idx
            Exit Function
        End If
    Next i
    FindFirstExistingCol = 0
End Function

Private Function GetCellIfExists(r As ListRow, colIdx As Long) As Variant
    If colIdx > 0 Then
        GetCellIfExists = r.Range.Columns(colIdx).Value
    Else
        GetCellIfExists = Empty
    End If
End Function

Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Or Not IsNumeric(v) Then
        NzD = 0#
    Else
        NzD = CDbl(v)
    End If
End Function

Private Function DictGetD(d As Object, k As String) As Double
    If d.Exists(k) Then
        DictGetD = NzD(d(k))
    Else
        DictGetD = 0#
    End If
End Function

Private Function SafeDate(v As Variant) As Variant
    If IsDate(v) Then SafeDate = CDate(v) Else SafeDate = ""
End Function

Private Sub ApplyRowColorExtremes(ws As Worksheet, totalCol As Long)
    Dim lastRow&: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim rng As Range: Set rng = ws.Range(ws.Cells(2, totalCol), ws.Cells(lastRow, totalCol))
    Dim mx As Double, mn As Double
    mx = Application.Max(rng)
    mn = Application.Min(rng)

    Dim r As Long
    For r = 2 To lastRow
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, totalCol))
            .Interior.Pattern = xlSolid
            .Interior.ColorIndex = xlNone
        End With

        If ws.Cells(r, totalCol).Value = mx Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, totalCol)).Interior.Color = vbGreen
        ElseIf ws.Cells(r, totalCol).Value = mn Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, totalCol)).Interior.Color = vbRed
            ws.Range(ws.Cells(r, 1), ws.Cells(r, totalCol)).Font.Color = vbWhite
        End If
    Next r
End Sub
