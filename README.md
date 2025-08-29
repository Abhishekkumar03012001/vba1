Public Sub UpdateLeaveFromUI()
    On Error GoTo ErrH
    
    ' Ask for Employee details
    Dim empId As Variant: empId = InputBox("Enter Employee ID to update:", "Update Leave")
    If empId = "" Then Exit Sub
    
    Dim empName As Variant: empName = InputBox("Enter Employee Name:", "Update Leave")
    If empName = "" Then Exit Sub
    
    Dim pl As Variant, sl As Variant, cl As Variant
    pl = InputBox("New PL balance (number):", "Update Leave")
    sl = InputBox("New SL balance (number):", "Update Leave")
    cl = InputBox("New CL balance (number):", "Update Leave")
    If pl = "" Or sl = "" Or cl = "" Then
        MsgBox "Update cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Load Leave Status table
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Leave Status").ListObjects(LV_TABLE)
    
    Dim r As ListRow, found As Boolean
    For Each r In lo.ListRows
        If CStr(r.Range.Columns(lo.ListColumns(EMP_ID).Index).Value) = CStr(empId) Then
            ' Update balances + name if row already exists
            r.Range.Columns(lo.ListColumns(EMP_NAME).Index).Value = empName
            r.Range.Columns(lo.ListColumns(LV_PL).Index).Value = CDbl(pl)
            r.Range.Columns(lo.ListColumns(LV_SL).Index).Value = CDbl(sl)
            r.Range.Columns(lo.ListColumns(LV_CL).Index).Value = CDbl(cl)
            found = True
            Exit For
        End If
    Next r
    
    ' Add new row if not found
    If Not found Then
        Set r = lo.ListRows.Add
        r.Range.Columns(lo.ListColumns(EMP_ID).Index).Value = empId
        r.Range.Columns(lo.ListColumns(EMP_NAME).Index).Value = empName
        r.Range.Columns(lo.ListColumns(LV_PL).Index).Value = CDbl(pl)
        r.Range.Columns(lo.ListColumns(LV_SL).Index).Value = CDbl(sl)
        r.Range.Columns(lo.ListColumns(LV_CL).Index).Value = CDbl(cl)
    End If
    
    MsgBox "Leave status updated.", vbInformation
    Exit Sub
    
ErrH:
    MsgBox "UpdateLeave error: " & Err.Description, vbCritical
End Sub
