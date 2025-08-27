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
' Priority:
' 1) Named range "AsOfDate"
' 2) Worksheet "Dashboard" cell B2
' 3) Worksheet "Settings" cell B2
' 4) Ask user via InputBox (default = today)
' 5) Fallback = Date (today)
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
    v = vbNull
    v = ThisWorkbook.Worksheets("Dashboard").Range("B2").Value
    If Err.Number = 0 Then
        If Not IsEmpty(v) Then AsOfDateVal = CDate(v): Exit Function
    End If
    Err.Clear

    ' 3) Settings!B2
    v = vbNull
    v = ThisWorkbook.Worksheets("Settings").Range("B2").Value
    If Err.Number = 0 Then
        If Not IsEmpty(v) Then AsOfDateVal = CDate(v): Exit Function
    End If
    Err.Clear

    ' 4) Prompt user (optional)
    Dim s As String
    s = InputBox("Enter As-Of date (e.g. 2025-08-27). Leave blank to use today:", "As-Of Date", Format(Date, "yyyy-mm-dd"))
    If Trim(s) <> "" Then
        If IsDate(s) Then
            AsOfDateVal = CDate(s)
            Exit Function
        End If
    End If

    ' 5) fallback to today
    AsOfDateVal = Date
End Function
