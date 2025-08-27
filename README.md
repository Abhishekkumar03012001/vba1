' === Safe Null/Empty to Double converter ===
Private Function NzD(val As Variant) As Double
    If IsError(val) Then
        NzD = 0
    ElseIf IsEmpty(val) Or IsNull(val) Or val = "" Then
        NzD = 0
    ElseIf IsNumeric(val) Then
        NzD = CDbl(val)
    Else
        NzD = 0
    End If
End Function
