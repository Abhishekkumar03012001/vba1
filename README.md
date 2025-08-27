A2: As Of Date
A4: Year of Joining
A5: Designation
A6: Team
A8: Leave Type
A9: Metric
A10: Top N

Sub RefreshDashboard()
    ThisWorkbook.RefreshAll
    MsgBox "✅ Dashboard has been refreshed!", vbInformation
End Sub

Sub ClearAllFilters()
    Dim ws As Worksheet
    Dim pvt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pvt In ws.PivotTables
            pvt.ClearAllFilters
        Next pvt
    Next ws
    MsgBox "🔄 All filters cleared!"
End Sub

B4:  Year of Joining   (dropdown from Lists!A:A; include “All”)
B5:  Designation       (dropdown from Lists!B:B; include “All”)
B6:  Team              (dropdown from Lists!C:C; include “All”)
B8:  Leave Type        (dropdown: Total, PL, SL, CL)
B9:  Metric            (dropdown: Balance, Accumulated, Taken)
B10: Top N             (number, default 10)
