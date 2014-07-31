Attribute VB_Name = "Program"
Option Explicit

Sub Main()

End Sub

Sub Clean()
    Dim s As Worksheet
    Dim PrevDispAlerts As Boolean

    Application.DisplayAlerts = False
    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Rows.Hidden = False
            s.Columns.Hidden = False
            s.Cells.Delete
        End If
    Next
End Sub
