Attribute VB_Name = "Program"
Option Explicit

Sub Main()

End Sub

Sub Clean()
    Dim PrevActiveBook As Workbook
    Dim PrevDispAlerts As Boolean
    Dim s As Worksheet

    Set PrevActiveBook = ActiveWorkbook
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
    
    PrevActiveBook.Activate
    Application.Dialogs = PrevDispAlerts
    
    Sheets("Macro").Select
    Range("C7").Select
End Sub
