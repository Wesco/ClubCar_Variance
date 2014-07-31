Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Import_Err
    MsgBox "Select Forecast P (1)"
    UserImportFile DestRange:=Sheets("Report1").Range("A1"), _
                   ShowAllData:=True, _
                   SourceSheet:="PivotTableP", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"

    MsgBox "Select Forecast A (1)"
    UserImportFile DestRange:=Sheets("Report1").Range("A" & Sheets("Report1").UsedRange.Rows.Count + 1), _
                   ShowAllData:=True, _
                   SourceSheet:="PivotTableA", _
                   Title:="Open Forecast A", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"

    MsgBox "Select Forecast P (2)"
    UserImportFile DestRange:=Sheets("Report2").Range("A1"), _
                   ShowAllData:=True, _
                   SourceSheet:="PivotTableP", _
                   Title:="Open Forecast P", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"

    MsgBox "Select Forecast A (2)"
    UserImportFile DestRange:=Sheets("Report2").Range("A" & Sheets("Report2").UsedRange.Rows.Count + 1), _
                   ShowAllData:=True, _
                   SourceSheet:="PivotTableA", _
                   Title:="Open Forecast A", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"
    On Error GoTo 0

    Exit Sub

Import_Err:
    MsgBox Prompt:="Error " & Err.Number & " (" & Err.Description & ") occurred in " & Err.Source & ".", _
           Title:="Oops!"
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
    Application.DisplayAlerts = PrevDispAlerts

    Sheets("Macro").Select
    Range("C7").Select
End Sub
