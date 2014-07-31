Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Import_Err
    UserImportFile DestRange:=Sheets("Report1").Range("A1"), _
                   ShowAllData:=True, _
                   SourceSheet:="Forecast", _
                   FileFilter:="Excel Files (*.xlsx; *.xls), *.xlsx; *.xls", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Order Report\"

    UserImportFile DestRange:=Sheets("Report2").Range("A1"), _
                   ShowAllData:=True, _
                   SourceSheet:="Forecast", _
                   FileFilter:="Excel Files (*.xlsx; *.xls), *.xlsx; *.xls", _
                   InitialFileName:="\\br3615gaps\gaps\Club Car\Order Report\"
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
