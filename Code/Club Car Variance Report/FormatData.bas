Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatReport(Sheet As Variant)
    Dim ColHeaders As Variant
    Dim TotalCols As Integer

    'Select sheet
    If TypeName(Sheet) = "Integer" Or TypeName(Sheet) = "String" Then
        Sheets(Sheet).Select
    ElseIf TypeName(Sheet) = "Worksheet" Then
        Sheet.Select
    Else
        Err.Raise 50002, "FormatReport", "Invalid type"
    End If
    
    'Store the column headerse
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value
    
    'Filter for duplicate column headers
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="=Item Number"
    
    'Remove duplicate headers
    Cells.Delete
    
    'Reinsert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub
