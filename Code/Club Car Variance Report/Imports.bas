Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False, Optional SourceSheet As String = "", Optional FileFilter = "")
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    File = Application.GetOpenFilename(FileFilter)

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File
        If SourceSheet = "" Then SourceSheet = ActiveSheet.Name
        If ShowAllData = True Then
            On Error Resume Next
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.UsedRange.Columns.Hidden = False
            ActiveSheet.UsedRange.Rows.Hidden = False
            On Error GoTo 0
        End If
        Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        If DelFile = True Then
            DeleteFile File
        End If
    Else
        Err.Raise 18
    End If
    Application.DisplayAlerts = OldDispAlert
End Sub
