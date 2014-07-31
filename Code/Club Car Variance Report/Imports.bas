Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 7/31/2014
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional SourceSheet As String = "", Optional ShowAllData = False, Optional FileFilter = "", Optional InitialFileName As String = "")
    Dim File As Variant             'Full path to user selected file
    Dim PrevDispAlerts As Boolean   'Original state of Application.DisplayAlerts

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .InitialFileName = InitialFileName
        .Show

        If .SelectedItems.Count = 1 Then
            File = .SelectedItems(1)
        End If
    End With

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    If File <> "" Then
        Workbooks.Open File

        If SourceSheet = "" Then SourceSheet = ActiveSheet.Name
        Sheets(SourceSheet).Select

        If ShowAllData = True Then
            ActiveSheet.AutoFilterMode = False
            ActiveSheet.Rows.Hidden = False
            ActiveSheet.Columns.Hidden = False
        End If

        Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange

        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close
    Else
        Err.Raise 18, "UserImportFile", "User canceled import"
    End If
    Application.DisplayAlerts = PrevDispAlerts
End Sub
