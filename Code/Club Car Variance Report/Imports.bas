Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ImportFile
' Date : 7/31/2014
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub ImportFile(DestRange As Range, Optional Title As String = "Open", Optional SourceSheet As String = "", _
                   Optional ShowAllData = False, Optional InitialFileName As String = "")
    Dim File As Variant             'Full path to user selected file
    Dim PrevDispAlerts As Boolean   'Original state of Application.DisplayAlerts
    Dim i As Integer

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .InitialFileName = InitialFileName
        .Title = Title
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

        On Error GoTo SELECT_ERR
        Sheets(SourceSheet).Select
        On Error GoTo 0

        If ShowAllData = True Then
            ActiveSheet.AutoFilterMode = False
            ActiveSheet.Rows.Hidden = False
            ActiveSheet.Columns.Hidden = False
            If ActiveSheet.ListObjects.Count > 0 Then
                For i = 1 To ActiveSheet.ListObjects.Count
                    ActiveSheet.ListObjects(i).Unlist
                Next
            End If
        End If

        Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange

        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close
    Else
        Err.Raise 18, "UserImportFile", "User canceled import"
    End If
    Application.DisplayAlerts = PrevDispAlerts
    Exit Sub

SELECT_ERR:
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlerts
    Err.Raise 50001, "UserImportFile", "Sheet """ & SourceSheet & """ does not exist"
End Sub
