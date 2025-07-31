Attribute VB_Name = "main"
Option Compare Database

Public Sub ExportQueryToFile(formatType As String)
' data export routine
' relies on query qSEL_export

    Dim fDialog As Object
    Dim exportPath As String
    Dim fallbackPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim queryName As String

    queryName = "qSEL_export"
    fallbackPath = Environ("USERPROFILE") & "\Documents"

    ' Folder picker dialog (late binding — no reference needed)
    Set fDialog = Application.FileDialog(4) ' 4 = Folder Picker
    With fDialog
        .Title = "Select Folder to Save Exported File"
        On Error Resume Next
        .InitialFileName = "C:\"
        .Show
        On Error GoTo 0

        If .SelectedItems.Count = 0 Then
            MsgBox "Export cancelled.", vbExclamation
            Exit Sub
        End If

        exportPath = .SelectedItems(1)
    End With

    If Len(exportPath) = 0 Then
        exportPath = fallbackPath
        MsgBox "Using fallback folder: " & exportPath, vbInformation
    End If

    ' Choose filename and export
    If formatType = "CSV" Then
        fileName = "CNS_omData_" & Format(Now(), "yyyymmdd_hhnnss") & ".csv"
        fullPath = exportPath & "\" & fileName

        DoCmd.TransferText acExportDelim, , queryName, fullPath, True

    ElseIf formatType = "Excel" Then
        fileName = "CNS_omData_" & Format(Now(), "yyyymmdd_hhnnss") & ".xlsx"
        fullPath = exportPath & "\" & fileName

        DoCmd.TransferSpreadsheet _
            TransferType:=acExport, _
            SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
            TableName:=queryName, _
            fileName:=fullPath, _
            HasFieldNames:=True
    End If

    MsgBox "Exported successfully to:" & vbCrLf & fullPath, vbInformation
End Sub


Public Function HideUI()
' hide UI - run from autoexec macro on startup

    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide
End Function

Public Function DevRestoreUI()
' restore UI for development

    On Error Resume Next
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    DoCmd.SelectObject acTable, , True
    MsgBox "Developer UI restored.", vbInformation
End Function

Public Function IsDevAuthorized() As Boolean
' check if user can enter development password (below)

    Dim pw As String
    pw = InputBox("Enter developer password:")

    If pw = "W1tch" Then
        IsDevAuthorized = True
    Else
        MsgBox "Access denied.", vbCritical
        IsDevAuthorized = False
    End If
End Function

