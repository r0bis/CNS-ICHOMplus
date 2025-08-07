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

'--- Get the field type for a bound control ---
' we get the bound (table) field type

Private Function GetFieldType(frm As Form, ctrl As Control) As Long
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    
    On Error Resume Next
    If Len(ctrl.ControlSource) > 0 Then
        Set rs = frm.RecordsetClone
        Set fld = rs.Fields(ctrl.ControlSource)
        GetFieldType = fld.Type   ' DAO DataTypeEnum constant
    End If
    On Error GoTo 0
End Function

'--- Clean a single bound control based on its field type ---
Private Sub CleanFieldToNullByType(frm As Form, ctrl As Control)
    Dim fldType As Long
    fldType = GetFieldType(frm, ctrl)
    
    ' If no type found, skip (unbound or not in recordset)
    If fldType = 0 Then Exit Sub
    
    Select Case fldType
        ' Text and Memo (Long Text) fields
        Case dbText, dbMemo
            If Trim(Nz(ctrl.Value, "")) = "" Then
                ctrl.Value = Null
            End If
        
        ' All numeric field types
        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDecimal
            If Trim(Nz(ctrl.Value, "")) = "" Then
                ctrl.Value = Null
            End If
        
        ' Other types (Date/Time, Yes/No, etc.) — ignored
        Case Else
            ' do nothing
    End Select
End Sub

'--- Loop through relevant bound controls and clean them ---
' we only operate on controls bound to specific table fields CONST
Public Sub CleanDataEntryFields(frm As Form)
    Dim ctl As Control
    Dim fieldName As String
    Debug.Print "start field processing from BeforeUpdate - only for text control fields updated "
    
    ' Only clean these three bound fields in the data table
    Const FIELDS_TO_CLEAN As String = "|numericValue|vasValue|freeText|"
    
    For Each ctl In frm.Controls
        'we only check text boxes
        If ctl.ControlType = acTextBox Then
            If Len(ctl.ControlSource) > 0 And Nz(ctl.Value, "") <> Nz(ctl.OldValue, "") Then
                fieldName = ctl.ControlSource
                ' Only clean if the bound field is in our list
                If InStr(1, FIELDS_TO_CLEAN, "|" & fieldName & "|", vbTextCompare) > 0 Then
                    Debug.Print "entering field processing for " & fieldName
                    CleanFieldToNullByType frm, ctl
                End If
            End If
        End If
    Next ctl
End Sub
