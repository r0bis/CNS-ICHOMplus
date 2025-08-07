Attribute VB_Name = "tech"
' this module is for some technical procedures that help in development


Public Sub printTableFields()
' prints table fields and descriptions in Immediate window
' depends on function HasProperty()
' change table name in recordset open line
    
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    
    Set rs = CurrentDb.OpenRecordset("answers")
    For Each fld In rs.Fields
        If HasProperty(fld, "description") Then
            Debug.Print fld.Name & " : " & fld.Properties("description").Value
        Else
            Debug.Print fld.Name & " ---- "
        End If
    Next fld
End Sub

Public Function HasProperty(Obj As Object, propName As String) As Variant
' needed for printTableFields

On Error GoTo errLbl
    Dim prpValue As String
    prpValue = Obj.Properties(propName).Value
    HasProperty = True
    Exit Function
errLbl:
    If Not Err.Number = 3270 Then
        Debug.Print Err.Number & Err.Description
    End If
End Function

Public Function IsFormReallyOpen(formName As String) As Boolean
' complex way to check if a form is open (maybe in the background)
    On Error Resume Next
    IsFormReallyOpen = (Not Forms(formName) Is Nothing)
    On Error GoTo 0
End Function

Sub TestMD5()
' for generating MD5 hashes
' needs .NET framework available on windows. Most often it is available
    Dim enc As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim s As String
    Dim i As Long
    Dim result As String

    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    s = "test123"
    bytes = StrConv(s, vbFromUnicode)
    hash = enc.ComputeHash_2(bytes)

    For i = 0 To UBound(hash)
        result = result & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    Debug.Print result ' <-- outputs MD5 hash to Immediate Window
End Sub

Public Function GetMD5HashShort(sText As String, Optional nChars As Integer = 8) As String
    Dim enc As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Long
    Dim fullHash As String

    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    bytes = StrConv(sText, vbFromUnicode)
    hash = enc.ComputeHash_2(bytes)

    For i = 0 To UBound(hash)
        fullHash = fullHash & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    GetMD5HashShort = Left(fullHash, nChars)
    
    ' “Patient identifiers are pseudonymised using an irreversible one-way MD5 hash.
    ' Only a truncated portion of the hash is stored (e.g. 8 hex characters), which
    ' provides a unique and non-reversible code for each patient.
    ' The original identifier cannot be derived from the stored hash.”

End Function

Public Sub ExportAllModules()
' Exports VBA from modules and forms and reports (if such exist)

    Dim vbComp As Object
    Dim exportPath As String
    
    
    'exportPath = "C:\tmp\"  ' Change this

    ' user choice where to export - preferably an empty directory
    Set fDialog = Application.FileDialog(4) ' 4 = Folder Picker
    With fDialog
        .Title = "Select Folder to Save Exported File"
        On Error Resume Next
        .InitialFileName = "C:\tmp\"
        .Show
        On Error GoTo 0

        If .SelectedItems.Count = 0 Then
            MsgBox "Export cancelled.", vbExclamation
            Exit Sub
        End If

        exportPath = .SelectedItems(1)
    End With
    
    'if path does not end with backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"

    ' Ensure folder exists
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        'Debug.Print exportPath
        Select Case vbComp.Type
            Case 1, 2, 3  ' Standard, Class, or Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 100  ' Forms/Reports
                vbComp.Export exportPath & vbComp.Name & ".vba"
        End Select
    Next vbComp

    MsgBox "All modules exported to: " & exportPath
End Sub
