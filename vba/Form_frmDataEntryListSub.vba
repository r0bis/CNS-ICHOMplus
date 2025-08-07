VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDataEntryListSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public PendingQuestionType As String

Private Sub cmbAnswer_AfterUpdate()
    ' Save the record
    If Me.Dirty Then Me.Dirty = False

    ' Clear selection only now — safe to do so
    Me.Parent.lstCurrQuestionnaire = Null
    Me.Parent.lstCurrQuestionnaire.Requery
End Sub


Private Sub cmbAnswer_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo SafeExit

    Debug.Print "SUBFRM: BeforeUpdate  fired at " & Now & " RecordID: " & Nz(Me.ID, "Null") & " Dirty = " & Me.Dirty

    ' Safeguard against no record context
    If Me.NewRecord Then Exit Sub
    If Me.CurrentRecord <= 0 Then Exit Sub
    If Me.Recordset.EOF Or Me.Recordset.BOF Then Exit Sub

    ' Your cleaning logic here
    Call CleanDataEntryFields(Me)

SafeExit:
    If Err.Number <> 0 Then
        Debug.Print "BeforeUpdate error: " & Err.Number & " - " & Err.Description
        Resume Next
    End If
    

End Sub

Private Sub Form_Load()
' hide fields upon first loading (caused by main form)

    Me.txtAnswer.Visible = False
    Me.txtNumeric.Visible = False
    Me.txtVAS.Visible = False
    Me.cmbAnswer.Visible = False

End Sub

Private Sub txtVAS_BeforeUpdate(Cancel As Integer)
' if value in control has changed and focus is directed away
' only checks if numeric btw 0 and 100.
    
    Dim val As Variant
    val = Me.txtVAS.Value

    If Not IsNumeric(val) Or val < 0 Or val > 100 Or val <> Int(val) Then
        MsgBox "Please enter a whole number between 0 and 100.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub txtVAS_KeyPress(KeyAscii As Integer)
' also forcing only digits and backspace in this control

    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Only digits and backspace allowed
    End If
End Sub

Public Sub LoadQuestion(questionType As String)
    Debug.Print "LD: QType -> " & questionType
    Debug.Print "LD: NewRecord = " & Me.NewRecord & " Dirty = " & Me.Dirty
    Debug.Print "LD: fkQuestion=" & Nz(Me.fkQuestion, "Null"), _
                "fkPatient=" & Nz(Me.fkPatient, "Null"), _
                "responseDate=" & Nz(Me.responseDate, "Null"), _
                "RecordID=" & Nz(Me!ID, "Null")
    Debug.Print "LD: Record count in recordset: " & Me.Recordset.RecordCount
    
    Dim answerSetID As Long
    answerSetID = Nz(Me.Parent.lstCurrQuestionnaire.Column(3), 0)
    
    ' --- Show the correct input control ---
    Select Case questionType
        Case "Likert", "Ordinal", "Binary"
            If answerSetID <> 0 Then
                Me.cmbAnswer.RowSource = _
                    "SELECT ID, answerText, answerScore FROM answers " & _
                    "WHERE fkAnswerSet = " & answerSetID & " " & _
                    "ORDER BY [order];"
            Else
                Me.cmbAnswer.RowSource = "SELECT ID, answerText FROM answers WHERE 1=0;"
            End If
            Me.cmbAnswer.Visible = True
            Me.txtAnswer.Visible = False
            Me.txtVAS.Visible = False
            Me.txtNumeric.Visible = False
            
        Case "Text"
            Me.txtAnswer.Visible = True
            Me.cmbAnswer.Visible = False
            Me.txtVAS.Visible = False
            Me.txtNumeric.Visible = False
            
        Case "VAS"
            Me.txtVAS.Visible = True
            Me.cmbAnswer.Visible = False
            Me.txtAnswer.Visible = False
            Me.txtNumeric.Visible = False
            
        Case "Numeric"
            Me.txtNumeric.Visible = True
            Me.cmbAnswer.Visible = False
            Me.txtAnswer.Visible = False
            Me.txtVAS.Visible = False
            
        Case Else
            MsgBox "Unknown question type: " & questionType, vbExclamation
    End Select
    
    Debug.Print "LD: Load finished -> requery listBox"
    Me.Parent.lstCurrQuestionnaire.Requery
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then ' Duplicate index; can fire for new record only - by definition
        MsgBox "DATA ERR: You have already answered this question for this patient on this date.", vbExclamation
        Response = acDataErrContinue
    End If
End Sub



Private Sub txtVAS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Cancel Enter
        ' Move focus explicitly (or ignore)
        ' prevents inserting a new record just based on VAS focus + Enter accident
    End If
End Sub



Private Sub Form_Current()
    If Len(PendingQuestionType) > 0 Then
        Debug.Print "SUBFORM: OnCurrent calling LoadQuestion with " & PendingQuestionType
        LoadQuestion PendingQuestionType
        PendingQuestionType = "" ' reset after use so it only runs once
    End If
End Sub

Public Sub PreloadSubformRecordDirect(qID As Long, pID As Long, rDate As Date)
    On Error GoTo ErrHandler
    Debug.Print "PRELOAD: starting, Q=" & qID & " P=" & pID & " D=" & rDate
    
    Dim rs As DAO.Recordset
    Dim qType As String
    Dim answerSetID As Long
    Dim defaultAnswerID As Long
    
    ' --- Look up question type and answer set ---
    qType = Nz(DLookup("questionType", "questions", "ID = " & qID), "")
    answerSetID = Nz(DLookup("fkAnswerSet", "questions", "ID = " & qID), 0)
    
    ' --- Decide default answer ---
    Select Case qType
        Case "Binary", "Likert", "Ordinal"
            defaultAnswerID = 91   ' placeholder "choose answer"
        
        Case "Text", "Numeric", "VAS"
            defaultAnswerID = GetFirstAnswerFromSet(answerSetID)
        
        Case Else
            defaultAnswerID = 0
    End Select
    
    Debug.Print "DIRECT PRELOAD: inserting record with defaultAnswerID=" & defaultAnswerID

    ' --- Insert recors directly into data table ---
    Set rs = CurrentDb.OpenRecordset("data", dbOpenDynaset)
    rs.AddNew
    rs!fkQuestion = qID
    rs!fkPatient = pID
    rs!responseDate = rDate
    rs!fkAnswer = defaultAnswerID
    rs.Update
    rs.Close
    
    Debug.Print "DIRECT PRELOAD: finished OK"
    Exit Sub
    
ErrHandler:
    Debug.Print "DIRECT PRELOAD: error " & Err.Number & " - " & Err.Description
End Sub

Private Function GetFirstAnswerFromSet(answerSetID As Long) As Long
' gets first answer from set. Intended to apply to Text, Numeric, VAS type answers
' in case if more than one answer in that set (shouldnt be really)
' a helper function

    Dim rsAns As DAO.Recordset
    If answerSetID = 0 Then
        GetFirstAnswerFromSet = 0
        Debug.Print "PRL: error getting answer set: " & answerSetID
        Exit Function
    End If
    
    Set rsAns = CurrentDb.OpenRecordset( _
        "SELECT ID FROM answers " & _
        "WHERE fkAnswerSet = " & answerSetID & " " & _
        "ORDER BY [order] ASC", dbOpenSnapshot)
    
    If Not rsAns.EOF Then
        GetFirstAnswerFromSet = rsAns!ID
    Else
        GetFirstAnswerFromSet = 0
        Debug.Print "PRL: error getting answer set => EOF: " & answerSetID
    End If
    
    rsAns.Close
End Function


