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

Private Sub cmbAnswer_AfterUpdate()
    ' Save the record
    If Me.Dirty Then Me.Dirty = False

    ' Clear selection only now — safe to do so
    Me.Parent.lstCurrQuestionnaire = Null
    Me.Parent.lstCurrQuestionnaire.Requery
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
' loads the question answers in the subform (Me)
    ' Debug.Print "LD: Loading Question ->"
    ' Debug.Print "LD: NewRecord: " & Me.NewRecord & ", Dirty: " & Me.Dirty & ", fkQuestion: " & Nz(Me.fkQuestion, "null")
    
    ' INSERT new record and set the correct answerID
    Call PreloadSubformRecord

    ' if there was a new record, now it is Me.MewRecord = FALSE
    ' so we can proceed with an existing record from here
    ' PreloadSubformRecord made sure of that
    
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim answerSetID As Variant
    Dim answerID As Variant
    Dim newAnswerID As Variant
    Dim existingAnswerID As Variant

    answerSetID = Me.Parent.lstCurrQuestionnaire.Column(3) ' Assuming fkAnswerSet is column index 3
    ' important that answerSet ID is taken from parent form
    ' likerty answers at this point have the false 91 answer set (so user needs make a choice)

    ' --- Populate the fkAnswer in the form's record source ---
    Dim qID As Long, pID As Long
    Dim rDate As Date
    Dim qType As String

    qID = Nz(Me.fkQuestion, 0)
    pID = Nz(Me.fkPatient, 0)
    rDate = Nz(Me.responseDate, Date) ' or use Me.txtSelectedDate from main form

    If qID = 0 Or pID = 0 Or IsNull(rDate) Then
        Debug.Print "LD: Skipping LoadQuestion — one or more key fields not ready yet"
        Exit Sub
    End If
    
    
    
    'If qID = 0 Or pID = 0 Then
    '    MsgBox "No good qID or pID", vbExclamation
    '    Exit Sub ' Can't proceed without question/patient
    'End If ' ? Moved this back in line to close the If properly

    qType = questionType
    'Debug.Print "question type: " & qType

    ' ? Check if answer already exists in 'data' table
    ' but it does if we are after preload
    ' a) preload loaded new record that now exists
    ' b) master-child fields loaded an existing line from data if the indexed fields matched
    existingAnswerID = DLookup("fkAnswer", "data", _
        "fkQuestion = " & qID & _
        " AND fkPatient = " & pID & _
        " AND responseDate = #" & Format(rDate, "mm\/dd\/yyyy") & "#")

    If IsNull(existingAnswerID) Then
        Debug.Print "LD: Big trouble - answer not in data table!"
        ' this really could not have heppened if master-child link works
        ' leae code in for now but don't expect it ti be triggered
        Exit Sub
    End If

    ' --- Now show/hide appropriate controls ---
    Select Case questionType
        Case "Likert", "Ordinal", "Binary"
            If Not IsNull(answerSetID) Then
                Me.cmbAnswer.RowSource = _
                    "SELECT ID, answerText, answerScore FROM answers " & _
                    "WHERE fkAnswerSet = " & answerSetID & " " & _
                    "ORDER BY [order];"
            Else
                Me.cmbAnswer.RowSource = "SELECT ID, answerText FROM answers WHERE 1=0;"
                'Debug.Print "trouble populating combo box!"
            End If

            With Me.cmbAnswer
                .RowSourceType = "Table/Query"
                .ColumnCount = 3
                .BoundColumn = 1
                .ColumnWidths = "0cm;5cm;1cm"
                .Requery
            End With
            'Debug.Print "cmbAnswer.Value 1: "; Nz(Me.cmbAnswer.Value, "NULL")
            Me.cmbAnswer.Visible = True
            'Debug.Print "cmbAnswer.Value 2: "; Nz(Me.cmbAnswer.Value, "NULL")
            Me.cmbAnswer.SetFocus
            'Debug.Print "cmbAnswer.Value 3: "; Nz(Me.cmbAnswer.Value, "NULL")
            Me.txtAnswer.Visible = False
            Me.txtVAS.Visible = False
            Me.txtNumeric.Visible = False

        Case "Text"
            Me.txtAnswer.Visible = True
            Me.txtAnswer.SetFocus
            Me.cmbAnswer.Visible = False
            Me.txtVAS.Visible = False
            Me.txtNumeric.Visible = False

        Case "VAS"
            Me.txtVAS.Visible = True
            Me.txtVAS.SetFocus
            Me.cmbAnswer.Visible = False
            Me.txtAnswer.Visible = False
            Me.txtNumeric.Visible = False
            
        Case "Numeric"
            Me.txtNumeric.Visible = True
            Me.txtNumeric.SetFocus
            Me.txtAnswer.Visible = False
            Me.cmbAnswer.Visible = False
            Me.txtVAS.Visible = False

        

        Case Else
            MsgBox "Unknown question type: " & questionType, vbExclamation
    End Select

    ' ? Requery listbox on parent (if it shows live changes)
    ' Debug.Print "LD: Load finished -> requery listBox"
    Me.Parent.lstCurrQuestionnaire.Requery

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then ' Duplicate index; can fire for new record only - by definition
        MsgBox "DATA ERR: You have already answered this question for this patient on this date.", vbExclamation
        Response = acDataErrContinue
    End If
End Sub

Private Sub PreloadSubformRecord()
' if new record needed, run this

    Dim qID As Long, pID As Long
    Dim rDate As Date
    
    ' Collect expected identity keys from parent
    qID = Nz(Me.Parent.lstCurrQuestionnaire.Column(0), 0)
    pID = Nz(Me.Parent.cmbPatient, 0)
    rDate = Nz(Me.Parent.txtSelectedDate, Date)
    ' Only runs if subform is on a new record
    
    If DCount("*", "data", _
        "fkQuestion = " & Nz(Me.fkQuestion, 0) & _
        " AND fkPatient = " & Nz(Me.fkPatient, 0) & _
        " AND responseDate = #" & Format(rDate, "mm\/dd\/yyyy") & "#") > 0 Then
        ' Debug.Print "PRL: Matching record already exists — skip preload"
        Exit Sub
    End If

    ' Defensive check - how can it be we don't have a question or the patient?
    If qID = 0 Or pID = 0 Then
        Debug.Print "PRL: Missing qID or pID — skipping preload"
        Exit Sub
    End If

    ' Safety check — maybe Access just hasn't loaded the record yet
    If DCount("*", "data", _
        "fkQuestion = " & qID & _
        " AND fkPatient = " & pID & _
        " AND responseDate = #" & Format(rDate, "mm\/dd\/yyyy") & "#") > 0 Then
        Debug.Print "PRL: Matching record already exists — skip preload"
        Exit Sub
    End If

    ' Populate required fields in (Me) record set
    Me.fkQuestion = qID
    Me.fkPatient = pID
    Me.responseDate = rDate

    ' Set fkAnswer based on question type (from parent form)
    Dim qType As String, answerSetID As Variant, answerID As Variant
    
    qType = Nz(Me.Parent.lstCurrQuestionnaire.Column(2), "")
    answerSetID = DLookup("fkAnswerSet", "questions", "ID = " & qID)

    ' set the expectable answer options - controls will be shown from LoadQuestion()
    Select Case qType
        Case "Text", "VAS", "Numeric"
            If Not IsNull(answerSetID) Then
                answerID = DLookup("ID", "answers", "fkAnswerSet = " & answerSetID)
                If Not IsNull(answerID) Then Me.fkAnswer = answerID
            End If
        Case "Likert", "Ordinal", "Binary"
            Me.fkAnswer = 91 ' Placeholder answer (e.g. "<CHOOSE ANSWER>")
    End Select

    ' Commit record (it will now exist in the data table)
    Me.Dirty = False
    ' Debug.Print "PRL: Preloaded new record: q=" & qID & ", p=" & pID & ", a=" & Me.fkAnswer
    
End Sub

Private Sub txtVAS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Cancel Enter
        ' Move focus explicitly (or ignore)
        ' prevents inserting a new record just based on VAS focus + Enter accident
    End If
End Sub


