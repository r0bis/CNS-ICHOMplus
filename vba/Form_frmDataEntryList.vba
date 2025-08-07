VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDataEntryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Private Sub btnConfirmNewDate_Click()
' what happens after new date is entered
' this is how one is meant to add data
' 0) choose patient, 1) set date, 2) choose which questionnaire to work on

    If Not IsNull(Me.txtNewDate) Then
        Me.cmbDate = Null
        Me.cmbDate.Enabled = False
        Me.txtSelectedDate = Me.txtNewDate
        Call PopulateQuestionnaireList
        Me.cmbQuestionnaire.Enabled = True
        Me.btnValidateQuestionnaire.Enabled = True
        
        Call ClearList
        
    End If
End Sub


Private Sub btnValidateQuestionnaire_Click()
' opportunity to check if data filled out correctly
' 1) all questions answered
' 2) no answers of type <CHOOSE ANSWER>
' 3) all answers within the question's answer set

    Dim qID As Long, pID As Long
    Dim rDate As Date
    Dim qCount As Long, validCount As Long
    Dim sql As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Get context
    qID = Nz(Me.cmbQuestionnaire, 0)
    pID = Nz(Me.cmbPatient, 0)
    rDate = Nz(Me.txtSelectedDate, Date)

    If qID = 0 Or pID = 0 Then
        MsgBox "Please select a questionnaire and patient.", vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb

    ' Count active questions
    qCount = DCount("*", "questions", "fkQuestionnaire = " & qID & " AND status = 'active'")

    ' Build SQL to count valid answers
    sql = _
        "SELECT COUNT(*) AS validCount " & _
        "FROM ((questions AS q " & _
        "INNER JOIN data AS d ON q.ID = d.fkQuestion) " & _
        "INNER JOIN answers AS a ON d.fkAnswer = a.ID) " & _
        "WHERE q.fkQuestionnaire = " & qID & " " & _
        "AND q.status = 'active' " & _
        "AND d.fkPatient = " & pID & " " & _
        "AND d.responseDate = #" & Format(rDate, "mm\/dd\/yyyy") & "# " & _
        "AND d.fkAnswer <> 91 " & _
        "AND a.fkAnswerSet = q.fkAnswerSet"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        validCount = rs!validCount
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Report
    If validCount = qCount Then
        MsgBox "GOOD: All active questions have valid answers.", vbInformation
    Else
        MsgBox "TROUBLE: " & (qCount - validCount) & " question(s) incomplete or invalid (O or X).", vbExclamation
    End If
    
End Sub


Private Sub cmbDate_AfterUpdate()
'reset subform not to display old data
    
    Me.txtNewDate = Null
    Me.txtSelectedDate = Me.cmbDate
    
    Call PopulateQuestionnaireList
    
    Me.cmbQuestionnaire.Enabled = True
    Me.btnValidateQuestionnaire.Enabled = True
    
    Call ClearList
    
End Sub

Private Sub cmbPatient_AfterUpdate()
' this is STEP 1 in workflow
' can be done from patients form

    'reset listBox not to display old data
    Call ClearList
    
    ' after selecting pt update row source for date box
    ' and set date values to 0 so user has to choose
    Dim sql As String
    sql = "SELECT DISTINCT responseDate FROM data WHERE fkPatient = " & Me.cmbPatient.Value & " ORDER BY responseDate DESC;"
    Me.cmbDate.RowSource = sql
    Me.cmbDate = Null   ' no value so user has to select
    Me.cmbDate.Enabled = True
    Me.txtNewDate.Enabled = True
    Me.txtNewDate = Null
    Me.txtSelectedDate.Enabled = True
    Me.txtSelectedDate = Null
    Me.cmbQuestionnaire = Null  ' so user has to choose (after date determined)
    Me.cmbQuestionnaire.Enabled = False
    Me.btnValidateQuestionnaire.Enabled = False
    Me.btnValidateQuestionnaire.Visible = True
    Me.btnConfirmNewDate.Enabled = True
    Me.btnConfirmNewDate.Visible = True
    

    
End Sub

Private Sub PopulateQuestionnaireList()
'TO populate the small combo box top left of the form (cmbQuestionnaire)

    Dim sql As String
    Dim responseDate As Variant

    If Not IsNull(Me.cmbDate) Then
        responseDate = Me.cmbDate
    ElseIf Not IsNull(Me.txtNewDate) Then
        responseDate = Me.txtNewDate
    Else
        Me.cmbQuestionnaire.RowSource = ""
        Exit Sub
    End If

    If IsNull(Me.cmbPatient) Then
        Me.cmbQuestionnaire.RowSource = ""
        Exit Sub
    End If

    If Not IsNull(Me.cmbDate) Then
        ' Case 1: Existing date selected: show only questionnaires with data
        sql = "SELECT DISTINCT qn.ID, qn.nameAbbr " & _
              "FROM (data AS d INNER JOIN questions AS q ON d.fkQuestion = q.ID) " & _
              "INNER JOIN questionnaires AS qn ON q.fkQuestionnaire = qn.ID " & _
              "WHERE d.fkPatient = " & Me.cmbPatient.Value & _
              " AND d.ResponseDate = #" & Format(responseDate, "yyyy-mm-dd") & "#" & _
              " ORDER BY qn.nameAbbr;"
    Else
        ' Case 2: New date entered: show all available questionnaires
        sql = "SELECT qn.ID, qn.nameAbbr FROM questionnaires AS qn ORDER BY nameAbbr;"
    End If

    Me.cmbQuestionnaire.RowSource = sql
    Me.cmbQuestionnaire = Null
    With Me.cmbQuestionnaire
        .ColumnCount = 2
        .BoundColumn = 1          ' ID stored
        .ColumnWidths = "0cm;5cm" ' Hide ID, show nameAbbr
        .Requery ' to be certain cb options are fresh
        .Value = Null
    End With
    
End Sub

Private Sub cmbQuestionnaire_AfterUpdate()
' this happens after the questionnaire has been selected in the combo box
    Call FilterList ' populates the listBox
    Me.btnValidateQuestionnaire.Enabled = True
End Sub

Private Sub lstCurrQuestionnaire_AfterUpdate()
    Dim qID As Long
    Dim pID As Long
    Dim rDate As Date
    Dim qType As String
    
    qID = Me.lstCurrQuestionnaire.Column(0)
    pID = Me.cmbPatient
    rDate = Me.txtSelectedDate
    qType = Me.lstCurrQuestionnaire.Column(2)
    
    '--- Ensure record exists ---
    If DCount("*", "data", _
        "fkQuestion = " & qID & _
        " AND fkPatient = " & pID & _
        " AND responseDate = #" & Format(rDate, "mm\/dd\/yyyy") & "#") = 0 Then
    ' so - if not wxist -> create one in preload Direct
        Debug.Print "MAIN: Preloading DIRECT from listbox AfterUpdate"
        Me.subAnswerPanel__.Form.PreloadSubformRecordDirect qID, pID, rDate
    End If

    ' --- Pass qType to subform ---
    Me.subAnswerPanel__.Form.PendingQuestionType = qType

    ' Refresh subform so it loads the record
    Me.subAnswerPanel__.Requery
    Debug.Print "MAIN: listBox AfterUpdate finished!"
End Sub




Private Sub lstCurrQuestionnaire_BeforeUpdate(Cancel As Integer)
' probably this is not needed now
    If Me.subAnswerPanel__.Form.Dirty Then
        If Me.subAnswerPanel__.Form.fkAnswer = 91 Then
            Me.subAnswerPanel__.Form.Undo
            Debug.Print "Undid incomplete record before switching question"
        End If
    End If
End Sub

Private Sub txtNewDate_AfterUpdate()
'review, maybe duplicating some function above
' but important - it sets txtSelectedDate hidden control
' to be the date used in further operations

    Me.txtSelectedDate = Me.txtNewDate
    Me.cmbDate = Null
End Sub

Private Sub Form_Load()
' on initial loading of the main form for data entry

    Me.cmbDate.Enabled = False
    Me.cmbDate = Null
    Me.txtNewDate.Enabled = False
    Me.txtNewDate = Null
    Me.cmbQuestionnaire.Enabled = False
    Me.cmbQuestionnaire = Null
    Me.btnValidateQuestionnaire.Enabled = False
    Me.btnValidateQuestionnaire.Visible = True
    Me.btnConfirmNewDate.Enabled = False
    Me.cmbPatient.Enabled = True
    Me.cmbPatient = Null
    Me.cmbPatient.SetFocus
    Me.lstCurrQuestionnaire.Enabled = True
    
    Call ClearList
    
    
    If Not IsNull(Me.OpenArgs) Then
        Dim ptID As Long
        ptID = CLng(Me.OpenArgs)
    
        Me.cmbPatient = ptID
        Call cmbPatient_AfterUpdate
    End If
    
    
End Sub

Private Sub FilterList()
'FOR populating listBox!!


    ' variables to hold components passed to subform
    Dim patientID, questionnaireID, responseDate As Variant
    Dim sql As String

    ' values obtained directly from controls
    patientID = Nz(Me.cmbPatient, 0)
    questionnaireID = Nz(Me.cmbQuestionnaire, 0)
    responseDate = Nz(Me.txtSelectedDate, "")

    If patientID = 0 Or questionnaireID = 0 Or Len(responseDate) = 0 Then
       'Debug.Print "ERR " & patientID & questionnaireID & Me.txtSelectedDate
        Me.lstCurrQuestionnaire.ControlSource = "SELECT * FROM data WHERE 1=0"
        Exit Sub
    End If
    

'SQL fields below (count of 6 by index 0):
    '0 = QuestionID (q.ID)
    '1 = DisplayText (q.questionCode, q.questionLong, a.answerText)
    '2 = q.questionType
    '3 = q.fkAnswerSet
    '4 = q.fkQuestionnaire
    '5 = DataID (d.ID)
    '6 = Answer status display (' ' or 'X' or 'O')
    '                           done improper notYetDone

    sql = _
        "SELECT " & _
        "    q.ID AS QuestionID, " & _
        "    q.questionCode & ' – ' & IIf(Len(q.questionLong) > 85, Left(q.questionLong,60)& ' ... ' & Right(q.questionLong, 20),q.questionLong) &" & _
        "        IIf(d.fkAnswer Is Not Null, ' [' & Nz(UCase(a.answerText), '') & ']', '') AS DisplayText, " & _
        "    q.questionType, " & _
        "    q.fkAnswerSet, " & _
        "    q.fkQuestionnaire, " & _
        "    d.ID AS DataID, " & _
        "    IIf(IsNull(d.fkAnswer), 'O', IIf(a.fkAnswerSet <> q.fkAnswerSet, 'X', '')) AS AnswerSetStatus " & _
        "FROM " & _
        "    (questions AS q " & _
        "        LEFT JOIN ( " & _
        "        SELECT * FROM data " & _
        "        WHERE fkPatient = " & Nz(Me.cmbPatient, 0) & " AND responseDate = #" & Format(Me.txtSelectedDate, "mm\/dd\/yyyy") & "# " & _
        "    ) AS d " & _
        "    ON q.ID = d.fkQuestion) " & _
        "LEFT JOIN answers AS a " & _
        "    ON d.fkAnswer = a.ID " & _
        "WHERE " & _
        "    q.fkQuestionnaire = " & Nz(Me.cmbQuestionnaire, 0) & " AND q.status = 'active' " & _
        "ORDER BY " & _
        "    q.questionCode;"

    ' Debug.Print "SQL is: " & sql
    With Me.lstCurrQuestionnaire
        .RowSourceType = "Table/Query"
        .RowSource = sql
        .ColumnCount = 7
        .BoundColumn = 1
        .ColumnWidths = "0cm;27cm;0cm;0cm;0cm;0cm;0.5cm"
        .Requery
    End With
        
    Debug.Print "MAIN: listBox count:"; Me.lstCurrQuestionnaire.ListCount
    'Debug.Print sql
        
End Sub

Private Sub ClearList()
'FOR clearing listbox

Dim sql As String

'for the listbox to preserve same column count/structure
sql = "SELECT q.ID AS QuestionID, '' AS DisplayText, '' AS questionType, " & _
        "'' AS fkAnswerSet, '' AS fkQuestionnaire, Null AS DataID " & _
        "FROM questions AS q WHERE 1 = 0;"

    With Me.lstCurrQuestionnaire
        .RowSourceType = "Table/Query"
        .RowSource = sql
        .ColumnCount = 5
        .BoundColumn = 1
        .ColumnWidths = "0cm;8cm;0cm;0cm;0cm"
        .Requery
    End With

End Sub

Public Sub LoadPatient(patientID As Long)
' exposed as LoadPatient to be sued by other code/forms (e.g. patient)

    Me.cmbPatient = patientID
    Call cmbPatient_AfterUpdate
End Sub

