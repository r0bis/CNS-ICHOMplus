VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnClearSearch_Click()
    Me.txtSearch = ""
    Me.lstResults.RowSource = ""
    Me.Requery ' reload full form recordset
    DoCmd.RunCommand acCmdRecordsGoToNew
End Sub



Private Sub cmdOpenDataEntry_Click()
' to open main data entry form with this patient preselected
    If Me.NewRecord Then
        MsgBox "Please select a patient first.", vbExclamation, "No Patient Selected"
        Exit Sub
    End If

    Dim patientID As Long
    patientID = Me.ID

    Dim formName As String
    formName = "frmDataEntryList"

    On Error Resume Next
    Dim isOpen As Boolean
    isOpen = (Not Forms(formName) Is Nothing)
    On Error GoTo 0

    If isOpen Then
        ' Form is open and addressable
        DoCmd.SelectObject acForm, formName
        Forms(formName).LoadPatient patientID
    Else
        ' Not open — pass via OpenArgs
        DoCmd.OpenForm formName:=formName, WindowMode:=acDialog, OpenArgs:=patientID
    End If
    
End Sub

Private Sub lstResults_Click()
' load the selected patient's data in form (Me)

    If Me.lstResults.ListIndex = -1 Then Exit Sub ' nothing selected

    Dim selectedID As Long
    selectedID = Me.lstResults.Column(0)

    Me.Recordset.FindFirst "ID = " & selectedID
End Sub

Private Sub txtSearch_AfterUpdate()
' search function to find records with partial match in all these fields:
' PAtient Name, RiO number, Patient (hashed) code
    
    Dim s As String
    s = Replace(Me.txtSearch, "'", "''") ' escape single quotes

    Dim q As String
    q = "SELECT ID, patientCode, patientName, RiO " & _
        "FROM patients " & _
        "WHERE patientCode LIKE '*" & s & "*' " & _
        "OR patientName LIKE '*" & s & "*' " & _
        "OR RiO LIKE '*" & s & "*' " & _
        "ORDER BY patientName;"

    Me.lstResults.RowSource = q
    Me.lstResults.Requery
    'Me.Recordset.MoveLast ' we are clearing whatever displayed in fields currently
    'Me.Recordset.MoveNext ' this moves to the "new" blank record
    If Not Me.NewRecord Then
        DoCmd.RunCommand acCmdRecordsGoToNew
    End If
End Sub

