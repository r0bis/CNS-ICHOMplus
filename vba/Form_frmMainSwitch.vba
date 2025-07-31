VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnExit_Click()
' exit application
    DoCmd.Quit
End Sub

Private Sub btnExportCSV_Click()
' export data for further analysis as CSV
    Call ExportQueryToFile("CSV")
End Sub

Private Sub btnExportExcel_Click()
' export data for further analysis as Excel
    Call ExportQueryToFile("Excel")
End Sub

Private Sub btnPatients_Click()
' show patients form as modal
    DoCmd.OpenForm "frmPatients", WindowMode:=acDialog
End Sub

Private Sub btnQuestionnaires_Click()
' show main data aentry form as modal
    DoCmd.OpenForm "frmDataEntryList", WindowMode:=acDialog
End Sub

Private Sub btnDevUI_Click()
' show normal controls for dev interface
' (UI first hidden to avoid user interface clutter)
    'Debug.Print "Dev pressed"
    Call DevRestoreUI
 
End Sub

Private Sub Form_Load()
' restore form size and placement
    DoCmd.Restore
End Sub

Private Sub lblDevUnlock_Click()
' click to trigger showing button for access to dev interface
    If IsDevAuthorized() Then
        Me.btnDevUI.Visible = True
    End If
End Sub
