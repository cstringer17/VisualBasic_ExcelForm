Private Sub ReEnterData_Click()
Sheet2.Unprotect Password:="Password"
Sheet2.Activate
Range("A1").Select
Do Until IsEmpty(ActiveCell)
    Dim IDComboValue As Integer
    Dim ActiveCellValue As Integer
    
    IDComboValue = IDCombo.Value
    ActiveCellValue = ActiveCell.Value
    
    
    If IDComboValue = ActiveCellValue Then
        ActiveCell.Offset(0, 1).Value = TitleTextBox.Value
        ActiveCell.Offset(0, 2).Value = BusinessCombo.Value
        ActiveCell.Offset(0, 3).Value = ActionCombo.Value
        ActiveCell.Offset(0, 4).Value = DescText.Value
        ActiveCell.Offset(0, 5).Value = CategoryCombo.Value
    End If
    ActiveCell.Offset(1, 0).Select
Loop
Sheet2.Protect Password:="Password"
End Sub

Private Sub IDCombo_Change()
Dim currentValue As String
currentValue = IDCombo.Value
'Fill current Values
Sheet2.Unprotect Password:="Password"
Sheet2.Activate
Range("A1").Select
Do Until IsEmpty(ActiveCell)
    If currentValue = ActiveCell.Value Then
        TitleTextBox.Value = ActiveCell.Offset(0, 1)
        BusinessCombo.Value = ActiveCell.Offset(0, 2)
        ActionCombo.Value = ActiveCell.Offset(0, 3)
        DescText.Value = ActiveCell.Offset(0, 4)
        CategoryCombo.Value = ActiveCell.Offset(0, 5)
    End If
    ActiveCell.Offset(1, 0).Select
Loop

Sheet2.Protect Password:="Password"
End Sub

Private Sub SubmitForm_Click()
'Validate Data
If Len(DescText) > 32 Then
    MsgBox "The description is too long, max 32 characters", vbOKOnly, "Input Error"
    If Len(DescText) > 32 Then
        DescText.SetFocus
    End If
    Exit Sub
End If

Dim emptyRow As Long

Sheet2.Unprotect Password:="Password"
Sheet2.Activate

emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

Dim IDGen As Integer
IDGen = 0
Range("A1").Select
Do Until IsEmpty(ActiveCell)
    If ActiveCell = IDGen Then
    IDGen = IDGen + 1
    End If
    ActiveCell.Offset(1, 0).Select
Loop
Cells(emptyRow, 1).Value = IDGen
Cells(emptyRow, 2).Value = TitleTextBox.Value
Cells(emptyRow, 3).Value = BusinessCombo.Value
Cells(emptyRow, 4).Value = ActionCombo.Value
Cells(emptyRow, 5).Value = DescText.Value
Cells(emptyRow, 6).Value = CategoryCombo.Value


Sheet2.Protect Password:="Password"
End Sub


Private Sub UserForm_Initialize()

ShowTitle = 0
ValueIsNull = False


TitleTextBox.Value = ""

BusinessCombo.Clear

With BusinessCombo
    .AddItem "CSV"
    .AddItem "GCP"
    .AddItem "GLP"
    .AddItem "GMP"
    .AddItem "Complaint"
End With

DescText.Value = ""

ActionCombo.Clear

With ActionCombo
    .AddItem "Corrective Action"
    .AddItem "Preventative Action"
    .AddItem "Correction"
End With

CategoryCombo.Clear

With CategoryCombo
    .AddItem "Archive"
    .AddItem "Calibration"
    .AddItem "Change Management"
    .AddItem "Computer System validation"
    .AddItem "Equipment"
    .AddItem "Facility"
    .AddItem "Follow up of previous inspection findings"
    .AddItem "Incident Management"
    .AddItem "Inspection"
    .AddItem "Inspection Findings"
    .AddItem "Maintenance"
    .AddItem "Master Schedule"
    .AddItem "Method Description"
    .AddItem "Missing Document"
    .AddItem "N / A"
    .AddItem "Organisation / Management"
    .AddItem "Process Level"
    .AddItem "QS Documents"
    .AddItem "Qualification"
    .AddItem "Quality Assurance"
    .AddItem "Raw Data / Records"
    .AddItem "Report"
    .AddItem "Sample Handling "
    .AddItem "Site Level"
    .AddItem "System / Study Documentation"
    .AddItem "Test / Reference Item"
    .AddItem "Training"
    .AddItem "Training Management"
    .AddItem "User Management"
    .AddItem "Validation Document"
    
End With
Sheet2.Unprotect Password:="Password"
Sheet2.Activate
Range("A1").Select
Do Until IsEmpty(ActiveCell)
    IDCombo.AddItem ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Loop

Sheet2.Protect Password:="Password"
End Sub
