
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
        ActiveCell.Offset(0, 6).Value = RootCause.Value
        ActiveCell.Offset(0, 7).Value = Evidence.Value
        ActiveCell.Offset(0, 8).Value = ActionTaken.Value
        ActiveCell.Offset(0, 9).Value = CurrentDueDate.Value
        ActiveCell.Offset(0, 10).Value = CompletionDate.Value
        
        If Check.Value = True Then
         ActiveCell.Offset(0, 11).Value = "Yes"
        End If
        
        If Check.Value = False Then
         ActiveCell.Offset(0, 11).Value = "No"
        End If
        
        ActiveCell.Offset(0, 12).Value = EffectivenessPlan.Value
        ActiveCell.Offset(0, 13).Value = Comments.Value
        ActiveCell.Offset(0, 14).Value = ActionPlan.Value
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
        If currentValue = 0 Then
            MsgBox "You can't change this", vbOKOnly, "Input Error"
            Exit Sub
        End If
        TitleTextBox.Value = ActiveCell.Offset(0, 1)
        BusinessCombo.Value = ActiveCell.Offset(0, 2)
        ActionCombo.Value = ActiveCell.Offset(0, 3)
        DescText.Value = ActiveCell.Offset(0, 4)
        CategoryCombo.Value = ActiveCell.Offset(0, 5)
        RootCause.Value = ActiveCell.Offset(0, 6)
        Evidence.Value = ActiveCell.Offset(0, 7)
        ActionTaken.Value = ActiveCell.Offset(0, 8)
        CurrentDueDate.Value = ActiveCell.Offset(0, 9)
        CompletionDate.Value = ActiveCell.Offset(0, 10)
        If Check.Value = True Then
         ActiveCell.Offset(0, 11).Value = "Yes"
        End If
        
        If Check.Value = False Then
         ActiveCell.Offset(0, 11).Value = "No"
        End If
        
        EffectivenessPlan.Value = ActiveCell.Offset(0, 12)
        Comments.Value = ActiveCell.Offset(0, 13)
        ActionPlan.Value = ActiveCell.Offset(0, 14)
    End If
    ActiveCell.Offset(1, 0).Select
Loop

Sheet2.Protect Password:="Password"
End Sub

Private Sub SubmitForm_Click()
'Validate Data
'If Len(DescText) > 32 Then
 '   MsgBox "The description is too long, max 32 characters", vbOKOnly, "Input Error"
 '   If Len(DescText) > 32 Then
 '       DescText.SetFocus
 '   End If
 '   Exit Sub
'End If

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
Cells(emptyRow, 7).Value = RootCause.Value
Cells(emptyRow, 8).Value = Evidence.Value
Cells(emptyRow, 9).Value = ActionTaken.Value
Cells(emptyRow, 10).Value = CurrentDueDate.Value
Cells(emptyRow, 11).Value = CompletionDate.Value
If Check.Value = True Then
    ActiveCell.Offset(0, 11).Value = "Yes"
End If
        
If Check.Value = False Then
    ActiveCell.Offset(0, 11).Value = "No"
End If
        
Cells(emptyRow, 13).Value = EffectivenessPlan.Value
Cells(emptyRow, 14).Value = Comments.Value
Cells(emptyRow, 15).Value = ActionPlan.Value

Sheet2.Protect Password:="Password"

Dim ctl As MSForms.Control

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.Value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl

End Sub


Private Sub UserForm_Initialize()

ShowTitle = 0
ValueIsNull = False

TitleTextBox.Value = ""

BusinessCombo.Clear


BusinessCombo.Value = "GCP"

DescText.Value = ""

ActionCombo.Clear

With ActionCombo
    .AddItem "Corrective Action"
    .AddItem "Preventative Action"
    .AddItem "Correction"
End With

CategoryCombo.Clear

With CategoryCombo
    .AddItem "Site Level"
    .AddItem "Process Level"
    .AddItem "Study Level"
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
