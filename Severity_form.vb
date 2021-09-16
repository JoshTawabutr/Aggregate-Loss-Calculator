Private Sub Severity_clear_Click()

Sheets("Severity").Select
Columns("A:B").Select
Range("A2").Activate
Selection.ClearContents
Range("A1:B1").Select
ActiveCell.FormulaR1C1 = "Severity Distribution"
Range("A2").Select
ActiveCell.FormulaR1C1 = "x"
Range("B2").Select
ActiveCell.FormulaR1C1 = "f(x)"

End Sub

Private Sub Severity_enter_Click()

Sheets("Severity").Select
Dim this_row As Long, prob_sum As Double
this_row = 3
prob_sum = 0
Do While Cells(this_row, 1).Value <> ""
    prob_sum = prob_sum + Cells(this_row, 2)
    this_row = this_row + 1
Loop
If IsNumeric(Severity_fx_input.Value) = False Then
    response = MsgBox("The probability input must be a number!", 48, "Invalid input")
ElseIf Severity_fx_input.Value <= 0 Then
    response = MsgBox("The probability input must be positive!", 48, "Invalid input")
ElseIf prob_sum + Severity_fx_input.Value > 1 Then
    response = MsgBox("The cumulative probability exceeded 1!", 48, "Invalid input")
Else
    Cells(this_row, 1).Value = Severity_x_input.Value
    Cells(this_row, 2).Value = Severity_fx_input.Value
End If
Severity_x_input.Value = ""
Severity_fx_input.Value = ""

End Sub

Private Sub Severity_exit_Click()

Unload Severity_form

End Sub