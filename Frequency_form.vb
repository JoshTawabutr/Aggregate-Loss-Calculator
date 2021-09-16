Private Sub Frequency_clear_Click()

Sheets("Frequency").Select
Columns("A:B").Select
Range("A2").Activate
Selection.ClearContents
Range("A1:B1").Select
ActiveCell.FormulaR1C1 = "Frequency Distribution"
Range("A2").Select
ActiveCell.FormulaR1C1 = "n"
Range("B2").Select
ActiveCell.FormulaR1C1 = "p_n"

End Sub

Private Sub Frequency_enter_Click()

Sheets("Frequency").Select
Dim this_row As Long, prob_sum As Double
this_row = 3
prob_sum = 0
Do While Cells(this_row, 1).Value <> ""
    prob_sum = prob_sum + Cells(this_row, 2)
    this_row = this_row + 1
Loop
If IsNumeric(Freq_pn_input.Value) = False Then
    response = MsgBox("The probability input must be a number!", 48, "Invalid input")
ElseIf Freq_pn_input.Value <= 0 Then
    response = MsgBox("The probability input must be positive!", 48, "Invalid input")
ElseIf prob_sum + Freq_pn_input.Value > 1 Then
    response = MsgBox("The cumulative probability exceeded 1!", 48, "Invalid input")
Else
    Cells(this_row, 1).Value = Freq_n_input.Value
    Cells(this_row, 2).Value = Freq_pn_input.Value
End If
Freq_n_input.Value = ""
Freq_pn_input.Value = ""

End Sub

Private Sub Frequency_exit_Click()

Unload Frequency_form

End Sub
