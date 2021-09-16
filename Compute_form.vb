Private Sub Compute_aggloss_button_Click()

'Check if the form is filled correctly
Dim max_L As Long, min_L As Long
If Option_allvalues = True Then
    min_L = 0
    max_L = -1
ElseIf Option_valuesupto = True Then
    If maxL_input.Value = "" Then
        response = MsgBox("Please specify the maximum aggregate loss value.", 48, "Parameter not specified")
        Unload Compute_form
        Exit Sub
    Else
        min_L = 0
        max_L = maxL_input.Value
    End If
ElseIf Option_valuesat = True Then
    If Specific_agg_input.Value = "" Then
        response = MsgBox("Please specify the desired aggregate loss value.", 48, "Parameter not specified")
        Unload Compute_form
        Exit Sub
    Else
        min_L = Specific_agg_input.Value
        max_L = Specific_agg_input.Value
    End If
Else
    response = MsgBox("Please choose the output format.", 48, "Parameter not specified")
    Unload Compute_form
    Exit Sub
End If


'Preliminary question
Sheets("Aggregate Losses").Select
If Cells(3, 1).Value <> "" Then
    response1 = MsgBox("This will overwrite the existing entries. Continue?", vbQuestion + vbOKCancel, "Warning")
    If response1 = vbCancel Then
        Unload Compute_form
        Exit Sub
    End If
End If

'Clear the existing content
Columns("A:B").Select
Range("A2").Activate
Selection.ClearContents
Range("A1:B1").Select
ActiveCell.FormulaR1C1 = "Aggregate Loss Distribution"
Range("A2").Select
ActiveCell.FormulaR1C1 = "x"
Range("B2").Select
ActiveCell.FormulaR1C1 = "f(x)"


'Check the given frequency distribution
Sheets("Frequency").Select
Dim this_row As Long, prob_sum As Double, nrow_f As Long, max_freq As Long, nonzero_pn() As Long
this_row = 3
prob_sum = 0
max_freq = 0
Do While Cells(this_row, 1).Value <> ""
    prob_sum = prob_sum + Cells(this_row, 2)
    If max_freq < Cells(this_row, 1) Then
        max_freq = Cells(this_row, 1)
    End If
    ReDim Preserve nonzero_pn(this_row - 3)
    nonzero_pn(this_row - 3) = Cells(this_row, 1)
    this_row = this_row + 1
Loop
If prob_sum < 1 Then
    response = MsgBox("The given probability distribution does not sum to one.", 48, "Missing probability")
    Unload Compute_form
End If
nrow_f = this_row - 3



'Check the given severity distribution
Sheets("Severity").Select
Dim nrow_s As Long, max_s As Double, nonzero_fx() As Long
this_row = 3
prob_sum = 0
max_s = 0
Do While Cells(this_row, 1).Value <> ""
    prob_sum = prob_sum + Cells(this_row, 2)
    If max_s < Cells(this_row, 1) Then
        max_s = Cells(this_row, 1)
    End If
    ReDim Preserve nonzero_fx(this_row - 3)
    nonzero_fx(this_row - 3) = Cells(this_row, 1)
    this_row = this_row + 1
Loop
If prob_sum < 1 Then
    response = MsgBox("The given probability distribution does not sum to one.", 48, "Missing probability")
    Unload Compute_form
End If
nrow_s = this_row - 3

'Set max_L for the maximal output case
If max_L < 0 Then
    max_L = max_s * max_freq
End If

'Construct p_n array
Sheets("Frequency").Select
Dim pn() As Double, i As Long, j As Long, isat As Long
i = 0
ReDim pn(0)
For i = 0 To max_freq
    ReDim Preserve pn(i)
    isat = -1
    j = 0
    For j = 0 To nrow_f - 1
        If nonzero_pn(j) = i Then
            isat = j
        End If
    Next j
    If isat >= 0 Then
        pn(i) = Cells(isat + 3, 2)
    Else
        pn(i) = 0
    End If
Next i


'A special case
Sheets("Aggregate Losses").Select
If max_L = 0 Then
    Cells(3, 1).Value = 0
    Cells(3, 2).Value = pn(0)
    Unload Compute_form
    Exit Sub
End If


'Construct f(x) array
Sheets("Severity").Select
Dim fx() As Double
ReDim fx(0)
For i = 0 To max_s - 1
    ReDim Preserve fx(i)
    isat = -1
    For j = 0 To nrow_s - 1
        If nonzero_fx(j) = i + 1 Then
            isat = j
        End If
    Next j
    If isat >= 0 Then
        fx(i) = Cells(isat + 3, 2)
    Else
        fx(i) = 0
    End If
Next i
For i = max_s To max_L - 1
    ReDim Preserve fx(i)
    fx(i) = 0
Next i

'Now, let's get working with the actual computation
'Computing f_x^*
Dim fxstar() As Double, k As Long, m As Long
ReDim fxstar(max_freq, max_L - 1)
For k = 0 To max_freq
    For i = 0 To max_L - 1
        fxstar(k, i) = 0
    Next i
Next k
For k = 1 To max_freq
    For i = 0 To max_L - 1
        If k = 1 Then
            fxstar(k, i) = fx(i)
        Else
            For m = 0 To i - 1
                fxstar(k, i) = fxstar(k, i) + (fx(m) * fxstar(k - 1, i - m - 1))
            Next m
        End If
    Next i
Next k

'Computing f_S
Dim fs() As Double
ReDim fs(max_L)
fs(0) = pn(0)
For i = 1 To max_L
    fs(i) = 0
    For k = 1 To max_freq
        fs(i) = fs(i) + (pn(k) * fxstar(k, i - 1))
    Next k
Next i

'Writing down the answer
Sheets("Aggregate Losses").Select
this_row = 3
For i = min_L To max_L
    If fs(i) <> 0 Then
        Cells(this_row, 1).Value = i
        Cells(this_row, 2).Value = fs(i)
        this_row = this_row + 1
    End If
Next i

'Some scratchwork output for prelim testing
'For this_row = 3 To max_freq + 3
'    Cells(this_row, 1) = this_row - 3
'    Cells(this_row, 2) = pn(this_row - 3)
'Next this_row
'For this_row = max_freq + 4 To max_L + max_freq + 3
'    Cells(this_row, 1) = this_row - 3 - max_freq
'    Cells(this_row, 2) = fx(this_row - 4 - max_freq)
'Next this_row

'Finishing up
Unload Compute_form
resp = MsgBox("The aggregate loss probability function is displayed in the table.", , "Results")

End Sub

