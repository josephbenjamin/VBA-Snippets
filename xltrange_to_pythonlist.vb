Public Function ToPyList(r As Range) As String

Dim row As Range
Dim cell As Range
Dim rng As Range: Set rng = r
Dim output_string As String: output_string = ""
Dim r_cnt As Integer: r_cnt = 0
Dim c_cnt As Integer: c_cnt = 0

output_string = output_string & "["
For Each row In rng.Rows
    If r_cnt = 0 Then
        output_string = output_string & "["
    Else
        output_string = output_string & ", ["
    End If

    For Each cell In row.Cells
        If c_cnt = 0 Then
            output_string = output_string & ""
        Else
            output_string = output_string & ","
        End If

        output_string = output_string & """" & cell.Value & """"
        c_cnt = c_cnt + 1
    Next cell

    output_string = output_string & "]"
    r_cnt = r_cnt + 1
    c_cnt = 0
Next row

output_string = output_string & "]"
ToPyList = output_string
End Function
