Sub DeleteTargetRows()
' Select a range in excel, and then run this script to delete the 
' *entire row* of where any cell contains the target string.
' There is no undo...Use with care.

Dim cell As Range
Dim target As String

'Delete rows containing target
target = "Exclude"
For Each cell In Selection
    If InStr(1, cell, target, vbTextCompare) > 0 Then
        cell.EntireRow.Delete
    End If
Next
End Sub
