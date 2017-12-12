Sub ZoomScroll()
' This script zooms a sheet to 80% and selects cell A1 on all sheets

Dim sht As Worksheet, csheet As Worksheet

Application.ScreenUpdating = False
Set csheet = ActiveSheet

For Each sht In ActiveWorkbook.Worksheets
  If sht.Visible Then
    sht.Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 80
  End If
Next sht

csheet.Activate
Application.ScreenUpdating = True

End Sub
