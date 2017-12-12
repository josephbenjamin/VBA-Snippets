Sub Beep_JH(ii)
' This script causes excel to say Beep ii times. It is usefull for monitoring
' a scripts progress when the excel worksheet would usually appear
' unresponsive. Since the alert is audible, one can continue working on
' something else while the task runs in the background and be alerted to
' completion, for example.

' The String can be changed to anything to get excel to say whatever you want.
Dim i As Integer
Dim N_beeps As String
Dim a_beep As String

a_beep = "Beep "

For i = 1 To ii
    N_beeps = N_beeps & a_beep
Next i

'Debug.Print (N_beeps)

For i = 1 To i_
    Application.Speech.Speak ("Beep")
Next i


End Sub
