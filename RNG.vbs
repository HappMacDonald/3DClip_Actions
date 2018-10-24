' I usually run this action with keyboard shortcut CTRL+ALT + R

Dim max,result
max=InputBox("Max Value?", "Random Number Generator", "1e9")
If(max <> "") Then 'Detects cancel click, silently exits
  If((not IsNull(max)) and IsNumeric(max)) Then
    If(max>=1) Then
      Randomize
      result = CStr(Int(CLng(max)*Rnd))
      MsgBox("Random Number selected, and placed in your clipboard! :D"& vbCrLf & vbCrLf & CStr(result))
      clipboard.value = result
    Else
      MsgBox("Sorry, we need max to be one or greater, not "& CStr(max) &".")
    End If
  Else
    MsgBox("Sorry.. """& CStr(max) &""" isn't really a valid max value.")
  End If
End If
