' Warning! Does not yet properly handle fractional hour timezones! Hard To test for that from the US. :P

Public Function ToIsoDateTime(datetime) 
  ToIsoDateTime = ToIsoDate(datetime) & "T" & ToIsoTime(datetime) & GetTimeZoneOffset()
End Function

Public Function ToIsoDate(datetime)
  ToIsoDate = CStr(Year(datetime)) & "-" & StrN2(Month(datetime),"") & "-" & StrN2(Day(datetime),"")
End Function    

Public Function ToIsoTime(datetime) 
  ToIsoTime = StrN2(Hour(datetime),"") & ":" & StrN2(Minute(datetime),"")
End Function

Function GetTimeZoneOffset()
  Const sComputer = "."

  Dim oWmiService : Set oWmiService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
      & sComputer & "\root\cimv2")

  Set cItems = oWmiService.ExecQuery("SELECT * FROM Win32_ComputerSystem")

  For Each oItem In cItems
    GetTimeZoneOffset = oItem.CurrentTimeZone / 60
    GetTimeZoneOffset = StrN2(GetTimeZoneOffset,"+") & ":00"
    Exit For
  Next
End Function

Private Function StrN2(n,positiveSign)
  n = CInt(n)
  If n<0 Then
    sign = "-"
    n = -n
  Else
    sign = positiveSign
  End if
  If Len(n) < 2 Then StrN2 = "0" & n Else StrN2 = n
  StrN2 = sign & StrN2
  StrN2 = CStr(StrN2)
End Function

Clipboard.Value = ToIsoDateTime(NOW())
