Option Explicit

Public Function FormatDate(ByVal dtmDate)
    Dim strYear, strMonth, strDay
    strYear = CStr(Year(dtmDate))
    strMonth = PadLeft(CStr(Month(dtmDate)), 2, "0")
    strDay = PadLeft(CStr(Day(dtmDate)), 2, "0")
    FormatDate = strYear & "-" & strMonth & "-" & strDay
End Function

