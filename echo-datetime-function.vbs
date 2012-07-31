Option Explicit

Private Sub EchoDateTimeFunction()
    WScript.Echo("Echo VBScript Date/Time Function result")
    Dim dtmToday, dtmNow
    dtmToday = Date()
    WScript.Echo("Date()         :" & dtmToday)
    WScript.Echo("Year(today)    :" & Year(dtmToday))
    WScript.Echo("Month(today)   :" & Month(dtmToday))
    WScript.Echo("Day(today)     :" & Day(dtmToday))
    WScript.Echo("Weekday(today) :" & Weekday(dtmToday))
    Select Case Weekday(dtmToday)
        Case vbSunday
            WScript.Echo("Sunday sun sun")
        Case vbMonday
            WScript.Echo("Monday mon mon")
        Case vbTuesday
            WScript.Echo("Tuesday tue tue")
        Case vbWednesday
            WScript.Echo("Wednesday wed wed")
        Case vbThursday
            WScript.Echo("Thursday thu thu")
        Case vbFriday
            WScript.Echo("Friday fri fri")
        Case vbSaturday
            WScript.Echo("Saturday sat sat")
        Case Else
            WScript.Echo("Invalid Weekday")
    End Select
    WScript.Echo()

    dtmNow = Now()
    WScript.Echo("Now()       :" & dtmNow)
    WScript.Echo("Hour(now)   :" & Hour(dtmNow))
    WScript.Echo("Minute(now) :" & Minute(dtmNow))
    WScript.Echo("Second(now) :" & Second(dtmNow))
    WScript.Echo()
End Sub

Private Function Main()
    Call EchoDateTimeFunction()
    Main = 0
End Function

WScript.Quit(Main())

