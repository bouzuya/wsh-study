Option Explicit

Private Function ValidateArguments()
    If Not WScript.Arguments.Named.Exists("srcfile") Then
        WScript.Echo("srcfile is required.")
        ValidateArguments = False
        Exit Function
    End If
    If Not WScript.Arguments.Named.Exists("dstdir") Then
        WScript.Echo("dstdir is required.")
        ValidateArguments = False
        Exit Function
    End If
    ValidateArguments = True
End Function

Private Function PadLeft(ByVal strValue, ByVal intTotalLength, ByVal strPadding)
    Dim strResult
    strResult = strValue
    While Len(strResult) < intTotalLength
        strResult = strPadding & strResult
    Wend
    PadLeft = strResult
End Function

Private Function FormatDate(ByVal dtmDate)
    Dim strYear, strMonth, strDay
    strYear = CStr(Year(dtmDate))
    strMonth = PadLeft(CStr(Month(dtmDate)), 2, "0")
    strDay = PadLeft(CStr(Day(dtmDate)), 2, "0")
    FormatDate = strYear & "-" & strMonth & "-" & strDay
End Function

Private Function Main()
    If Not ValidateArguments() Then
        Main = 1
        Exit Function
    End If
    Dim strSrcFile, strDstDir
    strSrcFile = WScript.Arguments.Named.Item("srcfile")
    strDstDir = WScript.Arguments.Named.Item("dstdir")

    Dim strDate
    strDate = FormatDate(Date())
    WScript.Echo("strDate: " & strDate)
    WScript.Echo()

    WScript.Echo("Hello, WSH!")
    WScript.Echo()

    WScript.Echo("srcfile:" & strSrcFile)
    WScript.Echo("dstdir :" & strDstDir)
    WScript.Echo()

    Main = 0
End Function

WScript.Quit(Main())
