Option Explicit

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
