Option Explicit

Private Function Main()
    If Not ValidateArguments() Then
        Main = 1
        Exit Function
    End If
    Dim strSrcFile, strDstDir
    strSrcFile = WScript.Arguments.Named.Item("srcfile")
    strDstDir = WScript.Arguments.Named.Item("dstdir")

    Dim strDstPath, strFileName
    strFileName = FormatDate(Date())
    strDstPath = PathCombine(strDstDir, strFileName)

    WScript.Echo("srcfile:" & strSrcFile)
    WScript.Echo("dstdir :" & strDstDir)
    WScript.Echo("dstpath:" & strDstPath)
    WScript.Echo()

    WScript.Echo("delete '" & strDstPath & "' if exists.")
    Call Delete(strDstPath)

    WScript.Echo("copy '" & strSrcFile & "' -> '" & strDstPath & "'.")
    Call Copy(strSrcFile, strDstPath)

    WScript.Echo("delete old files '" & strDstDir & "'.")
    Call DeleteOldFiles(strDstDir, 7)

    Main = 0
End Function

WScript.Quit(Main())

