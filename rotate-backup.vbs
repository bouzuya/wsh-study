Option Explicit

Public Sub Import(ByVal strFileName)
    Const ForReading = 1
    Dim objFso, objFile
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.OpenTextFile(strFileName, ForReading, False)
    ExecuteGlobal objFile.ReadAll()
    objFile.Close()
End Sub

Call Import("string.vbs")
Call Import("rotate-backup-arguments.vbs")
Call Import("rotate-backup-datetime.vbs")

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
