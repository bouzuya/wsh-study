Option Explicit

Private Function Main()
    Dim srcfile, dstdir
    If Not WScript.Arguments.Named.Exists("srcfile") Then
        WScript.Echo("srcfile is required.")
        Main = 1
        Exit Function
    End If
    srcfile = WScript.Arguments.Named.Item("srcfile")
    If Not WScript.Arguments.Named.Exists("dstdir") Then
        WScript.Echo("dstdir is required.")
        Main = 1
        Exit Function
    End If
    dstdir = WScript.Arguments.Named.Item("dstdir")

    WScript.Echo("Hello, WSH!")
    WScript.Echo()

    WScript.Echo("srcfile:" & srcfile)
    WScript.Echo("dstdir :" & dstdir)
    WScript.Echo()

    Main = 0
End Function

WScript.Quit(Main())
