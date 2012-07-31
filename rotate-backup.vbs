Option Explicit

Sub Main()
    Dim srcfile, dstdir
    If Not WScript.Arguments.Named.Exists("srcfile") Then
        WScript.Echo("srcfile is required.")
        WScript.Quit(1)
        Exit Sub
    End If
    srcfile = WScript.Arguments.Named.Item("srcfile")
    If Not WScript.Arguments.Named.Exists("dstdir") Then
        WScript.Echo("dstdir is required.")
        WScript.Quit(1)
        Exit Sub
    End If
    dstdir = WScript.Arguments.Named.Item("dstdir")

    WScript.Echo("Hello, WSH!")
    WScript.Echo()

    WScript.Echo("srcfile:" & srcfile)
    WScript.Echo("dstdir :" & dstdir)
    WScript.Echo()

    Call EchoWScriptProperties()

    WScript.Quit(0)
End Sub

Sub EchoWScriptProperties()
    WScript.Echo("WScript properties")
    WScript.Echo("WScript.Name          :" & WScript.Name)
    WScript.Echo("WScript.Version       :" & WScript.Version)
    WScript.Echo("WScript.Path          :" & WScript.Path)
    WScript.Echo("WScript.FullName      :" & WScript.FullName)
    WScript.Echo("WScript.Interactive   :" & WScript.Interactive)
    WScript.Echo("WScript.ScriptFullName:" & WScript.ScriptFullName)
    WScript.Echo("WScript.ScriptName    :" & WScript.ScriptName)
    WScript.Echo()
End Sub

Call Main()
