Option Explicit

Private Sub EchoWScriptProperties()
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

Private Function Main()
    Call EchoWScriptProperties()
    Main = 0
End Function

WScript.Quit(Main())

