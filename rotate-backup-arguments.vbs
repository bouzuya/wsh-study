Option Explicit

Public Function ValidateArguments()
    Dim objNamed
    Set objNamed = WScript.Arguments.Named
    If Not objNamed.Exists("srcfile") Then
        WScript.Echo("srcfile is required.")
        ValidateArguments = False
        Exit Function
    End If
    If Not objNamed.Exists("dstdir") Then
        WScript.Echo("dstdir is required.")
        ValidateArguments = False
        Exit Function
    End If
    Dim objFso
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If Not objFso.FileExists(objNamed.Item("srcfile")) Then
        WScript.Echo("srcfile does not exist.")
        ValidateArguments = False
        Exit Function
    End If
    If Not objFso.FolderExists(objNamed.Item("dstdir")) Then
        WScript.Echo("dstdir does not exist.")
        ValidateArguments = False
        Exit Function
    End If
    ValidateArguments = True
End Function

