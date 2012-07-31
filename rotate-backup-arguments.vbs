Option Explicit

Public Function ValidateArguments()
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

