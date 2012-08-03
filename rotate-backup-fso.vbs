Option Explicit

Public Function PathCombine(ByVal strPath1, ByVal strPath2)
    Dim objFso
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    PathCombine = objFso.BuildPath(strPath1, strPath2)
End Function

Public Sub Copy(ByVal strSrcPath, ByVal strDstPath)
    Dim objFso
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(strSrcPath) Then
        Call objFso.CopyFile(strSrcPath, strDstPath)
    ElseIf objFso.FolderExists(strSrcPath) Then
        Call objFso.CopyFolder(strSrcPath, strDstPath)
    Else
        ' do nothing
    End If
End Sub

Public Sub Move(ByVal strSrcPath, ByVal strDstPath)
    Dim objFso
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(strSrcPath) Then
        Call objFso.MoveFile(strSrcPath, strDstPath)
    ElseIf objFso.FolderExists(strSrcPath) Then
        Call objFso.MoveFolder(strSrcPath, strDstPath)
    Else
        ' do nothing
    End If
End Sub

Public Sub Delete(ByVal strPath)
    Dim objFso
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(strPath) Then
        Call objFso.DeleteFile(strPath)
    ElseIf objFso.FolderExists(strPath) Then
        Call objFso.DeleteFolder(strPath)
    Else
        ' do nothing
    End If
End Sub

Public Function GetPaths(ByVal strDir)
    Dim objFso, objFolder, objFiles, objFile
    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFso.GetFolder(strDir)
    Set objFiles = objFolder.Files
    Dim objPaths(), i
    ReDim objPaths(objFiles.Count - 1)
    i = 0
    For Each objFile In objFiles
        objPaths(i) = objFile.Path
        i = i + 1
    Next
    GetPaths = objPaths
End Function

Public Function CloneArray(ByVal objArray)
    Dim objNewArray(), i
    ReDim objNewArray(UBound(objArray))
    For i = LBound(objArray) To UBound(objArray)
        objNewArray(i) = objArray(i)
    Next
    CloneArray = objArray
End Function

Public Function SortPaths(ByVal objPaths)
    Const vbBinaryCompare = 0, vbTextCompare = 1
    Dim objNewPaths, i, j
    objNewPaths = CloneArray(objPaths)
    For i = LBound(objNewPaths) To UBound(objNewPaths)
        For j = i + 1 To UBound(objNewPaths)
            If StrComp(objNewPaths(i), objNewPaths(j), vbBinaryCompare) < 0 Then
                Dim strPath
                strPath = objNewPaths(i)
                objNewPaths(i) = objNewPaths(j)
                objNewPaths(j) = strPath
            End If
        Next
    Next
    SortPaths = objNewPaths
End Function

Public Sub DeleteOldFiles(ByVal strDir, ByVal intMaxCount)
    Dim objPaths, strPath, intCount
    objPaths = SortPaths(GetPaths(strDir))
    intCount = 0
    For Each strPath In objPaths
        intCount = intCount + 1
        If intCount > intMaxCount Then
            Call Delete(strPath)
        End If
    Next
End Sub

