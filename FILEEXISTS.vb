Function FILEEXISTS(strFileFullPath As String) As Boolean

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FILEEXISTS(strFileFullPath) Then FILEEXISTS = True
    Set objFSO = Nothing

End Function
