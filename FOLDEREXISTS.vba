Function FOLDEREXISTS(strFolderPath As String) As Boolean

'Define se um diretório existe, ex: =FOLDEREXISTS("C:\TESTE"), retornando Verdadeiro ou Falso.

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FOLDEREXISTS(strFolderPath) Then FOLDEREXISTS = True
    Set objFSO = Nothing

End Function
