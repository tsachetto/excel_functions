Public Function CRIADIR(strFolderPath As String) As String

'Verifica se um diretório existe, ex: =CRIADIR("C:\TESTE"), CASO CONTRÁRIO CRIARÁ o diretório.

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FOLDEREXISTS(strFolderPath) Then
    CRIADIR = "DIR EXISTENTE"
    
    Else
        
        MkDir (strFolderPath)
            
    CRIADIR = "DIR CRIADO"
            
    End If
    
    Set objFSO = Nothing
        
End Function
