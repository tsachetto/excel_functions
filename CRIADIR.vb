Public Function CRIADIR(strFolderPath As String, modo) As String

'Verifica se um diretório existe e realiza uma ação de modo:
'Se modo for 0, apenas verifica existência do diretório/pasta.
'Se modo for 1, além de verificar, cria o diretório/pasta casa não exita.
'Exemplo modo = 0
'=CRIADIR("C:\TESTE";0) retorna existencia ou não do diretório.
'Exemplo modo = 1
'=CRIADIR("C:\TESTE";1) retorna existencia e cria o diretório caso o mesmo não exista.
'Vr 2.0 by thz
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
'Define ação de modo 0/1
    
Select Case modo
Case Is = 0
    
        If objFSO.FOLDEREXISTS(strFolderPath) Then
        CRIADIR = "DIRETÓRIO EXISTENTE"
        Else
        CRIADIR = "DIRETÓRIO INEXISTENTE"
        End If
  
Case Is = 1
    
        If objFSO.FOLDEREXISTS(strFolderPath) Then
        CRIADIR = "DIR EXISTENTE"
        
        Else
        
        MkDir (strFolderPath)
        CRIADIR = "DIR CRIADO"
                
        End If
End Select
    
    Set objFSO = Nothing
        
End Function
