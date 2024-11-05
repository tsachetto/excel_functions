Public Function CNPJPFALTERDATA(CNPJ As String, Coluna As Integer) As String
    
    'Busca dentro da tabela WPHD Empresa qualquer informação por CNPJ

    Dim filePath As String
    Dim lineData As String
    Dim fields() As String
    Dim fileNumber As Integer
    
    ' Caminho do arquivo
    filePath = "\\192.168.1.251\Arquivos\Controle Contabilidade\BD\Tabelas\Empresas.txt"
    
    ' Remove caracteres especiais do CNPJ fornecido
    CNPJ = Replace(Replace(Replace(Replace(CNPJ, ".", ""), "/", ""), "-", ""), " ", "")
    
    ' Define o número do arquivo
    fileNumber = FreeFile
    
    ' Tenta abrir o arquivo para leitura
    On Error GoTo ErrorHandler
    Open filePath For Input As #fileNumber
    
    ' Lê o arquivo linha por linha
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        fields = Split(lineData, ";")
        
        ' Normaliza o CNPJ ou CPF do arquivo removendo caracteres especiais
        Dim fileCNPJ As String
        fileCNPJ = Replace(Replace(Replace(Replace(fields(17), ".", ""), "/", ""), "-", ""), " ", "")
        
        ' Verifica se o campo 18 (normalizado) é igual ao CNPJ fornecido
        If fileCNPJ = CNPJ Then
            ' Verifica se o índice da coluna é válido
            If Coluna >= 1 And Coluna <= UBound(fields) + 1 Then
                CNPJPFALTERDATA = fields(Coluna - 1) ' Retorna o valor da coluna especificada
            Else
                CNPJPFALTERDATA = "Índice de coluna inválido"
            End If
            Close #fileNumber
            Exit Function
        End If
    Loop
    
    ' Caso não encontre, retorna uma mensagem
    CNPJPFALTERDATA = "-"
    
ErrorHandler:
    ' Fecha o arquivo em caso de erro
    If fileNumber > 0 Then Close #fileNumber
End Function

Function CHAMADALTERDATA(Chamada As String, Coluna As Integer) As String

    'Busca dentro da tabela WPHD Empresa qualquer informação por Chamada
    
    Dim filePath As String
    Dim lineData As String
    Dim fields() As String
    Dim fileNumber As Integer
    
    ' Caminho do arquivo
    filePath = "\\192.168.1.251\Arquivos\Controle Contabilidade\BD\Tabelas\Empresas.txt"
    
    ' Define o número do arquivo
    fileNumber = FreeFile
    
    ' Tenta abrir o arquivo para leitura
    On Error GoTo ErrorHandler
    Open filePath For Input As #fileNumber
    
    ' Lê o arquivo linha por linha
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        fields = Split(lineData, ";")
        
        ' Verifica se a primeira coluna é igual ao código de chamada fornecido
        If fields(0) = Chamada Then
            ' Verifica se o índice da coluna é válido
            If Coluna >= 1 And Coluna <= UBound(fields) + 1 Then
                CHAMADALTERDATA = fields(Coluna - 1) ' Retorna o valor da coluna especificada
            Else
                CHAMADALTERDATA = "Índice de coluna inválido"
            End If
            Close #fileNumber
            Exit Function
        End If
    Loop
    
    ' Caso não encontre, retorna uma mensagem
    CHAMADALTERDATA = "Não encontrado"
    
ErrorHandler:
    ' Fecha o arquivo em caso de erro
    If fileNumber > 0 Then Close #fileNumber
End Function
