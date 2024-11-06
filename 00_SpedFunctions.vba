Public Function SPED(ln_sped As String, campo As Integer) As String

'Como usar? SERVE PARA EXTRAIR QUALQUER CAMPO DESEJADO DE UMA LINHA DE SPED
'Exemplo de linha de sped localizada na célula A5: |0000|1|00100|45454|X|TEXTO EXEMPLO|NF 1515|
'Para retornar o campo 6 da linha de sped localizada na linha 5, você pode ir na linha B5 e digitar:
'=SPED(A5,6) <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Exemplo de uso <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'O Excel retornará "TEXTO EXEMPLO" que corresponde ao campo 6 do conteúdo da célula A5.

    ' Quebra o texto em uma array de palavras
    Dim compos_sped() As String
    compos_sped = Split(ln_sped, "|")
    
    ' Retorna o valor do campo indicado pelo índice "campo"
    SPED = compos_sped(campo)
    
End Function

'-------------------------------------------------------------------------------------------------------------------------------

Function ALTSPED(ln_sped As String, campo As Integer, novo_conteudo As String) As String

'Como usar? SERVE PARA ALTERAR O CAMPO DESEJADO DENTRO DE UMA LINHA DE SPED
'Exemplo de linha de sped localizada na célula A5: |0000|1|00100|45454|X|TEXTO EXEMPLO|NF 1515|
'Para alterar o campo 6 da linha de sped localizada na linha 5, você pode ir na linha B5 e digitar:
'=ALTSPED(A5,6,"TEXTO ATUALIZADO") <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Exemplo de uso <<<<<<<<<<<<<<<<<<
'O Excel retornará "|0000|1|00100|45454|X|TEXTO ATUALIZADO|NF 1515|" que corresponde ao conteudo do campo 6 atualizado.


    ' Quebra o texto em uma array de palavras
    Dim compos_sped() As String
    compos_sped = Split(ln_sped, "|")
    
    ' Altera o campo indicado pelo índice "campo"
    compos_sped(campo) = novo_conteudo
    
    ' Concatena as palavras em uma string
    Dim resultado As String
    resultado = Join(compos_sped, "|")
    
    ' Retorna o resultado
    ALTSPED = resultado
End Function
