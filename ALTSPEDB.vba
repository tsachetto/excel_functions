Function ALTSPEDB(ln_sped As String, campos As Variant, novo_conteudo As String) As String

    ' Divide o texto em uma array de campos
    Dim compos_sped() As String
    compos_sped = Split(ln_sped, "|")
    
    ' Verifica se "campos" é um único valor ou uma lista
    Dim campo As Variant
    For Each campo In campos
        ' Verifica se o índice está dentro dos limites da array
        If campo > 0 And campo <= UBound(compos_sped) Then
            compos_sped(campo) = novo_conteudo
        End If
    Next campo
    
    ' Concatena os campos de volta em uma string
    Dim resultado As String
    resultado = Join(compos_sped, "|")
    
    ' Retorna o resultado com o primeiro e último delimitador "|"
    ALTSPEDB = resultado
End Function
