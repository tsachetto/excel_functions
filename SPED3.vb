Public Function SPED(cel As String, qtd As Integer)

'Como utilizar: a função extrai valores de qualquer coluna de uma referida linha do SPED, sem a necessidade de quebrar o arquivo em colunas.
' =SPED(Célula contendo o bloco de layout, Número Coluna que vc quer)
'Exemplo: =SPED(A1;5) irá retornar a quinta coluna do layout colado na célula A1.

Dim posicao As Integer
Dim incial As Integer
Dim final As Integer
Dim tamanho As Integer
Dim linhasped As String
Dim separador As String

separador = "|"
linhasped = cel
tamanho = Len(linhasped)

For i = 1 To tamanho

    carac = Mid(linhasped, i, 1)
        
    If carac = separador Then
    
        posicao = posicao + 1
    
        If posicao = qtd Then
            incial = i
        Else: End If
        
        If posicao = qtd + 1 Then
            final = i
            GoTo concluir
        Else: End If
    
    Else: End If
        
Next i
    
concluir:

    SPED = Mid(cel, (incial + 1), (final - incial - 1))

End Function
