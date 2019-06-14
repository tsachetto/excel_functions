Public Function ALTSPED(cel As String, nr As Long, vr) As String

'Como utilizar: a função ALTERA valores de qualquer coluna de uma referida linha do SPED, sem a necessidade de quebrar o arquivo em colunas.
'=ALTSPED(Célula contendo o bloco de layout, Número Coluna que vc quer, Valor a ser Alterado)
'Exemplo: =SPED(A1;6;"alt") irá retornar a linha com a sexta coluna do layout colado na célula A1 alterado para "alt".
'|C181|06|5405|218,88|0|218,88|0|||0|2|   <vira>   |C181|06|5405|218,88|0|alt|0|||0|2|

Dim posicao, inicial, final, xlen As Long
Dim tx, cr As String

cr = "|"
tx = cel
tamanho = Len(tx)

    For i = 1 To tamanho
        
        y = Mid(tx, i, 1)
            
            If y = cr Then
            
                posicao = posicao + 1
            
                    If posicao = nr Then
                        inicial = i
                    Else: End If
                    
                    If posicao = nr + 1 Then
                        final = i
                        GoTo xx
                    Else: End If
                    
            Else: End If
            
    Next i
    
xx:
p1 = Mid(cel, 1, inicial)
p2 = Mid(cel, final, 150)
ALTSPED = p1 & vr & p2

End Function
