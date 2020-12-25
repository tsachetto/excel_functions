'Módulo compilado de funções SPED
'Basta copiar tudo e colar em um novo módulo VBA no MS Excel


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

Public Function ALTSPED(cel As String, nr As Long, vr) As String

'Como utilizar: a função ALTERA valores de qualquer coluna de uma referida linha do SPED, sem a necessidade de quebrar o arquivo em colunas.
'=ALTSPED(Célula contendo o bloco de layout, Número Coluna que vc quer, Valor a ser Alterado)
'Exemplo: =SPED(A1;6;"alt") irá retornar a linha com a sexta coluna do layout colado na célula A1 alterado para "alt".
'|C181|06|5405|218,88|0|218,88|0|||0|2|   <vira>   |C181|06|5405|218,88|0|alt|0|||0|2|

Dim posicao as Integer
Dim inicial as Integer
Dim final as Integer
Dim tamanho As Long
Dim tx as String
Dim cr As String

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

Public Function CONCATE(ByRef Separador As String, ByRef Area As Range, ByRef Ref As Integer) As String

'Serve para concatenar ranges, sendo:
'CONCATE( "Defina um separador" ; "Defina uma range" ; "Defina a Forma 0, 1, 2 ou 3") 
'Formas:
'0 = Concatenar considerando campos vazios, sem separador nas bordas laterais		1| |2
'1 = Concatenar desprezando campos vazios, sem separadores nas laterais		1|2
'2 = Concatenar considerando campos vazios, com separadores nas laterais		|1||2|
'3 = Concatenar desprezando campos vazios, com separadores nas laterais		|1|2|

Dim tex as string, sep as string, virg As String

If Ref = 0 Then
   
    For Each cell In Area.Value
    
        tex = tex & Separador & cell
    
    Next

CONCATE = Replace(tex, Separador, "", 1, 1)

ElseIf Ref = 1 Then

    For Each cell In Area.Value
     If Len(cell) > 0 Then
        tex = tex & Separador & cell
     Else
     End If
    Next

CONCATE = Replace(tex, Separador, "", 1, 1)

ElseIf Ref = 2 Then

    For Each cell In Area.Value
    
        tex = tex & Separador & cell
        
    Next

CONCATE = tex & Separador

ElseIf Ref = 3 Then

    For Each cell In Area.Value
     If Len(cell) > 0 Then
        tex = tex & Separador & cell
     Else
     End If
    Next

CONCATE = tex & Separador

Else
End If


End Function
