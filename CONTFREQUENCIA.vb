Public Function CONTFREQUENCIA(ByVal minimo, ByVal maximo, ByRef area As Range) As Long

'Determina a frequência de uma calsse entre limite min e max baseado em uma lista de números.
'Exemplo: =CONTFREQUENCIA(MIN;MAX;RANGE) ou =CONTFREQUENCIA(1;55;A5:A525)

Dim contador As Integer

For Each cell In area

    If cell.Value >= minimo And cell.Value <= maximo Then
    
        contador = contador + 1

    Else: End If

Next

CONTFREQUENCIA = contador

End Function
