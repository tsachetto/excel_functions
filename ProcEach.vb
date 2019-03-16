Public Function PROCEACH(ByVal onde, ByRef oque As Range) As String

'Função elaborada para pesquisar uma cadeia de caracteres dentro de outra, um a um, exemplo:
'Em uma lista de números aleatórios X definida como 3251458.
'Preciso saber se os números 148 se encontram nela. (usando como uma range, mas basta setar para 1 arg)
'A fórmula verifica de maneira independnete da ordem e retorna Verdadeirou ou Falso.
'Neste caso, os números 1, 4 e 8 estão contidos na lista X = 3251458, logo, retornaria VERDADEIRO.

Dim lenonde, lenoque As Integer
Dim ttl As Integer

lenonde = onde
    
    For Each cell In oque
        ttl = 0
        lenoque = Len(cell)
        
        On Error GoTo nexto
        
        For i = 1 To lenoque
            
            If WorksheetFunction.Find(Mid(cell, i, 1), cell, 1) > 0 Then
                
                ttl = ttl + 1
                
            Else: End If
            
        Next i
        
nexto:
    
    Next cell

If ttl = len(cell) Then
    PROCEACH = True
Else
    PROCEACH = False
End If

End Function
