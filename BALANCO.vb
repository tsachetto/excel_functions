Function BALANCO(valor As String) As String

'FUNÇAO QUE ACERTA VALORES GERADOS NOS BALANÇOS PATRIMONIAIS DA ALTERDATA
'VALOR GERADO: (*****9000,00D) ALTERA-SE PARA (9000,00D)
'COMO USAR:
'=BALANCO(A1) p.exemplo

Dim tamanho As Integer
Dim i As Integer

    tamanho = Len(valor)

    For i = 1 To tamanho

        If IsNumeric(Mid(valor, i, 1)) Then
            bal = Mid(valor, i, tamanho)
            Exit Function
        Else: End If
        
    Next

End Function
