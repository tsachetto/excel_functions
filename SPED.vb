Public Function SPED(cel As String, nr As Long) As String

'Como utilizar: a função extrai valores de qualquer coluna de uma referida linha do SPED, sem a necessidade de quebrar o arquivo em colunas.
' =SPED(Célula contendo o bloco de layout, Número Coluna que vc quer)
'Exemplo: =SPED(A1;5) irá retornar a quinta coluna do layout colado na célula A1.

Dim plus, plusa, plusb, xlen As Long
Dim tx, cr As String

cr = "|"
tx = cel
xlen = Len(tx)

    For i = 1 To xlen
    
        y = Mid(tx, i, 1)
            
            If y = cr Then
            
                plus = plus + 1
            
                    If plus = nr Then
                        plusa = i
                    Else: End If
                    
                    If plus = nr + 1 Then
                        plusb = i
                        GoTo xx
                    Else: End If
            
            Else: End If
            
    Next i
    
xx:
    
SPED = Mid(cel, plusa + 1, plusb - plusa - 1)

End Function