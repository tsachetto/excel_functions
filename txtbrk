Public Function TXTBRK(cel As String, cr As String, nr As Integer) As String


'Como utilizar:

' =TXTBRK(Célula contendo o bloco de layout, Caractere de quebra de colunas do layout, Número Coluna que vc quer)
'Exemplo: =TXTBRK(A1;"|";5) irá retornar a quinta coluna do layout colado na célula A1.


Dim plus, plusa, plusb, xlen As Integer
Dim tx As String

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
    
TXTBRK = Mid(cel, plusa + 1, plusb - plusa - 1)

End Function
