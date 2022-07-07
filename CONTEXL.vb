Public Function CONTEXL(rng As Range, vlr) As Integer
'Function CONTEXL = Contar Registros Exclusivos
'Como usar:
'=CONTEXL(A1:A50;VERDADEIRO) >>> Irá resultar na contagem de VERDADEIRO dentro da range
'Caso tenha 3 células escritas verdadeiro dentro da range, irá retornar 3.
Dim qt As Long
qt = 0
    For Each cell In rng
        If cell = vlr Then
        qt = qt + 1
        Else: End If
    Next
CONTEXL = qt
End Function
