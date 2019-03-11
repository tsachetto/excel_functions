Public Function SOMADEC(ByVal alg As String) As Long

'SOMADEC = Soma da Decomposição dos Algarismos Interno de qualquer número, exemplo: 142 = 1+4+2 = 7 

Dim tam, sm As Long
Dim nr As Integer
Dim vr(500) As Long
Dim soma As String

tam = Len(alg)

    If tam <= 1 Then
        SOMADEC = alg
        Exit Function
    Else: End If
    
continua:

    For i = 1 To tam
            
        vr(i) = Mid(alg, i, 1)
                 
    Next

sm = WorksheetFunction.Sum(vr())
soma = sm
tam = Len(soma)

If tam > 1 Then
Erase vr
alg = soma
    GoTo continua
Else: End If

SOMADEC = sm

End Function
