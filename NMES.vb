Public Function NMES(ByVal dt) As String

'Basta selecionar uma data ou um número inteiro entre 1 e 12 que ele retorna o nome do mês correspondente.

Dim nm As Integer

    If IsDate(dt) Then
    
        nm = Month(dt)
        
    ElseIf IsNumeric(dt) Then
    
        If dt > 0 And dt < 13 Then
        
            nm = dt
            
            GoTo continua
        
        Else: End If
        
    Else
        
        NMES = "-"
    
    End If

continua:
    
    If nm = 1 Then
        NMES = "Janeiro"
    ElseIf nm = 2 Then
        NMES = "Fevereiro"
    ElseIf nm = 3 Then
        NMES = "Março"
    ElseIf nm = 4 Then
        NMES = "Abril"
    ElseIf nm = 5 Then
        NMES = "Maio"
    ElseIf nm = 6 Then
        NMES = "Junho"
    ElseIf nm = 7 Then
        NMES = "Julho"
    ElseIf nm = 8 Then
        NMES = "Agosto"
    ElseIf nm = 9 Then
        NMES = "Setembro"
    ElseIf nm = 10 Then
        NMES = "Outubro"
    ElseIf nm = 11 Then
        NMES = "Novembro"
    ElseIf nm = 12 Then
        NMES = "Dezembro"
    Else
        NMES = "N/A"
    End If

End Function
