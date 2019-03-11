Public Function POWERTRIM(ByVal valor As String) As String

'Selecione uma string com espaços duplos e extras que ele apara tudo, tanto nas laterais quanto no meio.

Dim result As String
Dim lenresult As Integer

result = Replace(valor, "  ", " ")

    For i = 1 To 10
    
        result = Replace(result, "  ", " ")
        
    Next i
    
        lenresult = Len(result)
    
    For j = 1 To 10
    
        If Left(result, 1) = " " Then
            result = Mid(result, 2, lenresult - 1)
        Else: End If
        
        If Right(result, 1) = " " Then
            result = StrReverse(result)
            result = Mid(result, 2, lenresult - 1)
            result = StrReverse(result)
        Else: End If
        

    Next j

POWERTRIM = result

End Function
