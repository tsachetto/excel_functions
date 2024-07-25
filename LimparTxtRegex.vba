Aqui está o código VBA para uma função que utiliza expressões regulares (Regex) para considerar apenas letras, números e espaços dentro de um texto, excluindo todo o restante:

```vba
Function LimparTexto(ByVal texto As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .Pattern = "[^a-zA-Z0-9\s]"
        LimparTexto = .Replace(texto, "")
    End With
    
    Set regex = Nothing
End Function
```

Este código cria uma função chamada `LimparTexto` que recebe uma string como argumento e retorna a mesma string, mas apenas com letras, números e espaços, removendo todos os outros caracteres.
