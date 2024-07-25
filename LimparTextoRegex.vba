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

