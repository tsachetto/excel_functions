Public Function READTexto(arquivo As String, linha As Long) As String

Dim textline As String, contador As Long
contador = 1

Open arquivo For Input As #1

    For i = 1 To 1048576
    
        Do Until EOF(1)
        
            Line Input #1, textline
            
                If contador = linha Then
                    GoTo fim
                Else: End If
                
            contador = contador + 1
        
        Loop
    Next
fim:
Close #1


READTexto = textline

End Function

