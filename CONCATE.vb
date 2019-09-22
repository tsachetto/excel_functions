Public Function CONCATE(ByRef Separador As String, ByRef Area As Range, ByRef Ref As Integer) As String

'Serve para concatenar ranges, sendo:
'CONCATE( "Defina um separador" ; "Defina uma range" ; "Defina a Forma 0, 1, 2 ou 3") 
'Formas:
'0 = Concatenar considerando campos vazios, sem separador nas bordas laterais		1| |2
'1 = Concatenar desprezando campos vazios, sem separadores nas laterais		1|2
'2 = Concatenar considerando campos vazios, com separadores nas laterais		|1||2|
'3 = Concatenar desprezando campos vazios, com separadores nas laterais		|1|2|

Dim tex as string, sep as string, virg As String

If Ref = 0 Then
   
    For Each cell In Area.Value
    
        tex = tex & Separador & cell
    
    Next

CONCATE = Replace(tex, Separador, "", 1, 1)

ElseIf Ref = 1 Then

    For Each cell In Area.Value
     If Len(cell) > 0 Then
        tex = tex & Separador & cell
     Else
     End If
    Next

CONCATE = Replace(tex, Separador, "", 1, 1)

ElseIf Ref = 2 Then

    For Each cell In Area.Value
    
        tex = tex & Separador & cell
        
    Next

CONCATE = tex & Separador

ElseIf Ref = 3 Then

    For Each cell In Area.Value
     If Len(cell) > 0 Then
        tex = tex & Separador & cell
     Else
     End If
    Next

CONCATE = tex & Separador

Else
End If


End Function
