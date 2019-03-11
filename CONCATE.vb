Public Function CONCATE(ByRef Separador As String, ByRef Área As Range, ByRef Ref As Integer) As String

'Serve para concatenar ranges, sendo:
'CONCATE( "Defina um separador" ; "Defina uma range" ; "Defina a Forma 0, 1, 2 ou 3") 
'Formas:
'0 = Concatenar considerando campos vazios, sem separador nas bordas laterais		1| |2
'1 = Concatenar desprezando campos vazios, sem separadores nas laterais		1|2
'2 = Concatenar considerando campos vazios, com separadores nas laterais		|1||2|
'3 = Concatenar desprezando campos vazios, com separadores nas laterais		|1|2|


Dim xx, sep, virg As String

If Ref = 0 Then
   
    For Each cell In Área.Value
    
        xx = xx & Separador & cell
    
    Next

CONCATE = Replace(xx, Separador, "", 1, 1)

ElseIf Ref = 1 Then

    For Each cell In Área.Value
     If Len(cell) > 0 Then
        xx = xx & Separador & cell
     Else
     End If
    Next

CONCATE = Replace(xx, Separador, "", 1, 1)

ElseIf Ref = 2 Then

    For Each cell In Área.Value
    
        xx = xx & Separador & cell
        
    Next

CONCATE = xx & Separador

ElseIf Ref = 3 Then

    For Each cell In Área.Value
     If Len(cell) > 0 Then
        xx = xx & Separador & cell
     Else
     End If
    Next

CONCATE = xx & Separador

Else
End If


End Function