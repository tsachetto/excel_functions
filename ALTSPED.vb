Public Function ALTSPED(sped As String, amtI As Integer, ByVal change As String)
'Utilize: =altsped("texto linha do sped |a|b|"; nr campo a ser alterado: 2, novo valor (ex. "x" ou 2)
Dim ini As Integer, fim As Integer, pt1 As String, pt2 As String
For i = 1 To amtI
ini = InStr((ini + 1), sped, "|")
fim = InStr((ini + 1), sped, "|")
Next
pt1 = Mid(sped, 1, ini)
pt2 = Mid(sped, fim, Len(sped))
ALTSPED = pt1 & change & pt2
End Function
