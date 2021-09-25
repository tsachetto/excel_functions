Public Function SPED(sped As String, amtI As Integer)
Dim ini As Integer, fim As Integer, count As Integer
ini = 0
For i = 1 To amtI
ini = InStr((ini + 1), sped, "|")
fim = InStr((ini + 1), sped, "|")
Next
SPED = Mid(sped, (ini + 1), (fim - ini - 1))
End Function
