Public Function SPED(spedtx As String, amtI As Integer)
Dim ini As Integer, fim As Integer
For i = 1 To amtI
ini = InStr((ini + 1), spedtx, "|")
fim = InStr((ini + 1), spedtx, "|")
Next
SPED = Mid(spedtx, (ini + 1), (fim - ini - 1))
End Function
