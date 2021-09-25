Public Function SPED2(sped As String, amtI As Integer)
  'Utilize =SPED2("linha de sped", 5)
Dim ini As Integer, fim As Integer, count As Integer
For i = 1 To Len(sped)
    If Mid(sped, i, 1) = "|" Then count = count + 1
    If count = amtI Then
        ini = i
        fim = InStr((i + 1), sped, "|")
        Exit For
    Else: End If
Next
SPED2 = Mid(sped, (ini + 1), (fim - ini - 1))
End Function
