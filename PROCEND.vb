Public Function PROCEND(ByVal strng As Long) As String

'A função irá retornar o endereço da célula de uma string ou número procurada mais próxiuma de seu resultado.

  PROCEND = Cells.Find(What:=strng, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Address
               
End Function
