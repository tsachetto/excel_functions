Public Function CPF(ByRef NBR As String) As String

'Converte uma sequência de até 11 caracteres numéricos inteiros para o padrão brasileiro de CPF: 000.000.000-00
'Exemplo: Sequência 12345678912 na célula A1, com a fórmula =CPF(A1) em outra célula teremos: 123.456.789-12

Dim parte1, parte2, parte3, parte4, parte5, cp As String

cp = Right("00000000000000" & NBR, 11)

  parte1 = Left(cp, 3)
  parte2 = Mid(cp, 4, 3)
  parte3 = Mid(cp, 7 , 3)
  parte4 = Right(cp, 2)

CPF = parte1 & "." & parte2 & "." & parte3 & "-" & parte4

End Function
