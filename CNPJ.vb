Public Function CNPJ(ByRef NBR As String) As String

'Converte uma sequência de até 14 caracteres numéricos inteiros para o padrão brasileiro de CNPJ: 00.000.000/0001-00
'Exemplo: Sequência 12123456000123 na célula A1, com a fórmula =CNPJ(A1) em outra célula teremos: 12.123.456/0001-23

Dim parte1, parte2, parte3, parte4, parte5, cnp As String

cnp = Right("00000000000000" & NBR, 14)

parte1 = Left(cnp, 2)
parte2 = Mid(cnp, 3, 3)
parte3 = Mid(cnp, 6, 3)
parte4 = Mid(cnp, 9, 4)
parte5 = Right(cnp, 2)

CNPJ = parte1 & "." & parte2 & "." & parte3 & "/" & parte4 & "-" & parte5

End Function
