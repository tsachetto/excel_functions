Public Function ALTERDATA(ByRef Area As Range) As String

'Função elaborada com a finalidade de concatenar dados no layout padrão Alterdata Contábil Windows.
'Sendo baseada em 08 colunas com os valores dispostos respectivamente na seguinte ordem:
'Débito, Crédito, Data, Valor, Histórico Padrão e Complemento Histórico.
'Ex. de Utilização, na Célular I1: =ALTERDATA(A1:H1)
'Ex. de Resultado: "","250","5","01/05/2013","1580,20","","SERVIÇO PRESTADO NF 2252"

Dim sep, virg, txt As String
Dim i As Integer

sep = """"
virg = ","

    For i = 1 To 8
    
        If i = 5 Then
              txt = txt & sep & Format(Round(Area(1, 5), 2), "0.00") & sep & virg
            ElseIf i = 7 Then
              txt = txt & sep & UCase(Area(1, 7)) & sep & virg
            Else
              txt = txt & sep & Area(1, i) & sep & virg
        End If
        
    Next
    
ALTERDATA = StrReverse(Replace(StrReverse(txt), StrReverse(virg), StrReverse(""), , 1))

End Function
