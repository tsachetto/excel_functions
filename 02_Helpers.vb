'Remover eventual filtro
If ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
End If

