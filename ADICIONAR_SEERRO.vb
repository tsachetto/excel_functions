Sub addseerros()
  
    'Função destinada a adicionar a fórmula "=SEERRO" em todas as células da Active Selection.
  
  Dim se_erro As String
  
  se_erro = """"""    
    
    For Each cell In Range(Selection, Selection)
  
      form = cell.FormulaLocal
      If Len(form) > 0 Then
        formul = "=SEERRO(" & Mid(form, 2, Len(form)) & ";" & se_erro & ")"
        cell.FormulaLocal = formul
      Else: End If
    
    Next
    
End Sub
