Function PROCSIMILAR(inputCompany As String, companyList As Range) As String
  'Esta função procura valores similares
  'Como utilizar: =PROCSMILIAR(valor procurado, lista com valores procurados)
   
  
    Dim cell As Range
    Dim minDistance As Integer
    Dim currentDistance As Integer
    Dim minCompany As String
    
    minDistance = LevenshteinDistance(inputCompany, companyList.Cells(1, 1).Value)
    minCompany = companyList.Cells(1, 1).Value
    
    For Each cell In companyList
        currentDistance = LevenshteinDistance(inputCompany, cell.Value)
        If currentDistance < minDistance Then
            minDistance = currentDistance
            minCompany = cell.Value
        End If
    Next
    If inputCompany = "" Then
    PROCSIMILAR = ""
    Else
    PROCSIMILAR = minCompany
    End If
    
End Function

Function LevenshteinDistance(s1 As String, s2 As String) As Integer
'Esta função é parte da anterior e não deve ser aplicada na planilha.
'Trata-se de um algoritmo Levenshtein Distance

    Dim i As Integer
    Dim j As Integer
    Dim dist() As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    Dim cost As Integer

    len1 = Len(s1)
    len2 = Len(s2)
    ReDim dist(len1, len2)

    For i = 0 To len1
        dist(i, 0) = i
    Next

    For j = 0 To len2
        dist(0, j) = j
    Next

    For i = 1 To len1
        For j = 1 To len2
            cost = Abs(StrComp(Mid$(s1, i, 1), Mid$(s2, j, 1), vbTextCompare) <> 0)
            dist(i, j) = Application.WorksheetFunction.Min(dist(i - 1, j) + 1, _
                                                            dist(i, j - 1) + 1, _
                                                            dist(i - 1, j - 1) + cost)
        Next
    Next

    LevenshteinDistance = dist(len1, len2)
End Function
