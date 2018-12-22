Public Function EXTNR(ByRef NBR As String) As Long

Dim NBROK, nblen As Long
nblen = Len(NBR)

    For i = 1 To nblen
    
    x = Mid(NBR, i, 1)
    
        If IsNumeric(x) = True Then
            NBROK = NBROK & Mid(NBR, i, 1)
        ElseIf NBROK > 0 Then
            EXTNR = NBROK
        Exit Function
        Else
        End If
    
    Next

EXTNR = NBROK

End Function
