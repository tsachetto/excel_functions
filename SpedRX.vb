Public Function SPEDRX(text As String, pattern As String, text_replace As String, Optional instance_num As Integer = 0, Optional match_case As Boolean = True) As String
    Dim text_result, text_find As String
    Dim matches_index, pos_start As Integer
    On Error GoTo ErrHandl
    
    text_result = text
    instance_num = 1
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    RegEx.pattern = pattern
    RegEx.Global = True
    RegEx.MultiLine = True
    
    If True = match_case Then
        RegEx.IgnoreCase = False
    Else
        RegEx.IgnoreCase = True
    End If
    
    Set matches = RegEx.Execute(text)
        
    If 0 < matches.Count Then
        If (0 = instance_num) Then
            text_result = RegEx.Replace(text, text_replace)
        Else
            If instance_num <= matches.Count Then
                pos_start = 1
                For matches_index = 0 To instance_num - 2
                    pos_start = InStr(pos_start, text, matches.Item(matches_index), vbBinaryCompare) + Len(matches.Item(matches_index))
                Next matches_index
                
                text_find = matches.Item(instance_num - 1)
                text_result = Left(text, pos_start - 1) & Replace(text, text_find, text_replace, pos_start, 1, vbBinaryCompare)
            End If
        End If
    End If
    
    SPEDRX = text_result
    Exit Function

ErrHandl:
    SPEDRX = CVErr(xlErrValue)
End Function
