Public Function BUSCAPHD(DADO As Variant, COLUNA As Integer) As Variant

    'Busca dados no cadastro de empresas do WPHD com base no código de chamada ou CPF/CNPJ
    'Aplicação:
    '=BUSCAPHD(DADO,COLUNA)
    'DADO = Chamada, CNPJ ou CPF
    'COLUNA = Use valores de 1 a 67 correspondentes às colunas da tabela de cadastro de Empresa    
    
    Dim FilePath As String
    Dim FileNum As Integer
    Dim LineData As String
    Dim Columns() As String
    Dim Found As Boolean
    Dim DADO_Clean As String
    Dim ColunaValue As Variant
    Dim tmp As Variant

    On Error GoTo ErrorHandler

    ' Primeiro, limpar o DADO
    DADO_Clean = Trim(CStr(DADO))

    If Len(DADO_Clean) <= 5 Then
        ' Se for até 5 dígitos, transformar em inteiro sem zeros à esquerda
        DADO_Clean = CStr(Val(DADO_Clean))
    Else
        ' Se tiver mais de 5 dígitos, remover ".", "/", "-"
        DADO_Clean = Replace(DADO_Clean, ".", "")
        DADO_Clean = Replace(DADO_Clean, "/", "")
        DADO_Clean = Replace(DADO_Clean, "-", "")
    End If

    ' Agora, abrir o arquivo e começar a busca
    FilePath = "\\192.168.1.251\Arquivos\Controle Contabilidade\BD\Tabelas\Empresas.txt"
    FileNum = FreeFile
    Open FilePath For Input As #FileNum

    Found = False

    Do While Not EOF(FileNum)
        Line Input #FileNum, LineData
        Columns = Split(LineData, ";")
        If Len(DADO_Clean) <= 5 Then
            ' Buscar na Coluna 1
            If UBound(Columns) >= COLUNA - 1 Then
                If Trim(Columns(0)) = DADO_Clean Then
                    ColunaValue = Columns(COLUNA - 1)
                    Found = True
                    Exit Do
                End If
            End If
        Else
            ' Buscar na Coluna 18 após limpar caracteres especiais
            If UBound(Columns) >= Application.WorksheetFunction.Max(17, COLUNA - 1) Then
                tmp = Replace(Columns(17), ".", "")
                tmp = Replace(tmp, "/", "")
                tmp = Replace(tmp, "-", "")
                If Trim(tmp) = DADO_Clean Then
                    ColunaValue = Columns(COLUNA - 1)
                    Found = True
                    Exit Do
                End If
            End If
        End If
    Loop

    Close #FileNum

    If Found Then
        BUSCAPHD = ColunaValue
    Else
        BUSCAPHD = CVErr(xlErrNA)
    End If

    Exit Function

ErrorHandler:
    BUSCAPHD = CVErr(xlErrValue)
End Function
