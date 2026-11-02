Attribute VB_Name = "modRegex"
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2
'
' Descricao: Identificacao e extracao de componentes de
' enderecos utilizando expressoes regulares (Regex).
' ============================================================
Public Function ExtrairTrecho( _
        ByVal txt As Variant, _
        Optional ByVal padroes As Variant, _
        Optional ByVal fallbackOriginal As Boolean = True, _
        Optional ByVal limparResultado As Boolean = True) As Variant

    On Error GoTo Falha

    ' -------- validação --------
    If IsError(txt) Then ExtrairTrecho = txt: Exit Function
    If IsNull(txt) Or IsEmpty(txt) Then ExtrairTrecho = vbNullString: Exit Function
    If IsArray(txt) Then ExtrairTrecho = CVErr(xlErrValue): Exit Function

    Dim s As String: s = CStr(txt)
    If Len(s) = 0 Then ExtrairTrecho = s: Exit Function
    If Len(s) > 32767 Then s = Left$(s, 32767)

    ' normaliza quebras/abas -> espaço (regex não cruza \n)
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    s = Trim$(s)

    ' -------- padrões (ordem = prioridade) --------
    Dim arr As Variant
    If IsMissing(padroes) Or IsEmpty(padroes) Then
        arr = Array( _
            "Pr[óo]ximo a \(precis[ãa]o \d+ metros\)\s*(.*?),\s*\b(Brazil|Brasil|BR|BRA)\b\.?", _
            "Pr[óo]ximo a \(precis[ãa]o \d+ metros\)\s*(.+)$", _
            "\|\s*(.+)$" _
        )
    ElseIf IsArray(padroes) Then
        arr = padroes
    Else
        arr = Array(CStr(padroes))
    End If

    ' -------- motor regex (cache entre chamadas) --------
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True      ' todas as ocorrências -> escolhe 1ª com captura válida
        re.IgnoreCase = True  ' Próximo/próximo, Brasil/BRASIL
        re.Multiline = False
    End If

    Dim i As Long, m As Object, col As Object, cap As String
    For i = LBound(arr) To UBound(arr)
        re.Pattern = CStr(arr(i))
        Set col = re.Execute(s)          ' 1 passada (Test+Execute unificados)
        If col.Count > 0 Then
            For Each m In col
                If m.SubMatches.Count > 0 Then
                    cap = CStr(m.SubMatches(0))
                    If Len(Trim$(cap)) > 0 Then
                        If limparResultado Then cap = LimparEndereco(cap)
                        ExtrairTrecho = cap
                        Exit Function
                    End If
                End If
            Next m
        End If
    Next i

    ' -------- fallback --------
    If limparResultado Then s = LimparEndereco(s)
    ExtrairTrecho = IIf(fallbackOriginal, s, vbNullString)
    Exit Function

Falha:
    ExtrairTrecho = txt
End Function
' Remove espaços/pontuação redundantes e normaliza o resultado
Private Function LimparEndereco(ByVal s As String) As String
    Dim c As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    s = Trim$(s)

    Do While Len(s) > 0
        c = Left$(s, 1)
        If c = "," Or c = ";" Or c = "-" Or c = "|" Or c = " " Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(s) > 0
        c = Right$(s, 1)
        If c = "," Or c = ";" Or c = "-" Or c = "|" Or c = " " Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    LimparEndereco = Trim$(s)
End Function

