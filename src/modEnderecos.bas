Attribute VB_Name = "modEnderecos"
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Normalizacao de enderecos, capitalizacao,
' abreviacoes, pontuacao e formatacao de CEP.
' ============================================================
Public Function ProcessarEndereco(ByVal texto As String, _
    ByVal capitalizar As Boolean, ByVal expandirAbrev As Boolean, _
    ByVal formatarCEP As Boolean, ByVal corrigirPontuacao As Boolean) As String

    Dim s As String
    s = CStr(ExtrairTrecho(texto))
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    s = Trim$(s)

    If expandirAbrev Then s = ExpandirAbreviaturas(s)
    If capitalizar Then s = CapitalizarEndereco(s)
    If formatarCEP Then s = FormatarCEPNoTexto(s)
    If corrigirPontuacao Then s = AjustarPontuacao(s)

    ProcessarEndereco = Trim$(s)
End Function
Private Function ExpandirAbreviaturas(ByVal s As String) As String
    s = TrocarAbrev(s, "R\.?", "Rua")
    s = TrocarAbrev(s, "Av\.?", "Avenida")
    s = TrocarAbrev(s, "Al\.?", "Alameda")
    s = TrocarAbrev(s, "P[cç]a?\.?", "Praça")
    s = TrocarAbrev(s, "Trav\.?", "Travessa")
    s = TrocarAbrev(s, "Rod\.?", "Rodovia")
    ExpandirAbreviaturas = s
End Function
Private Function TrocarAbrev(ByVal s As String, ByVal abreviacao As String, _
    ByVal nomeCompleto As String) As String

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(^|[ ,;])" & abreviacao & "[ ]+"
    TrocarAbrev = re.Replace(s, "$1" & nomeCompleto & " ")
End Function
Private Function CapitalizarEndereco(ByVal s As String) As String
    Dim partes() As String, i As Long, palavra As String, limpa As String
    Dim conectores As String

    conectores = "|a|as|o|os|de|da|do|das|dos|e|em|no|na|nos|nas|"
    s = StrConv(s, vbProperCase)
    partes = Split(s, " ")

    For i = LBound(partes) To UBound(partes)
        palavra = partes(i)
        limpa = LCase$(Replace(Replace(Replace(palavra, ",", ""), ".", ""), ";", ""))
        If i > 0 And InStr(1, conectores, "|" & limpa & "|", vbBinaryCompare) > 0 Then
            partes(i) = LCase$(palavra)
        End If
    Next i

    CapitalizarEndereco = Join(partes, " ")
End Function
Private Function FormatarCEPNoTexto(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(^|[^0-9])([0-9]{5})[- ]?([0-9]{3})(?=$|[^0-9])"
    FormatarCEPNoTexto = re.Replace(s, "$1$2-$3")
End Function
Private Function AjustarPontuacao(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[ ]*,[ ]*"
    s = re.Replace(s, ", ")
    Do While InStr(s, ", ,") > 0
        s = Replace(s, ", ,", ",")
    Loop
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    AjustarPontuacao = Trim$(s)
End Function
' O padrao 2 so extrai quando encontra uma virgula seguida de numero.
Public Function ObterLogradouroNumero(ByVal s As String, _
    ByRef reconhecido As Boolean) As String

    Dim re As Object, resultado As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(.+?),[ ]*([0-9]{1,6}[A-Za-z]?|S[./ ]?N)(?=$|[ ,;])"
    Set resultado = re.Execute(Trim$(s))

    reconhecido = (resultado.Count > 0)
    If reconhecido Then
        ObterLogradouroNumero = Trim$(resultado(0).SubMatches(0)) & ", " & _
                                Trim$(resultado(0).SubMatches(1))
    Else
        ObterLogradouroNumero = s
    End If
End Function
' Verifica apenas o FORMATO, nao a existencia do CEP.
Public Function TemCEPFormatado(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(^|[^0-9])[0-9]{5}-[0-9]{3}(?=$|[^0-9])"
    TemCEPFormatado = re.Test(s)
End Function
