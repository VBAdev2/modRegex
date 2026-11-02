Attribute VB_Name = "modViaCEP"
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Integracao com a API ViaCEP para consulta
' de enderecos a partir do CEP informado.
'
' API utilizada: https://viacep.com.br/
' ============================================================
Public Function ConsultarCEP(ByVal entrada As String, _
    ByRef mensagem As String) As Collection

    Dim cep As String, doc As Object, registros As Collection
    Dim no As Object
    cep = SomenteDigitos(entrada)

    If Len(cep) <> 8 Then
        mensagem = "Informe um CEP com 8 digitos."
        Exit Function
    End If

    Set doc = ObterXML("https://viacep.com.br/ws/" & cep & "/xml/", mensagem)
    If doc Is Nothing Then Exit Function

    Set registros = New Collection
    Set no = doc.SelectSingleNode("/xmlcep/cep")
    If Not no Is Nothing Then
        registros.Add LerRegistro(doc.DocumentElement)
    Else
        mensagem = "CEP nao encontrado."
    End If
    Set ConsultarCEP = registros
End Function
Public Function PesquisarCEPs(ByVal uf As String, ByVal cidade As String, _
    ByVal logradouro As String, ByRef mensagem As String) As Collection

    Dim doc As Object, registros As Collection
    Dim nos As Object, no As Object, endereco As String

    uf = UCase$(Trim$(uf))
    cidade = Trim$(cidade)
    logradouro = Trim$(logradouro)

    If Len(uf) <> 2 Or uf Like "*[!A-Z]*" Then
        mensagem = "Selecione uma UF valida."
        Exit Function
    End If
    If Len(cidade) < 3 Or Len(logradouro) < 3 Then
        mensagem = "Cidade e logradouro precisam ter 3 caracteres ou mais."
        Exit Function
    End If

    endereco = "https://viacep.com.br/ws/" & uf & "/" & _
               Application.WorksheetFunction.EncodeURL(cidade) & "/" & _
               Application.WorksheetFunction.EncodeURL(logradouro) & "/xml/"

    Set doc = ObterXML(endereco, mensagem)
    If doc Is Nothing Then Exit Function

    Set registros = New Collection
    Set nos = doc.SelectNodes("/xmlcep/enderecos/endereco | /xmlcep/endereco")
    For Each no In nos
        If Len(ValorNo(no, "cep")) > 0 Then registros.Add LerRegistro(no)
    Next no
    Set PesquisarCEPs = registros

    If registros.Count = 0 Then mensagem = "Nenhum endereco encontrado."
End Function
Private Function ObterXML(ByVal url As String, ByRef mensagem As String) As Object
    Dim http As Object, xml As Object, erro As Object
    On Error GoTo falhou

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 4000, 4000, 8000, 8000
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/xml"
    http.Send

    If http.Status <> 200 Then
        Select Case http.Status
            Case 400: mensagem = "Consulta invalida. Confira os dados."
            Case 404: mensagem = "Nenhum endereco encontrado."
            Case Else: mensagem = "ViaCEP indisponivel (HTTP " & http.Status & ")."
        End Select
        Exit Function
    End If

    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.validateOnParse = False
    xml.resolveExternals = False

    If Not xml.Load(http.responseBody) Then
        mensagem = "Nao foi possivel interpretar a resposta do ViaCEP."
        Exit Function
    End If

    Set erro = xml.SelectSingleNode("/xmlcep/erro")
    If Not erro Is Nothing Then
        mensagem = "CEP ou endereco nao encontrado."
        Exit Function
    End If

    Set ObterXML = xml
    Exit Function
falhou:
    mensagem = "Falha na consulta: " & Err.Description
End Function
Private Function LerRegistro(ByVal no As Object) As Variant
    LerRegistro = Array( _
        ValorNo(no, "cep"), ValorNo(no, "logradouro"), _
        ValorNo(no, "bairro"), ValorNo(no, "localidade"), _
        ValorNo(no, "uf"), ValorNo(no, "complemento"))
End Function
Private Function ValorNo(ByVal no As Object, ByVal nome As String) As String
    Dim item As Object
    Set item = no.SelectSingleNode(nome)
    If Not item Is Nothing Then ValorNo = Trim$(item.Text)
End Function
Private Function SomenteDigitos(ByVal s As String) As String
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then SomenteDigitos = SomenteDigitos & c
    Next i
End Function
