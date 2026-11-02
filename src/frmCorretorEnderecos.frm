VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorretorEnderecos 
   Caption         =   "modRegex - Corretor de Endereços - Versão: 0.2.4"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "frmCorretorEnderecos.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCorretorEnderecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Interface principal para correcao de
' enderecos e consulta de CEP.
' ============================================================
Private mRegistros As Collection
Private mRegistroAtual As Variant
Private mTemResultado As Boolean
Private mIndiceResultado As Long
Private mIgnorarAlteracaoBusca As Boolean
Private mInicializando As Boolean
Private mDicaAtiva As Boolean
Private mNumeroDigitado As String
Private Sub UserForm_Initialize()
    mInicializando = True
    Me.Caption = "modRegex - Corretor de Endereços - Versão: 0.2.4"

    With Me.cboPadrao
        .Clear
        .AddItem "Padrão Brasileiro Comum v2.1"
        .AddItem "Apenas Logradouro e Número"
        .AddItem "Endereço Completo (com CEP)"
        .ListIndex = 0
    End With

    Me.chkCapitalizar.Value = True
    Me.chkExpandir.Value = True
    Me.chkValidarCEP.Value = True
    Me.chkVirgulas.Value = False
    Me.optNovaColuna.Value = True

    Me.optConsultarCEP.Value = True
    Me.txtLogradouro.Locked = True
    Me.txtBairro.Locked = True
    Me.txtCidade.Locked = True
    Me.txtUF.Locked = True
    Me.txtCEPResultado.Locked = True

    mInicializando = False
    LimparResultados
    AtualizarModoConsulta
    AtualizarLabelSelecao
End Sub
Public Sub AtualizarLabelSelecao()
    Dim area As Range
    On Error GoTo SemSelecao
    If TypeName(Application.Selection) <> "Range" Then GoTo SemSelecao

    Set area = Application.Selection
    Me.lblSelecao.Caption = "Selecionado: " & area.Worksheet.Name & "!" & _
        area.Address(False, False) & " (" & Format$(area.CountLarge, "#,##0") & " células)"
    Exit Sub
SemSelecao:
    Me.lblSelecao.Caption = "Selecione células em uma planilha."
End Sub
Private Sub btnCorrigir_Click()
    Dim origem As Range, destino As Range, celula As Range, alvo As Range
    Dim antigoEventos As Boolean, antigaTela As Boolean
    Dim configuracaoAlterada As Boolean, falhou As Boolean, erroTexto As String
    Dim total As Long, ignorados As Long, naoReconhecidos As Long
    Dim s As String, novo As String, reconhecido As Boolean
    Dim modo As Long

    On Error GoTo Falha
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Selecione os endereços na planilha.", vbExclamation, "modRegex"
        Exit Sub
    End If
    Set origem = Application.Selection

    If origem.Areas.Count <> 1 Or origem.Columns.Count <> 1 Then
        MsgBox "Selecione um intervalo contínuo de uma coluna.", vbExclamation, "modRegex"
        Exit Sub
    End If
    If origem.CountLarge > 25000 Then
        MsgBox "Selecione até 25.000 células por operação.", vbExclamation, "modRegex"
        Exit Sub
    End If
    If origem.Worksheet.ProtectContents Then
        MsgBox "A planilha está protegida.", vbExclamation, "modRegex"
        Exit Sub
    End If
    If IsNull(origem.MergeCells) Then
        MsgBox "O intervalo pode conter células mescladas.", vbExclamation, "modRegex"
        Exit Sub
    End If
    If origem.MergeCells Then
        MsgBox "O intervalo contém células mescladas.", vbExclamation, "modRegex"
        Exit Sub
    End If

    If Me.optNovaColuna.Value Then
        If origem.Column = origem.Worksheet.Columns.Count Then
            MsgBox "Não existe coluna à direita.", vbExclamation, "modRegex"
            Exit Sub
        End If
        Set destino = origem.Offset(0, 1)
        If IsNull(destino.MergeCells) Then
            MsgBox "O destino pode conter células mescladas.", vbExclamation, "modRegex"
            Exit Sub
        End If
        If destino.MergeCells Then
            MsgBox "O destino contém células mescladas.", vbExclamation, "modRegex"
            Exit Sub
        End If
        If Application.WorksheetFunction.CountA(destino) > 0 Then
            MsgBox "A coluna ao lado contém dados. Escolha outro intervalo.", _
                vbExclamation, "modRegex"
            Exit Sub
        End If
    Else
        If MsgBox("Substituir os endereços selecionados?" & vbCrLf & _
            "O VBA pode limpar o histórico Desfazer do Excel.", _
            vbYesNo + vbQuestion, "modRegex") <> vbYes Then Exit Sub
        Set destino = origem
    End If

    modo = Me.cboPadrao.ListIndex
    antigoEventos = Application.EnableEvents
    antigaTela = Application.ScreenUpdating
    configuracaoAlterada = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each celula In origem.Cells
        If celula.EntireRow.Hidden Then
            ignorados = ignorados + 1
        ElseIf IsError(celula.Value2) Then
            ignorados = ignorados + 1
        ElseIf Len(Trim$(CStr(celula.Value2))) = 0 Then
            ignorados = ignorados + 1
        ElseIf Me.optSubstituir.Value And celula.HasFormula Then
            ignorados = ignorados + 1
        Else
            s = CStr(celula.Value2)
            novo = ProcessarEndereco(s, Me.chkCapitalizar.Value, _
                Me.chkExpandir.Value, Me.chkValidarCEP.Value, Me.chkVirgulas.Value)

            reconhecido = True
            If modo = 1 Then novo = ObterLogradouroNumero(novo, reconhecido)
            If modo = 2 Then reconhecido = TemCEPFormatado(novo)

            If Not reconhecido Or Len(novo) = 0 Then
                naoReconhecidos = naoReconhecidos + 1
            Else
                Set alvo = celula
                If Me.optNovaColuna.Value Then Set alvo = celula.Offset(0, 1)
                alvo.Value2 = novo
                total = total + 1
            End If
        End If
    Next celula

Finalizar:
    On Error Resume Next
    If configuracaoAlterada Then
        Application.EnableEvents = antigoEventos
        Application.ScreenUpdating = antigaTela
    End If
    On Error GoTo 0

    If falhou Then
        MsgBox "Falha: " & erroTexto & vbCrLf & _
            "Confira os dados; parte das células pode ter sido alterada.", _
            vbExclamation, "modRegex"
    Else
        MsgBox total & " endereço(s) processado(s)." & vbCrLf & _
            ignorados & " ignorado(s)." & vbCrLf & _
            naoReconhecidos & " não reconhecido(s) pelo padrão escolhido.", _
            vbInformation, "modRegex"
    End If
    AtualizarLabelSelecao
    Exit Sub
Falha:
    falhou = True
    erroTexto = Err.Description
    Resume Finalizar
End Sub
Private Sub btnCancelar_Click()
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    DesativarMonitoramento
End Sub
Private Sub optConsultarCEP_Click()
    If mInicializando Then Exit Sub
    AtualizarModoConsulta
End Sub
Private Sub optPesquisarEndereco_Click()
    If mInicializando Then Exit Sub
    AtualizarModoConsulta
End Sub
' Mantem o mesmo campo para consultar por CEP e pesquisar por endereco.
' O placeholder e apenas uma dica visual, nunca um valor para enviar a API.
Private Sub AtualizarModoConsulta()
    If mInicializando Then Exit Sub

    Me.lblCEPConsulta.Visible = True
    Me.txtConsultaCEP.Visible = True
    Me.txtConsultaCEP.Enabled = True

    If Me.optPesquisarEndereco.Value Then
        Me.lblCEPConsulta.Caption = "End.:"
        Me.txtConsultaCEP.ControlTipText = _
            "Ex.: Av. Paulista, 1000 - Sao Paulo/SP (numero opcional)"
        Me.btnConsultarCEP.Caption = "Pesquisar"
    Else
        Me.lblCEPConsulta.Caption = "CEP:"
        Me.txtConsultaCEP.ControlTipText = "Digite 8 numeros. O hifen sera inserido automaticamente."
        Me.btnConsultarCEP.Caption = "Consultar"
    End If

    mNumeroDigitado = vbNullString
    LimparResultados
    MostrarDica
End Sub
Private Function TextoDica() As String
    If Me.optPesquisarEndereco.Value Then
        TextoDica = "Rua, nº - Cidade/UF"
    Else
        TextoDica = "00000-000"
    End If
End Function
Private Sub MostrarDica()
    mIgnorarAlteracaoBusca = True
    mDicaAtiva = True
    Me.txtConsultaCEP.ForeColor = RGB(145, 145, 145)
    Me.txtConsultaCEP.Value = TextoDica()
    mIgnorarAlteracaoBusca = False
End Sub
Private Sub RemoverDica()
    If Not mDicaAtiva Then Exit Sub
    mIgnorarAlteracaoBusca = True
    Me.txtConsultaCEP.Value = vbNullString
    Me.txtConsultaCEP.ForeColor = RGB(35, 35, 35)
    mDicaAtiva = False
    mIgnorarAlteracaoBusca = False
End Sub
Private Sub txtConsultaCEP_Enter()
    RemoverDica
End Sub
Private Sub txtConsultaCEP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim$(Me.txtConsultaCEP.Value)) = 0 Then MostrarDica
End Sub
' No modo CEP, aplica a mascara 00000-000 enquanto o usuario digita.
' Tambem trata texto colado e preserva a posicao do cursor.
' No modo endereco, o texto permanece livre.
Private Sub txtConsultaCEP_Change()
    Dim entrada As String
    Dim numeros As String
    Dim formatado As String
    Dim digitosAntesCursor As Long
    Dim novaPosicao As Long

    If mIgnorarAlteracaoBusca Then Exit Sub

    If Me.optConsultarCEP.Value And Not mDicaAtiva Then
        entrada = CStr(Me.txtConsultaCEP.Value)
        digitosAntesCursor = ContarDigitosCEP(Left$(entrada, Me.txtConsultaCEP.SelStart))
        numeros = ExtrairDigitosCEP(entrada)

        If Len(numeros) > 5 Then
            formatado = Left$(numeros, 5) & "-" & Mid$(numeros, 6)
        Else
            formatado = numeros
        End If

        If digitosAntesCursor > Len(numeros) Then
            digitosAntesCursor = Len(numeros)
        End If

        novaPosicao = digitosAntesCursor
        If digitosAntesCursor > 5 Then novaPosicao = novaPosicao + 1

        If entrada <> formatado Then
            mIgnorarAlteracaoBusca = True
            Me.txtConsultaCEP.Value = formatado
            Me.txtConsultaCEP.SelStart = novaPosicao
            mIgnorarAlteracaoBusca = False
        End If
    End If

    ' Qualquer edicao invalida resultados de uma pesquisa anterior.
    mNumeroDigitado = vbNullString
    LimparResultados
End Sub
Private Function ExtrairDigitosCEP(ByVal texto As String) As String
    Dim i As Long
    Dim caractere As String
    Dim digitos As String

    For i = 1 To Len(texto)
        caractere = Mid$(texto, i, 1)
        If caractere >= "0" And caractere <= "9" Then
            digitos = digitos & caractere
            If Len(digitos) = 8 Then Exit For
        End If
    Next i

    ExtrairDigitosCEP = digitos
End Function

Private Function ContarDigitosCEP(ByVal texto As String) As Long
    Dim i As Long
    Dim caractere As String

    For i = 1 To Len(texto)
        caractere = Mid$(texto, i, 1)
        If caractere >= "0" And caractere <= "9" Then
            ContarDigitosCEP = ContarDigitosCEP + 1
        End If
    Next i
End Function
' Aceita o formato simples "Av. Paulista, 1000 - Sao Paulo/SP".
' O numero nao participa da consulta do ViaCEP, apenas da copia/insercao.
' Tambem aceita o formato antigo "SP; Sao Paulo; Avenida Paulista".
Private Function LerDadosPesquisa(ByVal entrada As String, _
    ByRef uf As String, ByRef cidade As String, _
    ByRef logradouro As String, ByRef numero As String) As Boolean

    Dim partes() As String
    Dim barra As Long, separador As Long, virgula As Long
    Dim anteriorUF As String, enderecoRua As String, candidatoNumero As String

    entrada = Trim$(entrada)
    numero = vbNullString

    If InStr(1, entrada, ";", vbBinaryCompare) > 0 Then
        partes = Split(entrada, ";")
        If UBound(partes) <> 2 Then GoTo FormatoInvalido
        uf = UCase$(Trim$(partes(0)))
        cidade = Trim$(partes(1))
        enderecoRua = Trim$(partes(2))
    Else
        barra = InStrRev(entrada, "/")
        If barra < 2 Then GoTo FormatoInvalido

        uf = UCase$(Trim$(Mid$(entrada, barra + 1)))
        anteriorUF = Trim$(Left$(entrada, barra - 1))
        separador = InStrRev(anteriorUF, " - ")
        If separador < 2 Then GoTo FormatoInvalido

        enderecoRua = Trim$(Left$(anteriorUF, separador - 1))
        cidade = Trim$(Mid$(anteriorUF, separador + 3))
    End If

    ' Remove do logradouro apenas um numero explicito apos virgula.
    ' Nomes como "Rua 25 de Marco" permanecem intactos.
    virgula = InStrRev(enderecoRua, ",")
    If virgula > 1 Then
        candidatoNumero = Trim$(Mid$(enderecoRua, virgula + 1))
        If EhNumeroDeImovel(candidatoNumero) Then
            numero = candidatoNumero
            enderecoRua = Trim$(Left$(enderecoRua, virgula - 1))
        End If
    End If

    logradouro = enderecoRua

    If Len(uf) <> 2 Or uf Like "*[!A-Z]*" Then GoTo DadosInvalidos
    If Len(cidade) < 3 Or Len(logradouro) < 3 Then GoTo DadosInvalidos

    LerDadosPesquisa = True
    Exit Function

FormatoInvalido:
    MsgBox "Informe: Rua, numero - Cidade/UF" & vbCrLf & _
           "Ex.: Av. Paulista, 1000 - Sao Paulo/SP" & vbCrLf & _
           "O numero e opcional.", vbInformation, "modRegex"
    Exit Function

DadosInvalidos:
    MsgBox "Confira a UF (2 letras), a cidade e o logradouro " & _
           "(minimo de 3 caracteres).", vbExclamation, "modRegex"
End Function
Private Function EhNumeroDeImovel(ByVal valor As String) As Boolean
    Dim regex As Object
    On Error GoTo Fim
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(n[.oº°]?\s*)?[0-9]+[A-Za-z]?$"
        .IgnoreCase = True
        .Global = False
    End With
    EhNumeroDeImovel = regex.Test(valor)
Fim:
End Function
Private Sub btnConsultarCEP_Click()
    Dim msg As String, uf As String, cidade As String, logradouro As String
    Dim pesquisaEndereco As Boolean, numero As String

    On Error GoTo Falha
    pesquisaEndereco = Me.optPesquisarEndereco.Value

    ' Quando a busca retorna varios enderecos, o mesmo botao avanca para o proximo.
    If pesquisaEndereco Then
        If Not mRegistros Is Nothing Then
            If mRegistros.Count > 1 Then
                mIndiceResultado = mIndiceResultado + 1
                If mIndiceResultado > mRegistros.Count Then mIndiceResultado = 1
                SelecionarRegistro mIndiceResultado
                AtualizarBotaoPesquisa
                Exit Sub
            End If
        End If
    End If

    LimparResultados
    If mDicaAtiva Or Len(Trim$(Me.txtConsultaCEP.Value)) = 0 Then
        MsgBox "Preencha o campo para consultar.", vbInformation, "modRegex"
        Exit Sub
    End If

    If pesquisaEndereco Then
        If Not LerDadosPesquisa(Me.txtConsultaCEP.Value, uf, cidade, logradouro, numero) Then Exit Sub
        mNumeroDigitado = numero
    End If

    Me.btnConsultarCEP.Enabled = False
    Me.MousePointer = fmMousePointerHourGlass
    Me.Repaint

    If pesquisaEndereco Then
        Set mRegistros = modViaCEP.PesquisarCEPs(uf, cidade, logradouro, msg)
    Else
        Set mRegistros = modViaCEP.ConsultarCEP(Me.txtConsultaCEP.Value, msg)
    End If

    If Not mRegistros Is Nothing Then
        If mRegistros.Count > 0 Then
            mIndiceResultado = 1
            SelecionarRegistro mIndiceResultado
        End If
    End If

    If mIndiceResultado = 0 Then
        If Len(msg) = 0 Then msg = "Nenhum endereco encontrado."
        MsgBox msg, vbInformation, "modRegex"
    End If

FinalizarConsulta:
    Me.MousePointer = fmMousePointerDefault
    Me.btnConsultarCEP.Enabled = True
    AtualizarBotaoPesquisa
    Exit Sub

Falha:
    MsgBox "Falha na consulta: " & Err.Description, vbExclamation, "modRegex"
    Resume FinalizarConsulta
End Sub
Private Sub AtualizarBotaoPesquisa()
    Dim proximo As Long

    If Me.optConsultarCEP.Value Then
        Me.btnConsultarCEP.Caption = "Consultar"
    ElseIf mRegistros Is Nothing Then
        Me.btnConsultarCEP.Caption = "Pesquisar"
    ElseIf mRegistros.Count <= 1 Then
        Me.btnConsultarCEP.Caption = "Pesquisar"
    Else
        proximo = mIndiceResultado + 1
        If proximo > mRegistros.Count Then proximo = 1
        Me.btnConsultarCEP.Caption = "Próximo " & proximo & "/" & mRegistros.Count
    End If
End Sub
Private Sub SelecionarRegistro(ByVal indice As Long)
    If mRegistros Is Nothing Then Exit Sub
    If indice < 1 Or indice > mRegistros.Count Then Exit Sub

    mRegistroAtual = mRegistros(indice)
    mTemResultado = True
    Me.txtCEPResultado.Value = CStr(mRegistroAtual(0))
    Me.txtLogradouro.Value = CStr(mRegistroAtual(1))
    Me.txtBairro.Value = CStr(mRegistroAtual(2))
    Me.txtCidade.Value = CStr(mRegistroAtual(3))
    Me.txtUF.Value = CStr(mRegistroAtual(4))
    Me.btnCopiar.Enabled = True
    Me.btnInserirExcel.Enabled = True
End Sub
Private Sub LimparResultados()
    mTemResultado = False
    Set mRegistros = Nothing
    mIndiceResultado = 0
    Me.txtLogradouro.Value = vbNullString
    Me.txtBairro.Value = vbNullString
    Me.txtCidade.Value = vbNullString
    Me.txtUF.Value = vbNullString
    Me.txtCEPResultado.Value = vbNullString
    Me.btnCopiar.Enabled = False
    Me.btnInserirExcel.Enabled = False
    Me.btnCopiar.Caption = "Copiar"
    Me.btnInserirExcel.Caption = "Inserir no Excel"
    AtualizarBotaoPesquisa
End Sub
Private Function EnderecoParaTexto() As String
    Dim partes As String, cidadeUF As String
    If Not mTemResultado Then Exit Function

    partes = CStr(mRegistroAtual(1))
    ' Numero fornecido pelo usuario; nao foi validado pelo ViaCEP.
    If Me.optPesquisarEndereco.Value And Len(mNumeroDigitado) > 0 Then
        If Len(partes) > 0 Then partes = partes & ", " & mNumeroDigitado
    End If
    If Len(CStr(mRegistroAtual(5))) > 0 Then
        If Len(partes) > 0 Then partes = partes & ", "
        partes = partes & CStr(mRegistroAtual(5))
    End If
    If Len(CStr(mRegistroAtual(2))) > 0 Then
        If Len(partes) > 0 Then partes = partes & " - "
        partes = partes & CStr(mRegistroAtual(2))
    End If

    cidadeUF = CStr(mRegistroAtual(3))
    If Len(CStr(mRegistroAtual(4))) > 0 Then
        If Len(cidadeUF) > 0 Then cidadeUF = cidadeUF & "/"
        cidadeUF = cidadeUF & CStr(mRegistroAtual(4))
    End If
    If Len(cidadeUF) > 0 Then
        If Len(partes) > 0 Then partes = partes & " - "
        partes = partes & cidadeUF
    End If
    If Len(CStr(mRegistroAtual(0))) > 0 Then
        If Len(partes) > 0 Then partes = partes & " - "
        partes = partes & CStr(mRegistroAtual(0))
    End If
    EnderecoParaTexto = partes
End Function
Private Sub btnCopiar_Click()
    Dim dados As MSForms.DataObject
    On Error GoTo Falha
    If Not mTemResultado Then Exit Sub
    Set dados = New MSForms.DataObject
    dados.SetText EnderecoParaTexto()
    dados.PutInClipboard
    Me.btnCopiar.Caption = "Copiado"
    Exit Sub
Falha:
    MsgBox "Não foi possível copiar: " & Err.Description, vbExclamation, "modRegex"
End Sub
Private Sub btnInserirExcel_Click()
    Dim destino As Range, conteudo As String
    On Error GoTo Falha
    If Not mTemResultado Then Exit Sub
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Selecione uma célula do Excel.", vbExclamation, "modRegex"
        Exit Sub
    End If

    Set destino = Application.Selection
    If destino.CountLarge <> 1 Then
        MsgBox "Selecione apenas uma célula de destino.", vbExclamation, "modRegex"
        Exit Sub
    End If
    If destino.Worksheet.ProtectContents Or destino.MergeCells Then
        MsgBox "O destino está protegido ou mesclado.", vbExclamation, "modRegex"
        Exit Sub
    End If

    conteudo = EnderecoParaTexto()
    If Not IsEmpty(destino.Value2) Or destino.HasFormula Then
        If MsgBox("A célula já contém dados. Substituir?", _
            vbYesNo + vbQuestion, "modRegex") <> vbYes Then Exit Sub
    End If

    destino.Value2 = conteudo
    Me.btnInserirExcel.Caption = "Inserido"
    Exit Sub
Falha:
    MsgBox "Não foi possível inserir: " & Err.Description, vbExclamation, "modRegex"
End Sub
