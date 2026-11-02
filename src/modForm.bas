Attribute VB_Name = "modForm"
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Abertura e encerramento do formulario,
' gerenciamento dos eventos e inicializacao do suplemento.
' ============================================================
Private mEventos As cAppEventos
' Macro utilizada pelo botao da Faixa de Opcoes.
Public Sub AbrirFormulario()

    Dim frm As Object

    On Error GoTo TratarErro

    ' Inicia o acompanhamento das selecoes.
    IniciarMonitoramento

    ' Evita abrir duas instancias do formulario.
    For Each frm In VBA.UserForms

        If TypeName(frm) = "frmCorretorEnderecos" Then

            frm.Show vbModeless
            frm.AtualizarLabelSelecao
            Exit Sub

        End If

    Next frm

    ' Abre somente quando solicitado pelo usuario.
    frmCorretorEnderecos.Show vbModeless
    frmCorretorEnderecos.AtualizarLabelSelecao

    If Not Application.EnableEvents Then
        MsgBox "Os eventos do Excel estao desativados." & _
               vbCrLf & "A selecao nao sera atualizada " & _
               "automaticamente ate que sejam reativados.", _
               vbExclamation, "modRegex"
    End If

    Exit Sub

TratarErro:

    MsgBox "Nao foi possivel abrir o corretor: " & _
           Err.Description, vbExclamation, "modRegex"

End Sub
' Mantem o objeto de eventos ativo enquanto o
' formulario estiver em utilizacao.
Public Sub IniciarMonitoramento()

    If mEventos Is Nothing Then
        Set mEventos = New cAppEventos
    End If

    Set mEventos.ExcelApp = Application

End Sub
' Libera os eventos da aplicacao.
Public Sub DesativarMonitoramento()

    On Error Resume Next

    If Not mEventos Is Nothing Then
        Set mEventos.ExcelApp = Nothing
    End If

    Set mEventos = Nothing

    On Error GoTo 0

End Sub
' Encerra o formulario e o monitoramento.
Public Sub EncerrarCorretor()

    Dim frm As Object

    On Error Resume Next

    For Each frm In VBA.UserForms

        If TypeName(frm) = "frmCorretorEnderecos" Then
            Unload frm
            Exit For
        End If

    Next frm

    DesativarMonitoramento

    On Error GoTo 0

End Sub
' Configuracao executada somente quando necessario.
Public Sub PrepararSuplemento()

    On Error GoTo TratarErro

    ThisWorkbook.IsAddin = True
    ThisWorkbook.Save

    MsgBox "Suplemento configurado com sucesso.", _
           vbInformation, "modRegex"

    Exit Sub

TratarErro:

    MsgBox "Erro ao preparar suplemento: " & _
           Err.Description, vbExclamation, "modRegex"

End Sub


