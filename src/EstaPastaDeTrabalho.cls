VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Controle do carregamento e encerramento
' do suplemento Excel (.xlam).
' ============================================================
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Verifica se o formulário está aberto
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = "frmCorretorEnderecos" Then
            ' Se estiver aberto, manda o formulário atualizar o Label
            frm.AtualizarLabelSelecao
            Exit For
        End If
    Next frm
End Sub
