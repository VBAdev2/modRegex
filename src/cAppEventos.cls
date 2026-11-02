VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Monitoramento dos eventos do Excel para
' atualizar a selecao de celulas no formulario.
' ============================================================
Public WithEvents ExcelApp As Excel.Application
Attribute ExcelApp.VB_VarHelpID = -1
' Executado quando o usuario seleciona celulas.
Private Sub ExcelApp_SheetSelectionChange( _
    ByVal Sh As Object, _
    ByVal Target As Excel.Range)

    AtualizarFormulario

End Sub
' Executado quando uma planilha e ativada.
Private Sub ExcelApp_SheetActivate( _
    ByVal Sh As Object)

    AtualizarFormulario

End Sub
' Executado quando uma pasta de trabalho e ativada.
Private Sub ExcelApp_WorkbookActivate( _
    ByVal Wb As Excel.Workbook)

    AtualizarFormulario

End Sub
' Executado quando outra janela do Excel e ativada.
Private Sub ExcelApp_WindowActivate( _
    ByVal Wb As Excel.Workbook, _
    ByVal Wn As Excel.Window)

    AtualizarFormulario

End Sub
' Atualiza o Label do formulario existente.
Private Sub AtualizarFormulario()

    Dim frm As Object

    On Error GoTo Finalizar

    For Each frm In VBA.UserForms

        If TypeName(frm) = "frmCorretorEnderecos" Then

            If frm.Visible Then
                frm.AtualizarLabelSelecao
            End If

            Exit For

        End If

    Next frm

Finalizar:

End Sub
