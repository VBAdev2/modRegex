Attribute VB_Name = "modTestes"
Option Explicit
' ============================================================
' modRegex - Corretor de Enderecos para Microsoft Excel
' Desenvolvimento: VBAdev2.
'
' Descricao: Testes de validacao das rotinas de
' normalizacao e tratamento de enderecos.
' ============================================================
' Testes locais: nao alteram planilhas nem consultam a internet.
Public Sub TestarModRegex()
    Dim falhas As Long, ok As Boolean, s As String

    s = ProcessarEndereco("R. DAS FLORES, 123", True, True, True, True)
    If s <> "Rua das Flores, 123" Then
        Debug.Print "FALHOU - capitalizacao: "; s
        falhas = falhas + 1
    End If

    s = ProcessarEndereco("Av. Paulista, 1000, 01310100", True, True, True, True)
    If s <> "Avenida Paulista, 1000, 01310-100" Then
        Debug.Print "FALHOU - CEP/abreviacao: "; s
        falhas = falhas + 1
    End If

    s = ExtrairTrecho("Proximo a (precisao 10 metros) Rua das Flores, 123, Brasil")
    If s <> "Rua das Flores, 123" Then
        Debug.Print "FALHOU - extracao: "; s
        falhas = falhas + 1
    End If

    s = ObterLogradouroNumero("Rua A, 15, Centro, SP", ok)
    If Not ok Or s <> "Rua A, 15" Then
        Debug.Print "FALHOU - logradouro/numero: "; s
        falhas = falhas + 1
    End If

    If Not TemCEPFormatado("01001-000") Then falhas = falhas + 1
    If TemCEPFormatado("0100100") Then falhas = falhas + 1

    If falhas = 0 Then
        MsgBox "Testes locais concluídos sem falhas.", vbInformation, "modRegex"
    Else
        MsgBox falhas & " teste(s) falharam. Consulte Ctrl+G no VBA.", _
            vbExclamation, "modRegex"
    End If
End Sub
