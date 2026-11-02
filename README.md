# modRegex

Suplemento para Microsoft Excel desenvolvido em VBA para correção, padronização e consulta de endereços.

## Funcionalidades

- Correção de endereços utilizando Regex.
- Padronização de textos, abreviações e CEPs.
- Processamento de células e intervalos selecionados.
- Consulta de endereços por CEP utilizando o ViaCEP.
- Cópia e inserção dos resultados no Excel.

## Instalação

1. Baixe o arquivo `modRegex.xlam` na seção **Releases**.
2. No Excel, acesse **Arquivo > Opções > Suplementos**.
3. Em **Gerenciar: Suplementos do Excel**, clique em **Ir > Procurar**.
4. Selecione o arquivo `modRegex.xlam` e habilite o suplemento.
5. Em **Personalizar Faixa de Opções**, adicione a macro `AbrirFormulario` a um botão.

**Observação:** utilize suplementos VBA somente de fontes confiáveis. Dependendo das configurações de segurança, o Excel poderá bloquear macros de arquivos baixados da internet.

## Requisitos

- Microsoft Excel Desktop para Windows.
- Suporte a macros VBA.
- Conexão com a internet para consultas ViaCEP.

## Desenvolvimento

[VBAdev2](https://github.com/VBAdev2)

## API utilizada

[ViaCEP](https://viacep.com.br/)

## Licença

MIT, caso essa seja a licença escolhida para o repositório.
