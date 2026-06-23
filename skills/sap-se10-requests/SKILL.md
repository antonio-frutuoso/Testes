---
name: sap-se10-requests
description: Automatiza a criacao de requests de Customizing no SAP TDD/700 via SE10. Use este skill quando o usuario mencionar criacao de requests no SAP TDD/700, pacotes da interface Namespace, requests de instalacao ou atualizacao SAP, SE10, ou patch TDD. Acione este skill mesmo que o usuario diga apenas "criar requests TDD", "fazer o SE10" ou "preparar o pacote Namespace".
---

# SAP TDD/700 — Criacao de Requests SE10 (Customizing)

## Visao Geral

Este skill guia a criacao automatizada das duas requests de Customizing para o pacote atual da interface Namespace no ambiente SAP TDD/700:

- **Request de INSTALACAO**: agrega o pacote atual de instalacao
- **Request de ATUALIZACAO**: agrega o pacote atual de atualizacao

O script utilizado e: `C:\Users\6121358\SAPScripts\Run-PatchTDD700.ps1`

---

## Pre-requisitos (confirmar com o usuario antes de prosseguir)

1. SAP GUI aberto com a conexao **TDD/700** na **tela de logon** (nao apenas o SAP Logon)
2. SAP GUI Scripting habilitado no cliente:  
   `Options > Accessibility & Scripting > Scripting > Enable scripting`  
   (os dois avisos "notify" devem estar **desmarcados**)

---

## Passo 1 — Coletar informacoes

Pergunte ao usuario os itens abaixo, **um por um**, na ordem apresentada.  
Use os defaults indicados quando o usuario pressionar ENTER em branco.

| # | Dado | Exemplo / Default |
|---|------|-------------------|
| 1 | **Usuario SAP** (obrigatorio) | `afrutuoso` |
| 2 | **Senha SAP** (obrigatorio, sensivel) | *(nao exibir no log)* |
| 3 | **Versao do pacote** | `v4126` |
| 4 | **Descricao da request de INSTALACAO** | `MSAF - Customizing INSTALACAO Namespace v4126` |
| 5 | **Request de Customizing INSTALACAO anterior** | `TDDK905001` |
| 6 | **Ha requests COM NOVOS AJUSTES para INSTALACAO?** (S/N) | N |
|   | *(se S)* Numeros das requests de ajuste (um por linha) | `TDDK905010` |
| 7 | **Descricao da request de ATUALIZACAO** | `MSAF - Customizing ATUALIZACAO Namespace v4126` |
| 8 | **Request de Customizing ATUALIZACAO anterior** | `TDDK905002` |
| 9 | **Ha requests COM NOVOS AJUSTES para ATUALIZACAO?** (S/N) | N |
|   | *(se S)* Numeros das requests de ajuste (um por linha) | `TDDK905020` |

> **Senha**: armazene apenas em memoria (variavel local), nunca grave em arquivo nem exiba em saida de log.

---

## Passo 2 — Executar o script

Com todos os dados coletados, execute o `Run-PatchTDD700.ps1` via PowerShell.  
Monte o comando de acordo com as respostas do usuario:

```powershell
# Define a senha como variavel de ambiente do processo (nunca em disco)
$env:SAP_BCODE = "<senha_do_usuario>"

# Monta os arrays de ajuste (apenas se houver requests de ajuste)
$ajInstal = @("TDDK905010")   # vazio se nao houver: @()
$ajAtual  = @("TDDK905020")   # vazio se nao houver: @()

# Executa o script com todos os parametros
powershell -ExecutionPolicy Bypass -File "C:\Users\6121358\SAPScripts\Run-PatchTDD700.ps1" `
  -Mandante    "700" `
  -Usuario     "<usuario>" `
  -Versao      "<versao>" `
  -DescricaoInstal "<descricao_instalacao>" `
  -ReqInstalAnterior "<request_instalacao_anterior>" `
  -AjustesInstal $ajInstal `
  -DescricaoAtual "<descricao_atualizacao>" `
  -ReqAtualAnterior "<request_atualizacao_anterior>" `
  -AjustesAtual $ajAtual

# Limpa a senha da memoria
Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue
```

**Importante**: use a ferramenta PowerShell (nao Bash) para este comando, pois o script usa sintaxe PowerShell e precisa de `$env:SAP_BCODE`.

---

## Passo 3 — Interpretar o resultado

### Sucesso
O script exibe `Concluido. Log completo: ...` em verde.  
Informe o usuario e indique o caminho do log para consulta.

### Falha — erros comuns e solucoes

| Mensagem | Causa provavel | Acao |
|----------|---------------|------|
| `nao foi possivel obter o Scripting Engine` | SAP GUI Scripting desabilitado | Habilitar nas opcoes do SAP GUI |
| `nenhuma sessao SAP scriptavel` | SAP esta na lista de sistemas, nao na tela de logon | Abrir a conexao TDD/700 ate chegar na tela de usuario/senha |
| `ERRO em abrir SE10` | Sessao esta em outra tela apos o logon | Usar `-SomenteSE10` se ja logado, ou verificar se logon foi bem-sucedido |
| `request com descricao '...' nao encontrada` | A request foi criada mas nao aparece na lista | O ID `btn[42]` (refresh) ou o caminho de busca pode precisar de ajuste — verificar IDs reais |
| `abrir dialogo incluir origem` | btn[35] diverge da sua versao do SAP | Gravar o passo via SAP GUI Script Recording e atualizar o ID em `Run-PatchTDD700.ps1` |

---

## Modo alternativo — sessao ja logada

Se o usuario ja estiver logado no SAP, pule as perguntas de usuario/senha e execute com `-SomenteSE10`:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\6121358\SAPScripts\Run-PatchTDD700.ps1" `
  -SomenteSE10 `
  -Versao "<versao>" `
  -DescricaoInstal "<descricao_instalacao>" `
  -ReqInstalAnterior "<request_instalacao_anterior>" `
  -AjustesInstal @() `
  -DescricaoAtual "<descricao_atualizacao>" `
  -ReqAtualAnterior "<request_atualizacao_anterior>" `
  -AjustesAtual @()
```

---

## Ajuste de IDs SAP GUI

Os IDs `btn[35]` (Incluir de origem) e `ctxtDV_0100_SOURCE_REQUEST` foram herdados  
do script SE01/TRQ800. Se a inclusao de origens falhar, grave o passo manualmente:

`SAP GUI > Alt+F12 > Script Recording and Playback > Record`

Execute o passo de "Incluir objeto de outra request" e copie os IDs gerados para  
as subs `IncluirOrigem` em `Run-PatchTDD700.ps1`.