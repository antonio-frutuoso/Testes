---
name: sap-patch-tdd700
description: Orquestrador end-to-end do patch SAP TDD/700 - faz o LOGON automatico e, na mesma sessao, cria as requests de Customizing (SE10). Use este skill quando o usuario pedir para "rodar o patch TDD do inicio ao fim", "logar e criar as requests TDD/700", "fazer o patch completo do Namespace", "automatizar tudo no SAP TDD" ou combinar logon + SE10 numa unica execucao. Acione mesmo que o usuario diga apenas "patch TDD completo" ou "logar e fazer o SE10".
---

# SAP TDD/700 — Patch completo (Logon + Requests SE10)

## Visao Geral

Este skill **orquestra de ponta a ponta** o patch do ambiente SAP TDD/700,
encadeando dois scripts PowerShell que ja funcionam de forma independente:

1. **`SAP-Logon.ps1`** — abre o SAP Logon, abre a conexao do TDD/700 e efetua o
   logon automatico (mandante, usuario, senha, idioma).
2. **`Run-PatchTDD700.ps1 -SomenteSE10`** — **reaproveita a sessao ja logada** e
   cria as duas requests de Customizing (INSTALACAO e ATUALIZACAO) via SE10,
   incluindo as origens.

Ambos usam SAP GUI Scripting via `cscript` (sem VBA) e leem a senha de
`$env:SAP_BCODE` — a senha **nunca** vai para disco nem linha de comando.

Scripts utilizados:
- `C:\Users\6121358\SAPScripts\SAP-Logon.ps1`
- `C:\Users\6121358\SAPScripts\Run-PatchTDD700.ps1`

> **Por que `-SomenteSE10` na 2a etapa?** O `SAP-Logon.ps1` ja deixa a sessao
> **autenticada** (passou da tela de logon). O `Run-PatchTDD700.ps1` sem esse
> switch tentaria preencher a tela de logon de novo e falharia. Com
> `-SomenteSE10`, ele apenas se anexa a sessao existente e navega para o SE10.

---

## Pre-requisitos (confirmar com o usuario)

1. SAP GUI instalado (`saplogon.exe` e o componente `SapROTWr.SapROTWrapper`).
2. SAP GUI Scripting habilitado no cliente:
   `Options > Accessibility & Scripting > Scripting > Enable scripting`
   (os dois avisos "notify" devem estar **desmarcados**).
3. Conhecer o **nome EXATO da conexao** do TDD/700 como aparece no SAP Logon
   (ex.: `1.1. Thomson Reuters TDD DEV700`).

---

## Passo 1 — Coletar informacoes

Pergunte ao usuario, **um por um**, na ordem. Use os defaults quando o usuario
apenas pressionar ENTER em branco. A **senha e coletada uma unica vez** e
reutilizada nas duas etapas.

### Dados de LOGON
| # | Dado | Obrigatorio | Default / Exemplo |
|---|------|-------------|-------------------|
| 1 | **Nome da CONEXAO** (igual ao SAP Logon) | Sim | `1.1. Thomson Reuters TDD DEV700` |
| 2 | **Usuario SAP** | Sim | `AFRUTUOSO` |
| 3 | **Mandante** | Nao | `700` |
| 4 | **Idioma** | Nao | `PT` (em branco = nao preenche) |
| 5 | **Senha SAP** (sensivel) | Sim | *(nao exibir no log)* |

### Dados do SE10 (requests)
| # | Dado | Obrigatorio | Default / Exemplo |
|---|------|-------------|-------------------|
| 6 | **Versao do pacote** | Sim | `v4126` |
| 7 | **Descricao da request de INSTALACAO** | Nao | `MSAF - Customizing INSTALACAO Namespace v4126` |
| 8 | **Request INSTALACAO anterior** | Sim | `TDDK905001` |
| 9 | **Ha ajustes para INSTALACAO?** (S/N) | Nao | N → *(se S, numeros das requests)* `TDDK905010` |
| 10 | **Descricao da request de ATUALIZACAO** | Nao | `MSAF - Customizing ATUALIZACAO Namespace v4126` |
| 11 | **Request ATUALIZACAO anterior** | Sim | `TDDK905002` |
| 12 | **Ha ajustes para ATUALIZACAO?** (S/N) | Nao | N → *(se S, numeros das requests)* `TDDK905020` |

> **Senha**: armazene apenas em memoria (variavel local) e passe via
> `$env:SAP_BCODE`. Nunca grave em arquivo, nunca passe na linha de comando,
> nunca exiba em log.

---

## Passo 2 — Executar a orquestracao

Use a ferramenta **PowerShell** (nao Bash) — os scripts dependem de
`$env:SAP_BCODE` e do COM do Windows. A senha e definida **uma vez** e vale para
as duas etapas; so e removida no final.

```powershell
# Senha definida UMA vez (nunca em disco); herdada pelas duas etapas
$env:SAP_BCODE = "<senha_do_usuario>"

$scripts = "C:\Users\6121358\SAPScripts"

# --- ETAPA 1: LOGON ---
powershell -ExecutionPolicy Bypass -File "$scripts\SAP-Logon.ps1" `
  -Conexao  "<nome_da_conexao>" `
  -Mandante "<mandante>" `
  -Usuario  "<usuario>" `
  -Idioma   "<idioma>"
$logonExit = $LASTEXITCODE

if ($logonExit -ne 0) {
    Write-Host "Logon falhou (exit $logonExit). Abortando o SE10." -ForegroundColor Red
    Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue
    return
}

# --- ETAPA 2: SE10 na sessao ja logada ---
# Monte os arrays de ajuste conforme as respostas (vazio = @())
$ajInstal = @()   # ex.: @("TDDK905010")
$ajAtual  = @()   # ex.: @("TDDK905020")

powershell -ExecutionPolicy Bypass -File "$scripts\Run-PatchTDD700.ps1" `
  -SomenteSE10 `
  -Versao            "<versao>" `
  -DescricaoInstal   "<descricao_instalacao>" `
  -ReqInstalAnterior "<request_instalacao_anterior>" `
  -AjustesInstal     $ajInstal `
  -DescricaoAtual    "<descricao_atualizacao>" `
  -ReqAtualAnterior  "<request_atualizacao_anterior>" `
  -AjustesAtual      $ajAtual
$se10Exit = $LASTEXITCODE

# Limpa a senha da memoria (os scripts tambem removem ao final)
Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue

Write-Host "Logon exit=$logonExit | SE10 exit=$se10Exit"
```

> **Importante:** so prossiga para a ETAPA 2 se a ETAPA 1 retornar `0`. Logar e
> depois tentar criar requests sem sessao valida apenas gera erros confusos.

---

## Passo 3 — Interpretar o resultado

### ETAPA 1 — Logon (`SAP-Logon.ps1`)
| Code | Significado | Acao |
|------|-------------|------|
| 0 | Login enviado com sucesso | Prosseguir para a ETAPA 2 |
| 2 | Conexao/usuario/senha nao informados | Recoletar o dado faltante |
| 3 | `saplogon.exe` nao encontrado | Confirmar caminho via `-SapExe` |
| 4 | Scripting Engine indisponivel | Habilitar SAP GUI Scripting nas opcoes |
| 5 | Conexao nao pode ser aberta | Conferir o nome EXATO no SAP Logon |
| 6 | Campo de senha nao encontrado | Tela nao e a de logon; verificar conexao |
| 1 | Falha inesperada | Consultar log em `...\SAPScripts\logs\` |

### ETAPA 2 — SE10 (`Run-PatchTDD700.ps1`)
| Mensagem | Causa provavel | Acao |
|----------|---------------|------|
| `nenhuma sessao SAP scriptavel` | Logon nao concluiu ou sessao fechada | Reexecutar a ETAPA 1 e confirmar exit 0 |
| `ERRO em abrir SE10` | Sessao em outra tela apos o logon | Garantir que o logon foi bem-sucedido |
| `request com descricao '...' nao encontrada` | `btn[42]` (refresh) ou busca divergem | Verificar IDs reais via Script Recording |
| `abrir dialogo incluir origem` | `btn[35]` diverge da sua versao do SAP | Gravar o passo e atualizar o ID no script |

Logs completos (com timestamp):
- `C:\Users\6121358\SAPScripts\logs\SAPLogon_<timestamp>.log`
- `C:\Users\6121358\SAPScripts\logs\PatchTDD700_<timestamp>.log`

---

## Modos alternativos

- **Ja logado no SAP** (pular a ETAPA 1): execute apenas o
  `Run-PatchTDD700.ps1 -SomenteSE10` — equivalente ao skill `sap-se10-requests`.
- **Apenas logon** (sem criar requests): execute apenas o `SAP-Logon.ps1` —
  equivalente ao skill `sap-logon`.
- **SAP Logon ja aberto**: adicione `-SemAbrir` na ETAPA 1 para nao reabrir o
  `saplogon.exe`.

---

## Relacao com os skills individuais

Este orquestrador **combina** os skills `sap-logon` e `sap-se10-requests` numa
unica execucao encadeada. Use-o quando o objetivo for o fluxo completo
(logar + criar requests). Para passos isolados, use os skills individuais.
