---
name: sap-logon
description: Efetua logon automatico no SAP via SAP GUI Scripting (PowerShell, sem VBA). Use este skill quando o usuario pedir para abrir o SAP, logar no SAP, entrar em um ambiente SAP (TDD/700, DEV700, etc.), fazer login no SAP Logon ou abrir uma conexao SAP. Acione mesmo que o usuario diga apenas "logar no SAP", "abrir o SAP" ou "fazer logon".
---

# SAP Logon automatico (PowerShell — sem VBA)

## Visao Geral

Este skill efetua o logon automatico no SAP. Substitui o antigo macro VBA
`SAP_Logon` (macro do Office) por um script **orquestrado em PowerShell**, sem
qualquer dependencia de Office/VBA.

> **Nota tecnica:** o PowerShell nao consegue chamar `SAPGUI.GetScriptingEngine()`
> por late-binding COM (falha com `TYPE_E_CANTLOADLIBRARY`). Por isso a camada
> fina de SAP GUI Scripting e um VBScript minimo executado via `cscript` —
> exatamente o mesmo mecanismo dos `Run-Patch*.ps1` que ja funcionam. O objeto
> `SAPGUI` e localizado pelo wrapper oficial `SapROTWr.SapROTWrapper`. Isso
> **nao e VBA**.

O script realiza:

1. Coleta de **Conexao**, **Mandante**, **Usuario** e **Senha** (senha sempre segura).
2. Abertura do `saplogon.exe` (se necessario) e espera o objeto `SAPGUI` registrar na ROT.
3. Gera e executa um VBScript via `cscript`: obtem o Scripting Engine.
4. Abertura da conexao pelo nome exato exibido no SAP Logon.
5. Preenchimento de mandante, usuario, senha e idioma e confirmacao (Enter).

Script utilizado: `C:\Users\6121358\SAPScripts\SAP-Logon.ps1`

---

## Pre-requisitos (confirmar com o usuario)

1. SAP GUI instalado (`saplogon.exe` e o componente `SapROTWr.SapROTWrapper`, que vem com o SAP GUI).
2. SAP GUI Scripting habilitado no cliente:
   `Options > Accessibility & Scripting > Scripting > Enable scripting`
   (os dois avisos "notify" devem estar **desmarcados**).

---

## Passo 1 — Coletar informacoes

Pergunte ao usuario, **um por um**. Use os defaults quando o usuario apenas
pressionar ENTER em branco.

| # | Dado | Obrigatorio | Observacao |
|---|------|-------------|------------|
| 1 | **Nome da CONEXAO** (igual ao SAP Logon) | Sim | sem default — informar sempre |
| 2 | **Usuario SAP** | Sim | sem default — informar sempre |
| 3 | **Mandante** (cliente) | Nao | em branco = nao preenche o campo |
| 4 | **Idioma** | Nao | em branco = nao preenche (ex.: `PT`, `EN`) |
| 5 | **Senha SAP** (sensivel) | Sim | *(nao exibir no log)* |

> Conexao e Usuario nao tem valores fixos no script: o script falha se nao forem
> informados. O caminho do `saplogon.exe` usa o local padrao de instalacao do
> SAP GUI e pode ser sobrescrito com `-SapExe`.

> **Senha**: armazene apenas em memoria (variavel local) e passe via
> `$env:SAP_BCODE`. Nunca grave em arquivo, nunca passe na linha de comando,
> nunca exiba em log.

---

## Passo 2 — Executar o script

Use a ferramenta **PowerShell** (nao Bash), pois o script depende de `$env:SAP_BCODE`
e do COM do Windows.

```powershell
# Define a senha como variavel de ambiente do processo (nunca em disco)
$env:SAP_BCODE = "<senha_do_usuario>"

powershell -ExecutionPolicy Bypass -File "C:\Users\6121358\SAPScripts\SAP-Logon.ps1" `
  -Conexao  "<nome_da_conexao>" `
  -Mandante "<mandante>" `
  -Usuario  "<usuario>" `
  -Idioma   "<idioma>"

# Limpa a senha da memoria (o proprio script tambem remove ao final)
Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue
```

### Modo interativo (sem `SAP_BCODE`)
Se a senha nao for definida em `SAP_BCODE`, o script a solicita de forma
mascarada via `Read-Host -AsSecureString`.

### Reaproveitar SAP Logon ja aberto
Se o `saplogon.exe` ja estiver aberto, o script reaproveita. Para nao tentar
abrir o executavel, adicione `-SemAbrir`.

---

## Passo 3 — Interpretar o resultado

O script retorna **exit code** (`$LASTEXITCODE`):

| Code | Significado | Acao |
|------|-------------|------|
| 0 | Login enviado com sucesso | Informar o usuario; indicar o log |
| 2 | Conexao/usuario/senha nao informados | Recoletar dado faltante |
| 3 | `saplogon.exe` nao encontrado | Confirmar caminho via `-SapExe` |
| 4 | Scripting Engine indisponivel | Habilitar SAP GUI Scripting nas opcoes |
| 5 | Conexao nao pode ser aberta | Conferir o nome EXATO no SAP Logon |
| 6 | Campo de senha nao encontrado | Tela nao e a de logon; verificar conexao |
| 1 | Falha inesperada | Consultar o log em `...\SAPScripts\logs\` |

O log completo fica em `C:\Users\6121358\SAPScripts\logs\SAPLogon_<timestamp>.log`.

---

## Encadeamento com outras skills

Apos o logon bem-sucedido, a sessao SAP fica disponivel para automacoes que
reaproveitam a sessao ja logada (ex.: `sap-se10-requests` no modo `-SomenteSE10`).
