<#
================================================================================
 Run-PatchTDD700.ps1
--------------------------------------------------------------------------------
 Orquestrador da automacao do patch no ambiente SAP TDD/700.

 Processo similar ao Run-PatchTRQ800.ps1, usando a transacao SE10 (Customizing)
 em vez de SE01 (Workbench). Cria dois tipos de request para o pacote da
 interface Namespace:
   1. Request de INSTALACAO  - agrega o pacote atual de instalacao
   2. Request de ATUALIZACAO - agrega o pacote atual de atualizacao

 Fluxo:
   1. Solicita mandante (700), usuario e SENHA MASCARADA (via SAP_BCODE).
   2. Solicita versao do pacote, descrições das requests, requests ANTERIORES
      (instalacao e atualizacao) e, opcionalmente, requests COM NOVOS AJUSTES.
   3. Gera dinamicamente um VBScript SAP GUI e executa via cscript:
      logon -> SE10 (criar as 2 requests -> incluir origens).
   4. Grava log com timestamp e trata erros.

 PRE-REQUISITOS:
   - SAP GUI instalado e com SAP GUI Scripting habilitado (cliente e servidor).
   - SAP Logon aberto com a conexao do TDD/700 ja iniciada na TELA DE LOGON.
   - SAP GUI Scripting habilitado no cliente: Options > Accessibility & Scripting
     > Scripting > Enable scripting (os dois 'notify' DESMARCADOS).

 SEGURANCA:
   - A senha NUNCA e passada via linha de comando nem gravada em disco.
   - Modo interativo: solicitada via Read-Host -AsSecureString.
   - Modo nao-interativo (chamado por skill/Claude): lida da variavel de
     ambiente SAP_BCODE do processo atual (definida pelo chamador e
     removida automaticamente ao final desta execucao).

 USO INTERATIVO (terminal do usuario):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTDD700.ps1

 USO NAO-INTERATIVO (chamado por automacao/skill):
   $env:SAP_BCODE = "<senha>"
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTDD700.ps1 `
     -Mandante 700 -Usuario "usuario" -Versao "v4126" `
     -DescricaoInstal "MSAF - Customizing INSTALACAO Namespace v4126" `
     -ReqInstalAnterior "TDDK905001" `
     -AjustesInstal @("TDDK905010","TDDK905011") `
     -DescricaoAtual "MSAF - Customizing ATUALIZACAO Namespace v4126" `
     -ReqAtualAnterior "TDDK905002" `
     -AjustesAtual @("TDDK905020")

 USO SEM LOGON (sessao ja logada):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTDD700.ps1 -SomenteSE10

 IMPORTANTE - IDs SAP GUI:
   - Os IDs de SE10 foram derivados da gravacao Script1_SE10. Os IDs de
     btn[35] (Incluir de origem) e ctxtDV_0100_SOURCE_REQUEST foram herdados
     do Run-PatchTRQ800.ps1 (SE01). VERIFIQUE-OS contra as suas telas reais
     caso a inclusao de origens falhe.

 Autor: gerado para Antonio Frutuoso
================================================================================
#>

[CmdletBinding()]
param(
    # Reaproveita sessao SAP ja logada: pula mandante/usuario/senha.
    [switch]$SomenteSE10,

    # Parametros opcionais para execucao nao-interativa (via automacao/skill).
    # Senha NUNCA via parametro: use a variavel de ambiente SAP_BCODE.
    [string]$Mandante,
    [string]$Usuario,
    [string]$Versao,
    [string]$DescricaoInstal,
    [string]$ReqInstalAnterior,
    [string[]]$AjustesInstal,
    [string]$DescricaoAtual,
    [string]$ReqAtualAnterior,
    [string[]]$AjustesAtual
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ #
#  Configuracao de caminhos / log                                     #
# ------------------------------------------------------------------ #
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir     = Join-Path $ScriptRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$LogFile = Join-Path $LogDir ("PatchTDD700_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'ERRO'  { Write-Host $line -ForegroundColor Red }
        'AVISO' { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line }
    }
}

# ------------------------------------------------------------------ #
#  Funcoes auxiliares                                                 #
# ------------------------------------------------------------------ #

# Grava conteudo .vbs em arquivo temporario usando Windows-1252 (acentos).
function New-VbsFile {
    param([string]$Content)
    $file = Join-Path $env:TEMP ("patch_tdd700_{0}.vbs" -f ([guid]::NewGuid().ToString('N')))
    [System.IO.File]::WriteAllText($file, $Content, [System.Text.Encoding]::GetEncoding(1252))
    return $file
}

# Executa um .vbs via cscript, captura saida/erro e valida o exit code.
function Invoke-Vbs {
    param([string]$Path, [string]$Desc)

    Write-Log "Executando etapa: $Desc"
    $outFile = Join-Path $LogDir ("{0}.out.log" -f ($Desc -replace '\W','_'))
    $errFile = Join-Path $LogDir ("{0}.err.log" -f ($Desc -replace '\W','_'))

    $proc = Start-Process -FilePath 'cscript.exe' `
        -ArgumentList '//nologo', ('"{0}"' -f $Path) `
        -Wait -PassThru -NoNewWindow `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile

    $stdout = (Get-Content $outFile -Raw -ErrorAction SilentlyContinue)
    $stderr = (Get-Content $errFile -Raw -ErrorAction SilentlyContinue)
    if ($stdout) { Write-Log ("Saida ({0}): {1}" -f $Desc, $stdout.Trim()) }
    if ($stderr) { Write-Log ("Stderr ({0}): {1}" -f $Desc, $stderr.Trim()) 'AVISO' }

    if ($proc.ExitCode -ne 0) {
        Write-Log ("FALHA em '{0}' (exit code {1})." -f $Desc, $proc.ExitCode) 'ERRO'
        throw "Etapa '$Desc' falhou."
    }
    Write-Log "Etapa concluida: $Desc" 'OK'
}

# Pergunta S/N e, se sim, le requests de ajuste em loop.
function Read-AdjustmentRequests {
    param([string]$Tipo)
    $lista = @()
    $resp = Read-Host "Ha requests de Customizing COM NOVOS AJUSTES para $Tipo? (S/N)"
    if ($resp -match '^[Ss]') {
        Write-Host "  Informe os numeros das requests de ajuste ($Tipo). ENTER vazio encerra." -ForegroundColor Cyan
        do {
            $r = Read-Host "  Request de ajuste ($Tipo)"
            $r = if ($r) { $r.Trim() } else { '' }
            if ($r) {
                $lista += $r
                Write-Log "Ajuste ($Tipo) - request adicionada: $r"
            }
        } while ($r)
    }
    return ,$lista
}

# Monta a sequencia .vbs de inclusoes para uma ordem, localizando-a pela
# descricao unica (sem depender de indices de linha da lista ABAP).
function Build-IncludeSequence {
    param([string]$Descr, [string]$PrevReq, [string[]]$Adjs)
    $linhas = @()
    $linhas += "SelecionarOrdemPorDescricao `"$Descr`""
    $linhas += "IncluirOrigem `"$PrevReq`""
    foreach ($a in $Adjs) {
        $linhas += "SelecionarOrdemPorDescricao `"$Descr`""
        $linhas += "IncluirOrigem `"$a`""
    }
    return ($linhas -join "`r`n")
}

# ================================================================== #
#  Coleta de parametros                                              #
# ================================================================== #
Write-Host ""
Write-Host "=== Automacao do patch SAP TDD/700 - SE10 Customizing ===" -ForegroundColor Cyan
Write-Log  "Inicio da execucao. Log: $LogFile"

# Modo automacao: SAP_BCODE definido (chamado por skill/script) OU parametros passados via linha de comando.
$modoAutomacao = (-not [string]::IsNullOrEmpty([System.Environment]::GetEnvironmentVariable('SAP_BCODE', 'Process'))) -or
                 $PSBoundParameters.ContainsKey('Usuario') -or $SomenteSE10.IsPresent

$senhaSec = $null
$senhaViaEnv = $false

if ($SomenteSE10) {
    Write-Host "Modo -SomenteSE10: reaproveitando sessao SAP ja logada (sem logon)." -ForegroundColor Yellow
    Write-Log  "Modo -SomenteSE10 ativo: logon ignorado."
}
else {
    if ($Mandante) { $mandt = $Mandante.Trim() }
    else {
        $mandt = Read-Host "Mandante (ENTER = 700)"
        if ([string]::IsNullOrWhiteSpace($mandt)) { $mandt = '700' }
    }

    if ($Usuario) { $usuario = $Usuario.Trim() }
    else {
        $usuario = Read-Host "Usuario SAP"
        if ([string]::IsNullOrWhiteSpace($usuario)) { throw "Usuario e obrigatorio." }
    }

    # Senha: prefere SAP_BCODE (modo automacao) ou Read-Host (modo interativo).
    $envSenha = [System.Environment]::GetEnvironmentVariable('SAP_BCODE', 'Process')
    if (-not [string]::IsNullOrEmpty($envSenha)) {
        Write-Log "Senha obtida via variavel de ambiente SAP_BCODE (modo automacao)."
        $senhaViaEnv = $true
    }
    else {
        $senhaSec = Read-Host "Senha SAP" -AsSecureString
        if ($null -eq $senhaSec -or $senhaSec.Length -eq 0) { throw "Senha e obrigatoria." }
    }
}

if ($Versao) { $versao = $Versao.Trim() }
else {
    $versao = Read-Host "Versao do pacote (ENTER = v4126)"
    if ([string]::IsNullOrWhiteSpace($versao)) { $versao = 'v4126' }
}

# --- Request de INSTALACAO ---
if ($DescricaoInstal) { $textoInstal = $DescricaoInstal.Trim() }
else {
    $textoInstal = Read-Host "Descricao da request de INSTALACAO (ENTER = 'MSAF - Customizing INSTALACAO Namespace $versao')"
    if ([string]::IsNullOrWhiteSpace($textoInstal)) { $textoInstal = "MSAF - Customizing INSTALACAO Namespace $versao" }
}

if ($ReqInstalAnterior) { $reqInstalAnt = $ReqInstalAnterior.Trim() }
else {
    $reqInstalAnt = Read-Host "Request de Customizing INSTALACAO ANTERIOR (ex: TDDK905001)"
    if ([string]::IsNullOrWhiteSpace($reqInstalAnt)) { throw "Request de instalacao anterior e obrigatoria." }
}

if ($AjustesInstal -and $AjustesInstal.Count -gt 0) {
    $ajInstal = $AjustesInstal | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() }
} elseif ($PSBoundParameters.ContainsKey('AjustesInstal') -or $modoAutomacao) {
    $ajInstal = @()
} else {
    $ajInstal = Read-AdjustmentRequests -Tipo 'INSTALACAO'
}

# --- Request de ATUALIZACAO ---
if ($DescricaoAtual) { $textoAtual = $DescricaoAtual.Trim() }
else {
    $textoAtual = Read-Host "Descricao da request de ATUALIZACAO (ENTER = 'MSAF - Customizing ATUALIZACAO Namespace $versao')"
    if ([string]::IsNullOrWhiteSpace($textoAtual)) { $textoAtual = "MSAF - Customizing ATUALIZACAO Namespace $versao" }
}

if ($ReqAtualAnterior) { $reqAtualAnt = $ReqAtualAnterior.Trim() }
else {
    $reqAtualAnt = Read-Host "Request de Customizing ATUALIZACAO ANTERIOR (ex: TDDK905002)"
    if ([string]::IsNullOrWhiteSpace($reqAtualAnt)) { throw "Request de atualizacao anterior e obrigatoria." }
}

if ($AjustesAtual -and $AjustesAtual.Count -gt 0) {
    $ajAtual = $AjustesAtual | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() }
} elseif ($PSBoundParameters.ContainsKey('AjustesAtual') -or $modoAutomacao) {
    $ajAtual = @()
} else {
    $ajAtual = Read-AdjustmentRequests -Tipo 'ATUALIZACAO'
}

# Log dos parametros coletados.
if ($SomenteSE10) { Write-Log "Parametros: (SomenteSE10) versao=$versao" }
else { Write-Log "Parametros: mandante=$mandt; usuario=$usuario; versao=$versao" }
Write-Log "Instalacao: texto='$textoInstal'; anterior=$reqInstalAnt; ajustes=$($ajInstal -join ', ')"
Write-Log "Atualizacao: texto='$textoAtual'; anterior=$reqAtualAnt; ajustes=$($ajAtual -join ', ')"

# ================================================================== #
#  Boilerplate de anexar a sessao SAP (com retry, igual ao TRQ800)   #
# ================================================================== #
$attach = @"
Dim SapGuiAuto, application, connection, session

On Error Resume Next
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Err.Number <> 0 Or Not IsObject(application) Then
   WScript.Echo "ERRO: nao foi possivel obter o Scripting Engine do SAP GUI." & _
      " O SAP Logon esta aberto e o scripting habilitado no cliente?" & _
      " (Err " & Err.Number & ": " & Err.Description & ")"
   WScript.Quit 1
End If
Err.Clear

Dim tentativa
For tentativa = 1 To 15
   If Not IsObject(connection) Then Set connection = application.Children(0)
   If IsObject(connection) And Not IsObject(session) Then Set session = connection.Children(0)
   If IsObject(session) Then Exit For
   Err.Clear
   WScript.Sleep 1000
Next
Err.Clear

If Not IsObject(session) Then
   WScript.Echo "ERRO: nenhuma sessao SAP scriptavel encontrada." & _
      " Verifique: (1) a janela do TDD/700 esta aberta na TELA DE LOGON;" & _
      " (2) SAP GUI Scripting habilitado no cliente (Options > Accessibility &" & _
      " Scripting > Scripting > Enable scripting, os dois 'notify' DESMARCADOS)" & _
      " e no servidor; (3) deixe apenas a janela do TDD/700 aberta."
   WScript.Quit 1
End If

If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
On Error Resume Next
"@

# ================================================================== #
#  Bloco de LOGON. Senha via env SAP_BCODE (nunca em disco).         #
# ================================================================== #
$logonBody = @"
' ===== LOGON (senha via env SAP_BCODE) =====
Set shell = CreateObject("WScript.Shell")
pwd = shell.Environment("PROCESS").Item("SAP_BCODE")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "$mandt"
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "$usuario"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pwd
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then Falha "logon"
Err.Clear
WScript.Echo "Logon OK."
"@

# ================================================================== #
#  Sequencias de inclusao (geradas a partir dos parametros)          #
# ================================================================== #
$instalSeq = Build-IncludeSequence -Descr $textoInstal -PrevReq $reqInstalAnt -Adjs $ajInstal
$atualSeq  = Build-IncludeSequence -Descr $textoAtual  -PrevReq $reqAtualAnt  -Adjs $ajAtual

# ================================================================== #
#  Bloco principal do SE10 (criar requests de Customizing).          #
#  Abre SE10, exibe a lista, cria as 2 requests e inclui as origens. #
#  Localiza cada request pela DESCRICAO UNICA (sem indices de linha) #
# ================================================================== #
$se10Main = @"
' ===== SE10: criar/incluir requests de Customizing =====
Dim gFoundId

' Fecha quaisquer popups modais remanescentes antes de navegar.
Do While session.Children.Count > 1
   session.findById("wnd[1]").sendVKey 12
   WScript.Sleep 300
Loop
Err.Clear
session.findById("wnd[0]").maximize
' ===== Abrir transacao SE10 =====
session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE10"
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then Falha "abrir SE10"
Err.Clear

' ===== Exibir lista de requests de Customizing =====
session.findById("wnd[0]/usr/subCOMMONSUBSCREEN:RDDM0001:0220/btn%_AUTOTEXT028").press
If Err.Number <> 0 Then Falha "exibir lista SE10"
Err.Clear

' ===== Criar as duas requests de Customizing =====
CriarOrdem "$textoInstal"
CriarOrdem "$textoAtual"

' ===== Inclusoes na request de INSTALACAO =====
$instalSeq

' ===== Inclusoes na request de ATUALIZACAO =====
$atualSeq

WScript.Echo "SE10 OK: requests criadas e origens incluidas."
WScript.Quit 0
"@

# ================================================================== #
#  Sub-rotinas compartilhadas.                                       #
# ================================================================== #
$subs = @"
' ============================ SUB-ROTINAS ============================
Sub Falha(etapa)
   WScript.Echo "ERRO em " & etapa & ": " & Err.Description
   WScript.Quit 1
End Sub

Sub CriarOrdem(descr)
   On Error Resume Next
   session.findById("wnd[0]/tbar[1]/btn[6]").press
   If Err.Number <> 0 Then Falha "abrir dialogo criar request (" & descr & ")"
   Err.Clear

   ' Em SE10 o botao criar abre PRIMEIRO um popup de selecao de TIPO de request
   ' (radio "Ordem customizing" vs "Ordem de workbench"); so depois vem a tela
   ' de descricao. Selecionar Customizing e confirmar (btn[0]).
   If Not session.findById("wnd[1]/usr/radKO042-REQ_CUST_W", False) Is Nothing Then
      session.findById("wnd[1]/usr/radKO042-REQ_CUST_W").select
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      Err.Clear
   End If

   ' Aguarda a tela de descricao (campo AS4TEXT) aparecer.
   Dim tCri
   For tCri = 1 To 20
      If Not session.findById("wnd[1]/usr/txtKO013-AS4TEXT", False) Is Nothing Then Exit For
      WScript.Sleep 250
   Next
   If session.findById("wnd[1]/usr/txtKO013-AS4TEXT", False) Is Nothing Then Falha "campo de descricao nao apareceu (" & descr & ")"

   session.findById("wnd[1]/usr/txtKO013-AS4TEXT").text = descr
   If Err.Number <> 0 Then Falha "preencher descricao (" & descr & ")"
   Err.Clear
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   If Err.Number <> 0 Then Falha "gravar request (" & descr & ")"
   Err.Clear
   FecharPopups
   WScript.Echo "Request criada: " & descr & " | sbar=" & session.findById("wnd[0]/sbar").text
   Err.Clear
End Sub

Sub SelecionarOrdemPorDescricao(descr)
   On Error Resume Next
   ' As requests recem-criadas ja aparecem na lista; NAO usar btn[42] (refresh)
   ' aqui: ele dispara um re-render assincrono que esvazia a arvore de labels no
   ' instante da busca e faz BuscarLabel falhar. Em vez disso, busca com retry.
   Dim tSel
   For tSel = 1 To 10
      gFoundId = ""
      BuscarLabel session.findById("wnd[0]/usr"), descr
      If gFoundId <> "" Then Exit For
      WScript.Sleep 500
   Next
   If gFoundId = "" Then
      WScript.Echo "ERRO: request com descricao '" & descr & "' nao encontrada na lista."
      WScript.Quit 1
   End If
   session.findById(gFoundId).setFocus
   Err.Clear
End Sub

Sub BuscarLabel(container, alvo)
   On Error Resume Next
   Dim obj
   For Each obj In container.Children
      If gFoundId <> "" Then Exit Sub
      If obj.Type = "GuiLabel" Then
         If Trim(obj.Text) = Trim(alvo) Then
            gFoundId = obj.Id
            Exit Sub
         End If
      End If
      If InStr(obj.Type, "Container") > 0 Or obj.Type = "GuiUserArea" Then
         BuscarLabel obj, alvo
      End If
   Next
End Sub

Sub IncluirOrigem(origem)
   On Error Resume Next
   ' btn[35] = "Incluir objetos de outra ordem/request" na toolbar SE10.
   ' VERIFIQUE este ID contra a sua tela real se a inclusao falhar.
   session.findById("wnd[0]/tbar[1]/btn[35]").press
   If Err.Number <> 0 Then Falha "abrir dialogo incluir origem (" & origem & ")"
   session.findById("wnd[1]/usr/ctxtDV_0100_SOURCE_REQUEST").text = origem
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   If Err.Number <> 0 Then Falha "confirmar origem " & origem
   Err.Clear
   FecharPopups
   WScript.Echo "Incluida origem " & origem & " | sbar=" & session.findById("wnd[0]/sbar").text
   Err.Clear
End Sub

Sub FecharPopups
   On Error Resume Next
   Dim n
   n = 0
   Do While session.Children.Count > 1 And n < 5
      session.findById("wnd[1]").sendVKey 0
      n = n + 1
   Loop
   Err.Clear
End Sub
"@

# ================================================================== #
#  Montagem do script VBS unico (logon + SE10 na MESMA sessao).      #
# ================================================================== #
if ($SomenteSE10) {
    $mainVbs = ($attach, $se10Main, $subs) -join "`r`n"
}
else {
    $mainVbs = ($attach, $logonBody, $se10Main, $subs) -join "`r`n"
}

# ================================================================== #
#  Execucao                                                          #
# ================================================================== #
$mainFile = $null
try {
    $mainFile = New-VbsFile -Content $mainVbs
    Write-Log "Script temporario gerado em $env:TEMP"

    if ($SomenteSE10) {
        Invoke-Vbs -Path $mainFile -Desc 'PatchTDD700_SE10'
    }
    else {
        # Injeta a senha no ambiente do processo (nunca em disco).
        $bstr = $null
        try {
            if ($senhaViaEnv) {
                # Senha ja esta em SAP_BCODE (definida pelo chamador): usa direto.
                Invoke-Vbs -Path $mainFile -Desc 'PatchTDD700'
            }
            else {
                # Modo interativo: converte SecureString e injeta.
                $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($senhaSec)
                $env:SAP_BCODE = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                Invoke-Vbs -Path $mainFile -Desc 'PatchTDD700'
            }
        }
        finally {
            if ($bstr) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
            Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue
        }
    }

    Write-Log "Automacao concluida com sucesso." 'OK'
    Write-Host ""
    Write-Host "Concluido. Log completo: $LogFile" -ForegroundColor Green
}
catch {
    Write-Log ("Execucao interrompida: {0}" -f $_.Exception.Message) 'ERRO'
    Write-Host ""
    Write-Host "ERRO. Verifique o log: $LogFile" -ForegroundColor Red
    exit 1
}
finally {
    if ($mainFile -and (Test-Path $mainFile)) { Remove-Item $mainFile -Force -ErrorAction SilentlyContinue }
    if ($senhaSec) { $senhaSec.Dispose() }
}
