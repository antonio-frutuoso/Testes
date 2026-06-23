<#
================================================================================
 Run-PatchTDP100.ps1
--------------------------------------------------------------------------------
 Orquestrador da automacao do patch no ambiente SAP TDP/100.

 Implementa as etapas descritas em:
   .vscode\Etapas a serem executadas no ambiente SAP TDP100.txt

 DIFERENCA em relacao a TDD/700 e TRQ/800:
   - TDD/700 e TRQ/800 usam a transacao SE01 (criar/incluir ordens).
   - TDP/100 NAO usa SE01. O fluxo aqui e:
       1. CG3Z  -> envia (upload) os arquivos da request (cofile K* e
                   data file R*) do PC (C:\PATCH) para o servidor de
                   transporte (\\sapintdevqas\sapmnt\trans).
       2. /ZUP_REQUEST -> executa o programa ZUPLOAD_REQUEST, que importa
                   a request para o ambiente TDP/100.
     Esse par (CG3Z x2 + ZUPLOAD_REQUEST) e repetido para cada request.

 Fluxo:
   1. Solicita mandante (100), usuario e SENHA MASCARADA.
   2. Solicita a pasta de origem (C:\PATCH) e o diretorio de transporte
      do servidor (\\sapintdevqas\sapmnt\trans).
   3. Solicita as requests a importar (a 1a com default TDDK907041; as
      demais conforme as etapas 5-7 do spec). ENTER em branco pula.
   4. Para cada request: valida os arquivos locais, gera dinamicamente o
      script SAP GUI (.vbs) e executa logon -> CG3Z(cofile) ->
      CG3Z(datafile) -> ZUPLOAD_REQUEST, via cscript, na MESMA sessao.
   5. Grava log com timestamp e trata erros.

 PRE-REQUISITOS:
   - SAP GUI instalado e com SAP GUI Scripting habilitado (cliente e servidor).
   - SAP Logon aberto com a conexao do TDP/100 ja iniciada na TELA DE LOGON
     (mesma condicao em que os scripts originais foram gravados).
   - Os arquivos da request existem em  C:\PATCH\<REQUEST>\  no formato
     padrao de transporte: cofile  K<seq>.<SID>  e data file  R<seq>.<SID>.
       Ex.: C:\PATCH\TDDK907041\K907041.TDD  e  C:\PATCH\TDDK907041\R907041.TDD

 SEGURANCA:
   - A senha e lida como SecureString e injetada no processo cscript apenas
     via variavel de ambiente do processo (SAP_BCODE), removida ao final.
     Os arquivos .vbs gerados NAO contem a senha.

 IMPORTANTE - IDs DE CAMPO A VERIFICAR:
   - Os scripts de TDD/700 e TRQ/800 foram GRAVADOS, garantindo os ids
     exatos da SE01. Para CG3Z e ZUPLOAD_REQUEST NAO ha gravacao; os ids
     abaixo (secao "CONFIG - IDs SAP GUI") sao a melhor aproximacao e
     DEVEM ser conferidos contra as suas telas reais. Para gravar:
       SAP GUI > Personalizar layout local (Alt+F12) > Script Recording and
       Playback > Record, execute CG3Z e /ZUP_REQUEST manualmente, e copie
       os session.findById(...) gerados para os campos $Cg3z* / $Zup* abaixo.

 USO:
   # Fluxo completo (faz logon e depois CG3Z + ZUPLOAD_REQUEST):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTDP100.ps1

   # Reaproveitar uma sessao SAP JA logada (pula o logon):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTDP100.ps1 -SomenteImportacao

 Autor: gerado para Antonio Frutuoso
================================================================================
#>

[CmdletBinding()]
param(
    # Reaproveita uma sessao SAP JA logada: pula o logon (mandante/usuario/senha)
    # e executa somente as etapas de CG3Z + ZUPLOAD_REQUEST.
    [switch]$SomenteImportacao,

    # --- Parametros opcionais p/ execucao NAO-interativa ---
    # Se informados, o prompt correspondente (Read-Host) e ignorado.
    # No modo -SomenteImportacao, informar -Requests torna a execucao 100%
    # nao-interativa (nenhum prompt). No fluxo completo, a SENHA continua
    # sendo solicitada de forma mascarada (nunca via parametro/disco).
    [string]$Mandante,
    [string]$Usuario,
    [string]$PatchDir,
    [string]$ServerTrans,
    [string[]]$Requests
)

$ErrorActionPreference = 'Stop'

# ================================================================== #
#  CONFIG - IDs SAP GUI  (VERIFICAR contra suas telas reais!)        #
#  Edite SOMENTE aqui se a gravacao mostrar ids diferentes.          #
# ================================================================== #
# CG3Z (RC1TCG3Z) - tela de upload PC -> servidor de aplicacao:
$Cg3zIdOrigem  = 'wnd[0]/usr/ctxtDY_PATH'      # arquivo de ORIGEM no frontend (PC)
$Cg3zIdDestino = 'wnd[0]/usr/ctxtDY_FILENAME'  # arquivo DESTINO no servidor de aplicacao
# ZUPLOAD_REQUEST (/ZUP_REQUEST) - campo onde se informa o numero da request:
$ZupIdRequest  = 'wnd[0]/usr/ctxtP_REQUEST'    # campo da REQUEST no programa Z

# ------------------------------------------------------------------ #
#  Configuracao de caminhos / log                                     #
# ------------------------------------------------------------------ #
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir     = Join-Path $ScriptRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$LogFile = Join-Path $LogDir ("PatchTDP100_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

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
    $file = Join-Path $env:TEMP ("patch_tdp100_{0}.vbs" -f ([guid]::NewGuid().ToString('N')))
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

# Deriva os nomes dos arquivos de transporte a partir do numero da request,
# valida que existem em <PatchDir>\<REQUEST>\ e monta os caminhos de destino
# no servidor (subpastas cofiles\ e data\, padrao SAP /usr/sap/trans).
#   Ex.: TDDK907041 -> SID=TDD, seq=907041
#        cofile  = K907041.TDD   data = R907041.TDD
function Get-RequestInfo {
    param(
        [string]$Numero,
        [string]$Rotulo,
        [string]$PatchDir,
        [string]$ServerTrans
    )
    $req = $Numero.Trim().ToUpper()
    if ($req.Length -lt 5) { throw "Numero de request invalido: '$Numero' (esperado algo como TDDK907041)." }

    $sid = $req.Substring(0,3)        # ex.: TDD
    $seq = $req.Substring(4)          # ex.: 907041 (apos as 3 letras + a letra de tipo)
    $cofileName = "K$seq.$sid"        # ex.: K907041.TDD
    $dataName   = "R$seq.$sid"        # ex.: R907041.TDD

    $folder      = Join-Path $PatchDir $req
    $localCofile = Join-Path $folder $cofileName
    $localData   = Join-Path $folder $dataName

    if (-not (Test-Path $localCofile)) { throw "Cofile nao encontrado: $localCofile" }
    if (-not (Test-Path $localData))   { throw "Data file nao encontrado: $localData" }

    # Destinos no servidor (UNC). Subpastas padrao do diretorio de transporte.
    # VERIFICAR: se o seu ZUPLOAD_REQUEST espera os arquivos direto na raiz
    # de \trans (sem cofiles\/data\), ajuste as duas linhas abaixo.
    $serverCofile = "$ServerTrans\cofiles\$cofileName"
    $serverData   = "$ServerTrans\data\$dataName"

    return [pscustomobject]@{
        Numero       = $req
        Rotulo       = $Rotulo
        LocalCofile  = $localCofile
        LocalData    = $localData
        ServerCofile = $serverCofile
        ServerData   = $serverData
    }
}

# ================================================================== #
#  Coleta de parametros                                              #
# ================================================================== #
Write-Host ""
Write-Host "=== Automacao do patch SAP TDP/100 ===" -ForegroundColor Cyan
Write-Log  "Inicio da execucao. Log: $LogFile"

$senhaSec = $null
if ($SomenteImportacao) {
    Write-Host "Modo -SomenteImportacao: reaproveitando sessao SAP ja logada (sem logon)." -ForegroundColor Yellow
    Write-Log  "Modo -SomenteImportacao ativo: logon ignorado."
}
else {
    if ($Mandante) { $mandt = $Mandante.Trim() }
    else {
        $mandt = Read-Host "Mandante (ENTER = 100)"
        if (-not $mandt.Trim()) { $mandt = '100' }
    }

    if ($Usuario) { $usuario = $Usuario.Trim() }
    else {
        $usuario = Read-Host "Usuario SAP"
        if (-not $usuario.Trim()) { throw "Usuario e obrigatorio." }
    }

    $senhaSec = Read-Host "Senha SAP" -AsSecureString
    if ($senhaSec.Length -eq 0) { throw "Senha e obrigatoria." }
}

if ($PatchDir) { $patchDir = $PatchDir.Trim() }
else {
    $patchDir = Read-Host "Pasta de origem das requests (ENTER = C:\PATCH)"
    if (-not $patchDir.Trim()) { $patchDir = 'C:\PATCH' }
}
if (-not (Test-Path $patchDir)) { throw "Pasta de origem nao encontrada: $patchDir" }

if ($ServerTrans) { $serverTrans = $ServerTrans.Trim() }
else {
    $serverTrans = Read-Host "Diretorio de transporte no servidor (ENTER = \\sapintdevqas\sapmnt\trans)"
    if (-not $serverTrans.Trim()) { $serverTrans = '\\sapintdevqas\sapmnt\trans' }
}
$serverTrans = $serverTrans.TrimEnd('\')

# --- Requests a importar (conforme etapas do spec). ENTER em branco pula. ---
$prompts = @(
    @{ Rotulo = 'Instalacao/controle';                                   Default = 'TDDK907041' },
    @{ Rotulo = 'Pacote ANTERIOR da interface Namespace';                Default = '' },
    @{ Rotulo = 'Pacote ANTERIOR do Web Service';                        Default = '' },
    @{ Rotulo = 'Customizing instalacao ANTERIOR da interface Namespace';Default = '' }
)

$reqList = @()
if ($Requests -and $Requests.Count -gt 0) {
    # Modo nao-interativo: usa as requests passadas por parametro, na ordem.
    $idx = 0
    foreach ($num in $Requests) {
        $num = $num.Trim()
        if (-not $num) { continue }
        $rotulo = if ($idx -lt $prompts.Count) { $prompts[$idx].Rotulo } else { "Request $($idx+1)" }
        $reqList += (Get-RequestInfo -Numero $num -Rotulo $rotulo -PatchDir $patchDir -ServerTrans $serverTrans)
        $idx++
    }
}
else {
    foreach ($p in $prompts) {
        $hint = if ($p.Default) { "ENTER = $($p.Default)" } else { "ENTER pula" }
        $resp = Read-Host ("Request - {0} ({1})" -f $p.Rotulo, $hint)
        $resp = $resp.Trim()
        if (-not $resp -and $p.Default) { $resp = $p.Default }
        if ($resp) {
            $reqList += (Get-RequestInfo -Numero $resp -Rotulo $p.Rotulo -PatchDir $patchDir -ServerTrans $serverTrans)
        }
    }
}

if ($reqList.Count -eq 0) { throw "Nenhuma request informada - nada a fazer." }

if ($SomenteImportacao) { Write-Log "Parametros: (SomenteImportacao)" }
else { Write-Log "Parametros: mandante=$mandt; usuario=$usuario" }
Write-Log "Origem=$patchDir; servidor=$serverTrans"
foreach ($r in $reqList) {
    Write-Log ("Request: {0} [{1}] | cofile {2} -> {3} | data {4} -> {5}" -f `
        $r.Numero, $r.Rotulo, $r.LocalCofile, $r.ServerCofile, $r.LocalData, $r.ServerData)
}

# ================================================================== #
#  Boilerplate de anexar a sessao SAP (com retry curto p/ cscript)   #
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
   WScript.Sleep 1000   ' aguarda a sessao registrar no scripting
Next
Err.Clear

If Not IsObject(session) Then
   WScript.Echo "ERRO: nenhuma sessao SAP scriptavel em application.Children(0).Children(0)." & _
      " Verifique: (1) a janela do SAP TDP/100 esta REALMENTE aberta com a tela de logon visivel;" & _
      " (2) SAP GUI Scripting habilitado no cliente (Options > Accessibility & Scripting >" & _
      " Scripting > Enable scripting, com os dois 'notify' DESMARCADOS) e no servidor;" & _
      " (3) deixe apenas a janela do TDP/100 aberta."
   WScript.Quit 1
End If

If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
On Error Resume Next
"@

# ================================================================== #
#  Bloco de LOGON (somente quando NAO e -SomenteImportacao).         #
#  Mesma abordagem da gravacao Script1_700_logon (mandante 100).     #
#  Senha via env SAP_BCODE (nunca em disco).                         #
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
#  Bloco de IMPORTACAO (gerado a partir das requests).               #
#  Para cada request: CG3Z(cofile) + CG3Z(data) + ZUPLOAD_REQUEST.   #
# ================================================================== #
$importLinhas = @()
$i = 0
foreach ($r in $reqList) {
    $i++
    $importLinhas += ""
    $importLinhas += "' ----- [$i] $($r.Rotulo): request $($r.Numero) -----"
    $importLinhas += "EnviarArquivoCG3Z `"$($r.LocalCofile)`", `"$($r.ServerCofile)`""
    $importLinhas += "EnviarArquivoCG3Z `"$($r.LocalData)`", `"$($r.ServerData)`""
    # VERIFICAR: se o ZUPLOAD_REQUEST importa o buffer inteiro de uma vez (e nao
    # uma request por execucao), mova esta chamada para FORA do loop (uma so vez).
    $importLinhas += "ImportarRequestZUP `"$($r.Numero)`""
}
$importBlock = ($importLinhas -join "`r`n")

$workMain = @"
' ===== IMPORTACAO TDP/100: CG3Z (upload) + ZUPLOAD_REQUEST =====
session.findById("wnd[0]").maximize
$importBlock

WScript.Echo "Importacao TDP/100 concluida: arquivos enviados e ZUPLOAD_REQUEST executado."
WScript.Quit 0
"@

# ================================================================== #
#  Sub-rotinas compartilhadas.                                       #
#  IDs de CG3Z / ZUPLOAD_REQUEST vem da secao CONFIG (VERIFICAR).    #
# ================================================================== #
$subs = @"
' ============================ SUB-ROTINAS ============================
Sub Falha(etapa)
   WScript.Echo "ERRO em " & etapa & ": " & Err.Description
   WScript.Quit 1
End Sub

' Envia (upload) UM arquivo do PC para o servidor de aplicacao via CG3Z.
Sub EnviarArquivoCG3Z(arqLocal, arqServidor)
   On Error Resume Next
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nCG3Z"
   session.findById("wnd[0]").sendVKey 0
   If Err.Number <> 0 Then Falha "abrir CG3Z (" & arqLocal & ")"
   Err.Clear

   ' --- Campos da tela CG3Z (VERIFICAR ids na secao CONFIG do .ps1) ---
   session.findById("$Cg3zIdOrigem").text = arqLocal
   If Err.Number <> 0 Then Falha "CG3Z: campo de ORIGEM nao encontrado (VERIFICAR id '$Cg3zIdOrigem')"
   session.findById("$Cg3zIdDestino").text = arqServidor
   If Err.Number <> 0 Then Falha "CG3Z: campo de DESTINO nao encontrado (VERIFICAR id '$Cg3zIdDestino')"
   Err.Clear

   ' Executar upload (F8). Transporte = binario; CG3Z usa binario por padrao.
   session.findById("wnd[0]").sendVKey 8
   If Err.Number <> 0 Then Falha "CG3Z: executar upload (" & arqServidor & ")"
   Err.Clear

   ' Confirma popups (sobrescrever arquivo existente, avisos, etc.)
   FecharPopups
   WScript.Echo "CG3Z: " & arqLocal & " -> " & arqServidor & " | sbar=" & session.findById("wnd[0]/sbar").text
   Err.Clear
End Sub

' Importa a request executando o programa ZUPLOAD_REQUEST (/ZUP_REQUEST).
Sub ImportarRequestZUP(req)
   On Error Resume Next
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUP_REQUEST"
   session.findById("wnd[0]").sendVKey 0
   If Err.Number <> 0 Then Falha "abrir /ZUP_REQUEST (" & req & ")"
   Err.Clear

   ' Informar a request no ZUPLOAD_REQUEST (VERIFICAR id na secao CONFIG do .ps1)
   session.findById("$ZupIdRequest").text = req
   If Err.Number <> 0 Then Falha "ZUPLOAD_REQUEST: campo da request nao encontrado (VERIFICAR id '$ZupIdRequest')"
   Err.Clear

   ' Executar (F8)
   session.findById("wnd[0]").sendVKey 8
   If Err.Number <> 0 Then Falha "ZUPLOAD_REQUEST: executar (" & req & ")"
   Err.Clear

   FecharPopups
   WScript.Echo "ZUPLOAD_REQUEST executado para " & req & " | sbar=" & session.findById("wnd[0]/sbar").text
   Err.Clear
End Sub

' Confirma/fecha popups simples (ENTER), ate 5 vezes.
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
#  Montagem do script UNICO (uma sessao, um cscript).                #
#  Ordem: anexar -> [logon] -> importacao -> sub-rotinas.            #
# ================================================================== #
if ($SomenteImportacao) {
    $mainVbs = ($attach, $workMain, $subs) -join "`r`n"
}
else {
    $mainVbs = ($attach, $logonBody, $workMain, $subs) -join "`r`n"
}

# ================================================================== #
#  Execucao                                                          #
# ================================================================== #
$mainFile = $null
try {
    $mainFile = New-VbsFile -Content $mainVbs
    Write-Log "Script temporario gerado em $env:TEMP"

    if ($SomenteImportacao) {
        Invoke-Vbs -Path $mainFile -Desc 'Patch_TDP100_Import'
    }
    else {
        # Injeta a senha apenas no ambiente do processo (nunca em disco).
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($senhaSec)
        try {
            $env:SAP_BCODE = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            Invoke-Vbs -Path $mainFile -Desc 'Patch_TDP100'
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
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
