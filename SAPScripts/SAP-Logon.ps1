<#
================================================================================
 SAP-Logon.ps1
--------------------------------------------------------------------------------
 Abre o SAP Logon e efetua login automaticamente no SAP - SEM VBA (sem macro
 do Office). Orquestrado em PowerShell; a camada fina de SAP GUI Scripting e
 um VBScript minimo executado via cscript (mesmo mecanismo dos Run-Patch*.ps1).

 POR QUE NAO E PURO POWERSHELL:
   O PowerShell nao consegue chamar SAPGUI.GetScriptingEngine() por late-binding
   COM (falha com TYPE_E_CANTLOADLIBRARY 0x80029C4A). O cscript/VBScript usa
   IDispatch puro e funciona. Para localizar o objeto "SAPGUI" usamos o wrapper
   oficial do SAP "SapROTWr.SapROTWrapper" (GetObject("SAPGUI") por moniker
   tambem falha neste ambiente).

 Equivalente ao antigo macro VBA "SAP_Logon", agora sem dependencia do Office:
   1. Coleta Conexao, Mandante, Usuario e Senha (senha sempre segura).
   2. Abre o saplogon.exe (se ainda nao estiver aberto) e espera o objeto
      "SAPGUI" registrar na Running Object Table (ROT).
   3. Gera um VBScript que: obtem o Scripting Engine -> abre a conexao pelo
      nome -> preenche mandante/usuario/senha/idioma -> confirma (Enter).
   4. Executa via cscript, captura o resultado e trata erros.

 PRE-REQUISITOS:
   - SAP GUI instalado (inclui o componente SapROTWr.SapROTWrapper).
   - SAP GUI Scripting habilitado no cliente:
       Options > Accessibility & Scripting > Scripting > Enable scripting
       (os dois avisos "notify" DESMARCADOS).

 SEGURANCA:
   - A senha NUNCA e passada via linha de comando nem gravada em disco.
   - O VBScript le a senha da variavel de ambiente SAP_BCODE (herdada pelo
     processo filho cscript); o .vbs em disco NAO contem a senha.
   - Modo interativo: senha via Read-Host -AsSecureString.
   - Modo automacao (skill/Claude): senha via SAP_BCODE definida pelo chamador.

 USO INTERATIVO:
   powershell -ExecutionPolicy Bypass -File .\SAP-Logon.ps1 `
     -Conexao "1.1. Thomson Reuters TDD DEV700" -Usuario "AFRUTUOSO"

 USO NAO-INTERATIVO (automacao/skill):
   $env:SAP_BCODE = "<senha>"
   powershell -ExecutionPolicy Bypass -File .\SAP-Logon.ps1 `
     -Conexao "1.1. Thomson Reuters TDD DEV700" -Usuario "AFRUTUOSO" `
     -Mandante "700" -Idioma "PT"
   Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue

 RETORNO (exit code; use $LASTEXITCODE):
   0 sucesso | 2 parametro faltando | 3 saplogon.exe nao encontrado
   4 Scripting Engine indisponivel | 5 conexao nao abriu
   6 campo de senha nao encontrado | 1 falha inesperada

 Autor: gerado para Antonio Frutuoso (conversao do VBA SAP_Logon).
================================================================================
#>

[CmdletBinding()]
param(
    # Nome da conexao EXATAMENTE como aparece no SAP Logon. OBRIGATORIO.
    [Parameter(Mandatory = $true)]
    [string]$Conexao,

    # Usuario SAP. OBRIGATORIO.
    [Parameter(Mandatory = $true)]
    [string]$Usuario,

    # Mandante (cliente). Em branco = nao preenche o campo.
    [string]$Mandante = "",

    # Idioma de logon (PT, EN, ...). Em branco = nao preenche.
    [string]$Idioma = "",

    # Caminho do saplogon.exe.
    [string]$SapExe = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",

    # Timeout (segundos) para o objeto SAPGUI registrar na ROT.
    [int]$TimeoutSeg = 30,

    # Nao abre o saplogon.exe (assume que ja esta aberto).
    [switch]$SemAbrir
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ #
#  Log                                                                #
# ------------------------------------------------------------------ #
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir     = Join-Path $ScriptRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir ("SAPLogon_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

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

function ConvertFrom-SecureToPlain {
    param([System.Security.SecureString]$Secure)
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try   { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}

# Aguarda o objeto "SAPGUI" aparecer na ROT (via SapROTWr; so deteccao de presenca).
function Wait-SapGuiRot {
    param([int]$Seg)
    for ($i = 0; $i -lt $Seg; $i++) {
        Start-Sleep -Seconds 1
        try {
            $rot = New-Object -ComObject SapROTWr.SapROTWrapper
            $obj = $rot.GetROTEntry("SAPGUI")
            if ($obj) { return $true }
        } catch { }
    }
    return $false
}

Write-Host ""
Write-Host "=== SAP Logon automatico (PowerShell + cscript / SAP GUI Scripting) ===" -ForegroundColor Cyan
Write-Log  "Inicio da execucao. Log: $LogFile"

# ------------------------------------------------------------------ #
#  1) Parametros / senha                                              #
# ------------------------------------------------------------------ #
$Conexao  = $Conexao.Trim()
$Usuario  = $Usuario.Trim()
$Mandante = $Mandante.Trim()
$Idioma   = $Idioma.Trim()

if ([string]::IsNullOrWhiteSpace($Conexao)) { Write-Log "Conexao nao informada." 'ERRO'; exit 2 }
if ([string]::IsNullOrWhiteSpace($Usuario)) { Write-Log "Usuario nao informado." 'ERRO'; exit 2 }

$senhaEnv = [System.Environment]::GetEnvironmentVariable('SAP_BCODE', 'Process')
if (-not [string]::IsNullOrEmpty($senhaEnv)) {
    $senha = $senhaEnv
    Write-Log "Senha obtida via variavel de ambiente SAP_BCODE."
} else {
    $sec   = Read-Host "Digite a senha SAP" -AsSecureString
    $senha = ConvertFrom-SecureToPlain -Secure $sec
}
if ([string]::IsNullOrEmpty($senha)) { Write-Log "Senha nao informada." 'ERRO'; exit 2 }

Write-Log ("Conexao: '{0}' | Mandante: '{1}' | Usuario: '{2}' | Idioma: '{3}'" -f $Conexao, $Mandante, $Usuario, $Idioma)

$vbsFile = $null
try {
    # -------------------------------------------------------------- #
    #  2) Abre o SAP Logon e aguarda registro na ROT                  #
    # -------------------------------------------------------------- #
    if (-not $SemAbrir) {
        if (-not (Test-Path $SapExe)) { Write-Log ("saplogon.exe nao encontrado em: {0}" -f $SapExe) 'ERRO'; exit 3 }
        if (-not (Get-Process -Name 'saplogon' -ErrorAction SilentlyContinue)) {
            Write-Log "Abrindo o SAP Logon..."
            Start-Process -FilePath $SapExe | Out-Null
        } else {
            Write-Log "SAP Logon ja esta aberto; reaproveitando." 'AVISO'
        }
    }

    Write-Log "Aguardando o objeto SAPGUI na Running Object Table..."
    if (-not (Wait-SapGuiRot -Seg $TimeoutSeg)) {
        Write-Log "Objeto SAPGUI nao registrou na ROT. Verifique se o SAP GUI Scripting esta habilitado." 'ERRO'
        exit 4
    }
    Write-Log "Objeto SAPGUI presente na ROT." 'OK'

    # -------------------------------------------------------------- #
    #  3) Gera o VBScript de logon                                    #
    #     Conexao/mandante/usuario/idioma e SENHA vem de variaveis    #
    #     de ambiente (nada sensivel gravado no .vbs).                #
    # -------------------------------------------------------------- #
    $vbs = @'
Option Explicit
Dim sh, conexao, mandante, usuario, idioma, senha
Dim rot, sapgui, app, connection, session, pwdField
Set sh = CreateObject("WScript.Shell")
conexao  = sh.Environment("PROCESS").Item("SAP_CONN")
mandante = sh.Environment("PROCESS").Item("SAP_MANDT")
usuario  = sh.Environment("PROCESS").Item("SAP_USER")
idioma   = sh.Environment("PROCESS").Item("SAP_LANG")
senha    = sh.Environment("PROCESS").Item("SAP_BCODE")

On Error Resume Next
Set rot = CreateObject("SapROTWr.SapROTWrapper")
If Err.Number <> 0 Then WScript.Echo "ERRO|engine|" & Err.Description : WScript.Quit 4
Set sapgui = rot.GetROTEntry("SAPGUI")
If sapgui Is Nothing Then WScript.Echo "ERRO|engine|SAPGUI nao encontrado na ROT" : WScript.Quit 4
Set app = sapgui.GetScriptingEngine
If Err.Number <> 0 Then WScript.Echo "ERRO|engine|" & Err.Description : WScript.Quit 4

Err.Clear
Set connection = app.OpenConnection(conexao, True)
If Err.Number <> 0 Or connection Is Nothing Then WScript.Echo "ERRO|conn|" & Err.Description : WScript.Quit 5
Set session = connection.Children(0)
WScript.Sleep 2000
session.findById("wnd[0]").maximize

If Not session.findById("wnd[1]", False) Is Nothing Then session.findById("wnd[1]").sendVKey 0

If mandante <> "" Then
  If Not session.findById("wnd[0]/usr/txtRSYST-MANDT", False) Is Nothing Then session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = mandante
End If

If session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing Then WScript.Echo "ERRO|tela|campo de usuario nao encontrado (sessao nao esta na tela de logon)" : WScript.Quit 6
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario

Set pwdField = session.findById("wnd[0]/usr/pwdRSYST-BCODE", False)
If pwdField Is Nothing Then Set pwdField = session.findById("wnd[0]/usr/txtRSYST-BCODE", False)
If pwdField Is Nothing Then WScript.Echo "ERRO|pwd|campo de senha nao encontrado" : WScript.Quit 6
pwdField.Text = senha

If idioma <> "" Then
  If Not session.findById("wnd[0]/usr/txtRSYST-LANGU", False) Is Nothing Then session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = idioma
End If

session.findById("wnd[0]").sendVKey 0
WScript.Echo "OK|login enviado"
WScript.Quit 0
'@

    $vbsFile = Join-Path $env:TEMP ("sap_logon_{0}.vbs" -f ([guid]::NewGuid().ToString('N')))
    [System.IO.File]::WriteAllText($vbsFile, $vbs, [System.Text.Encoding]::GetEncoding(1252))

    # -------------------------------------------------------------- #
    #  4) Executa via cscript com os dados em variaveis de ambiente   #
    # -------------------------------------------------------------- #
    $env:SAP_CONN  = $Conexao
    $env:SAP_MANDT = $Mandante
    $env:SAP_USER  = $Usuario
    $env:SAP_LANG  = $Idioma
    $env:SAP_BCODE = $senha   # herdada pelo cscript; removida no finally

    $outFile = Join-Path $LogDir 'logon.out.log'
    $errFile = Join-Path $LogDir 'logon.err.log'
    $proc = Start-Process -FilePath "$env:WINDIR\System32\cscript.exe" `
        -ArgumentList '//nologo', ('"{0}"' -f $vbsFile) `
        -Wait -PassThru -NoNewWindow `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile

    $stdout = (Get-Content $outFile -Raw -ErrorAction SilentlyContinue)
    $stderr = (Get-Content $errFile -Raw -ErrorAction SilentlyContinue)
    if ($stdout) { Write-Log ("Saida cscript: {0}" -f $stdout.Trim()) }
    if ($stderr) { Write-Log ("Stderr cscript: {0}" -f $stderr.Trim()) 'AVISO' }

    switch ($proc.ExitCode) {
        0 { Write-Log "Login enviado com sucesso." 'OK'; exit 0 }
        4 { Write-Log "Scripting Engine indisponivel (habilite o SAP GUI Scripting)." 'ERRO'; exit 4 }
        5 { Write-Log ("Nao foi possivel abrir a conexao '{0}'. Confira o nome exato no SAP Logon." -f $Conexao) 'ERRO'; exit 5 }
        6 { Write-Log "Campo de logon/senha nao encontrado (a sessao nao esta na tela de logon)." 'ERRO'; exit 6 }
        default { Write-Log ("Falha no logon (exit cscript {0})." -f $proc.ExitCode) 'ERRO'; exit 1 }
    }
}
catch {
    Write-Log ("FALHA inesperada: {0}" -f $_.Exception.Message) 'ERRO'
    exit 1
}
finally {
    # Limpeza: senha em memoria, variaveis de ambiente e arquivo .vbs.
    $senha = ""
    Remove-Item Env:\SAP_BCODE -ErrorAction SilentlyContinue
    Remove-Item Env:\SAP_CONN  -ErrorAction SilentlyContinue
    Remove-Item Env:\SAP_MANDT -ErrorAction SilentlyContinue
    Remove-Item Env:\SAP_USER  -ErrorAction SilentlyContinue
    Remove-Item Env:\SAP_LANG  -ErrorAction SilentlyContinue
    if ($vbsFile -and (Test-Path $vbsFile)) { Remove-Item $vbsFile -ErrorAction SilentlyContinue }
}
