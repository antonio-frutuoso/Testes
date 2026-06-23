<#
================================================================================
 Run-PatchTRQ800.ps1
--------------------------------------------------------------------------------
 Orquestrador da automacao do patch no ambiente SAP TRQ/800.

 Implementa as etapas descritas em:
   .vscode\Automatizar_patch_TRQ800

 Processo SIMILAR ao Run-PatchTDD700.ps1, mudando o ambiente de TDD/700 para
 TRQ/800 e usando requests de WORKBENCH (em vez de Customizing) para dois
 pacotes:
   1. Pacote da interface Namespace
   2. Pacote do Web Service

 Fluxo:
   1. Solicita mandante, usuario e SENHA MASCARADA.
   2. Solicita versao do pacote, as descricoes das duas novas requests e as
      requests de Workbench ANTERIORES (namespace e web service).
   3. Pergunta (S/N) + loop pelas requests COM OS OBJETOS corrigidos/atualizados,
      separadamente para namespace e para web service.
   4. Gera dinamicamente os scripts SAP GUI (.vbs) a partir de templates.
   5. Executa: logon -> SE01 (criar/incluir requests), via cscript.
   6. Grava log com timestamp e trata erros.

 PRE-REQUISITOS:
   - SAP GUI instalado e com SAP GUI Scripting habilitado (cliente e servidor).
   - SAP Logon aberto com a conexao do TRQ/800 ja iniciada na TELA DE LOGON
     (mesma condicao em que os scripts originais foram gravados).

 SEGURANCA:
   - A senha e lida como SecureString e injetada no processo cscript apenas
     via variavel de ambiente do processo (SAP_BCODE), removida ao final.
     Os arquivos .vbs gerados NAO contem a senha.

 ESTRATEGIA DE LOCALIZACAO:
   - A lista de ordens da SE01 (tela 120) e uma lista ABAP cujas linhas mudam
     a cada execucao; por isso NAO se usa indice de linha (lbl[col,linha]).
     As ordens recem-criadas sao localizadas pela DESCRICAO unica que o
     proprio script define, tornando a automacao reproduzivel.
   - Pre-requisito: a descricao de cada ordem deve ser unica no momento da
     execucao (rodar duas vezes no mesmo dia/versao gera descricoes iguais).

 OBSERVACAO sobre requests de Workbench:
   - O filtro da SE01 marca REQ_WB (Workbench) em vez de REQ_CUST (Customizing).
   - Ao criar a ordem (btn[6]), se o SAP exibir um dialogo de SELECAO DE TIPO
     de request (Workbench x Customizing) antes da tela de descricao, capture o
     ID desse controle e ajuste a Sub CriarOrdem. Na gravacao base (TDD) a tela
     de descricao (KO013-AS4TEXT) abriu diretamente.

 USO:
   # Fluxo completo (faz logon e depois SE01):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTRQ800.ps1

   # Reaproveitar uma sessao SAP JA logada (pula o logon, so SE01):
   powershell -ExecutionPolicy Bypass -File .\Run-PatchTRQ800.ps1 -SomenteSE01

 Autor: gerado para Antonio Frutuoso
================================================================================
#>

[CmdletBinding()]
param(
    # Reaproveita uma sessao SAP JA logada: pula o logon (mandante/usuario/senha)
    # e executa somente a etapa SE01 (criar/incluir ordens).
    [switch]$SomenteSE01
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ #
#  Configuracao de caminhos / log                                     #
# ------------------------------------------------------------------ #
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir     = Join-Path $ScriptRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$LogFile = Join-Path $LogDir ("PatchTRQ800_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

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
    $file = Join-Path $env:TEMP ("patch_trq800_{0}.vbs" -f ([guid]::NewGuid().ToString('N')))
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

# Pergunta S/N e, se sim, le requests (que contem os objetos corrigidos) em loop.
function Read-AdjustmentRequests {
    param([string]$Tipo)
    $lista = @()
    $resp = Read-Host "Ha requests Workbench COM OS OBJETOS corrigidos/atualizados para $Tipo? (S/N)"
    if ($resp -match '^[Ss]') {
        Write-Host "  Informe os numeros das requests com objetos ($Tipo). ENTER vazio encerra." -ForegroundColor Cyan
        do {
            $r = Read-Host "  Request de objetos ($Tipo)"
            $r = if ($r) { $r.Trim() } else { '' }
            if ($r) {
                $lista += $r
                Write-Log "Objetos ($Tipo) - request adicionada: $r"
            }
        } while ($r)
    }
    return ,$lista
}

# Monta a sequencia .vbs de inclusoes para uma ordem, localizando-a pela
# descricao unica (sem depender de indices de linha da lista).
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
Write-Host "=== Automacao do patch SAP TRQ/800 ===" -ForegroundColor Cyan
Write-Log  "Inicio da execucao. Log: $LogFile"

$senhaSec = $null
if ($SomenteSE01) {
    Write-Host "Modo -SomenteSE01: reaproveitando sessao SAP ja logada (sem logon)." -ForegroundColor Yellow
    Write-Log  "Modo -SomenteSE01 ativo: logon ignorado."
}
else {
    $mandt = Read-Host "Mandante (ENTER = 800)"
    if ([string]::IsNullOrWhiteSpace($mandt)) { $mandt = '800' }

    $usuario = Read-Host "Usuario SAP"
    if ([string]::IsNullOrWhiteSpace($usuario)) { throw "Usuario e obrigatorio." }

    $senhaSec = Read-Host "Senha SAP" -AsSecureString
    if ($null -eq $senhaSec -or $senhaSec.Length -eq 0) { throw "Senha e obrigatoria." }
}

$versao = Read-Host "Versao do pacote (ENTER = V4126)"
if ([string]::IsNullOrWhiteSpace($versao)) { $versao = 'V4126' }

# --- Pacote da interface Namespace ---
$textoNs = Read-Host "Descricao da request Workbench - INTERFACE NAMESPACE (ENTER = 'MSAF - INTERFACE NAMESPACE MSAF $versao')"
if ([string]::IsNullOrWhiteSpace($textoNs)) { $textoNs = "MSAF - INTERFACE NAMESPACE MSAF $versao" }

$reqNsAnterior = Read-Host "Request Workbench ANTERIOR do pacote Interface Namespace (ENTER = TRQK901251)"
if ([string]::IsNullOrWhiteSpace($reqNsAnterior)) { $reqNsAnterior = 'TRQK901251' }

# --- Pacote do Web Service ---
$textoWs = Read-Host "Descricao da request Workbench - WEB SERVICE (ENTER = 'MSAF - PACOTE WEBSERVICE MSAF $versao')"
if ([string]::IsNullOrWhiteSpace($textoWs)) { $textoWs = "MSAF - PACOTE WEBSERVICE MSAF $versao" }

$reqWsAnterior = Read-Host "Request Workbench ANTERIOR do pacote Web Service (ENTER = TRQK901253)"
if ([string]::IsNullOrWhiteSpace($reqWsAnterior)) { $reqWsAnterior = 'TRQK901253' }

# --- Requests com objetos corrigidos/atualizados ---
$objsNs = Read-AdjustmentRequests -Tipo 'INTERFACE NAMESPACE'
$objsWs = Read-AdjustmentRequests -Tipo 'WEB SERVICE'

if ($SomenteSE01) { Write-Log "Parametros: (SomenteSE01) versao=$versao" }
else { Write-Log "Parametros: mandante=$mandt; usuario=$usuario; versao=$versao" }
Write-Log "Namespace: texto='$textoNs'; anterior=$reqNsAnterior; objetos=$($objsNs -join ', ')"
Write-Log "WebService: texto='$textoWs'; anterior=$reqWsAnterior; objetos=$($objsWs -join ', ')"

# ================================================================== #
#  Boilerplate de anexar a sessao SAP                                #
# ================================================================== #
$attach = @"
' ----- Anexa a sessao SAP (mesma abordagem da gravacao Script1_800_logon) -----
'  A gravacao NAO testa Children.Count: indexa direto application.Children(0)
'  e connection.Children(0). Em algumas instalacoes Children.Count volta 0
'  mesmo havendo sessao scriptavel, e a varredura por Count falha sempre;
'  por isso aqui replicamos o acesso direto por indice, so com um retry curto
'  porque, via cscript, a sessao pode levar um instante para ficar scriptavel.
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
      " Verifique: (1) a janela do SAP TRQ/800 esta REALMENTE aberta com a tela de logon visivel," & _
      " e nao apenas o SAP Logon (a lista de sistemas);" & _
      " (2) SAP GUI Scripting habilitado no cliente (Options > Accessibility & Scripting >" & _
      " Scripting > Enable scripting, com os dois avisos de 'notify' DESMARCADOS) e no servidor;" & _
      " (3) se usar varias janelas/SAP Business Client, deixe so a janela do TRQ/800 aberta."
   WScript.Quit 1
End If

If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
On Error Resume Next
"@

# ================================================================== #
#  Sequencias de inclusao (geradas a partir dos parametros)          #
# ================================================================== #
$nsSeq = Build-IncludeSequence -Descr $textoNs -PrevReq $reqNsAnterior -Adjs $objsNs
$wsSeq = Build-IncludeSequence -Descr $textoWs -PrevReq $reqWsAnterior -Adjs $objsWs

# ================================================================== #
#  Bloco de LOGON (somente quando NAO e -SomenteSE01).               #
#  Roda na MESMA sessao/processo do SE01 (sem re-anexar), igual a    #
#  gravacao Script1_800_logon. Senha via env SAP_BCODE.              #
# ================================================================== #
$logonBody = @"
' ===== LOGON (senha via env SAP_BCODE, nunca em disco) =====
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
#  Bloco principal do SE01 (criar/incluir ordens Workbench).         #
#  Localiza a ordem pela DESCRICAO unica, sem depender de indices     #
#  de linha (lbl[col,linha]) da lista ABAP.                          #
# ================================================================== #
$se01Main = @"
' ===== SE01: criar/incluir ordens =====
Dim gFoundId
Dim base
base = "wnd[0]/usr/tabsMAINTABSTRIP/tabpTSCM/ssubCOMMONSUBSCREEN:RDDM0001:0220/"

session.findById("wnd[0]").maximize
' ===== Abrir transacao SE01 =====
session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE01"
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then Falha "abrir SE01"
Err.Clear

' ===== Selecionar aba TSCM (Transportes) e filtrar ordens de WORKBENCH =====
session.findById("wnd[0]/usr/tabsMAINTABSTRIP/tabpTSCM").select
session.findById(base & "chkTRDYSE01CM-REQ_WB").selected   = true
session.findById(base & "chkTRDYSE01CM-REQ_CUST").selected = false
session.findById(base & "chkTRDYSE01CM-REQ_COP").selected  = false
session.findById(base & "btn%_AUTOTEXT028").press
If Err.Number <> 0 Then Falha "filtrar/Exibir lista SE01"
Err.Clear

' ===== Criar as duas ordens de Workbench =====
CriarOrdem "$textoNs"
CriarOrdem "$textoWs"

' ===== Inclusoes na ordem da INTERFACE NAMESPACE =====
$nsSeq

' ===== Inclusoes na ordem do WEB SERVICE =====
$wsSeq

WScript.Echo "SE01 OK: ordens criadas e inclusoes aplicadas."
WScript.Quit 0
"@

# ================================================================== #
#  Sub-rotinas compartilhadas (logon + SE01).                        #
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
   If Err.Number <> 0 Then Falha "abrir dialogo criar ordem (" & descr & ")"
   ' Se o SAP abrir um dialogo de selecao de TIPO de request (Workbench x
   ' Customizing) antes da descricao, trate-o aqui antes de preencher AS4TEXT.
   session.findById("wnd[1]/usr/txtKO013-AS4TEXT").text = descr
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   If Err.Number <> 0 Then Falha "gravar ordem (" & descr & ")"
   Err.Clear
   FecharPopups
   WScript.Echo "Ordem criada: " & descr & " | sbar=" & session.findById("wnd[0]/sbar").text
   Err.Clear
End Sub

Sub SelecionarOrdemPorDescricao(descr)
   On Error Resume Next
   ' atualiza a lista para refletir as ordens recem-criadas
   session.findById("wnd[0]/tbar[1]/btn[42]").press
   Err.Clear
   gFoundId = ""
   BuscarLabel session.findById("wnd[0]/usr"), descr
   If gFoundId = "" Then
      WScript.Echo "ERRO: ordem com descricao '" & descr & "' nao encontrada na lista."
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
   session.findById("wnd[0]/tbar[1]/btn[35]").press
   If Err.Number <> 0 Then Falha "abrir 'Incluir objetos' (origem " & origem & ")"
   session.findById("wnd[1]/usr/ctxtDV_0100_SOURCE_REQUEST").text = origem
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   If Err.Number <> 0 Then Falha "informar origem " & origem
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
#  Montagem do script UNICO (uma sessao, um cscript).                #
#  Ordem: anexar -> [logon] -> SE01 -> sub-rotinas.                  #
#  As Subs ficam apos o WScript.Quit 0 (nunca executadas em linha,   #
#  apenas chamadas), e ficam visiveis para todo o script.            #
# ================================================================== #
if ($SomenteSE01) {
    $mainVbs = ($attach, $se01Main, $subs) -join "`r`n"
}
else {
    $mainVbs = ($attach, $logonBody, $se01Main, $subs) -join "`r`n"
}

# ================================================================== #
#  Execucao (UM unico cscript: logon + SE01 na MESMA sessao)         #
# ================================================================== #
$mainFile = $null
try {
    $mainFile = New-VbsFile -Content $mainVbs
    Write-Log "Script temporario gerado em $env:TEMP"

    if ($SomenteSE01) {
        # Reaproveita a sessao ja logada: executa direto.
        Invoke-Vbs -Path $mainFile -Desc 'Patch_TRQ800_SE01'
    }
    else {
        # Injeta a senha apenas no ambiente do processo (nunca em disco).
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($senhaSec)
        try {
            $env:SAP_BCODE = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            Invoke-Vbs -Path $mainFile -Desc 'Patch_TRQ800'
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
