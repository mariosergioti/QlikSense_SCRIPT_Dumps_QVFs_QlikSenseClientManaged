<#
.SYNOPSIS
Cubotimize - Backup de Aplicações Qlik Sense Client-Managed (QVF)

.DESCRIPTION
Script para controle de backups dos apps (Publicados e Não Publicados) do Qlik Sense Client-Managed.
Realiza o dump direto para o Storage de Rede (ClusterFS), gerencia retenção e envia relatório executivo 
em HTML via E-mail com ícones de status (Alto Contraste) e cores dinâmicas.
Script universal: identifica automaticamente o servidor (hostname) para direcionar o destino.

.NOTES
Versão: 2.4.0
Licença: MIT License
Créditos: Mario Sergio Soares
Bio Page: https://cubo.plus/mariosergioti
Direitos Reservados: https://cubotimize.com
Versão Data: 16/03/2026
#>

# =================================================================
# CONFIGURAÇÕES DO AMBIENTE E RETENÇÃO
# =================================================================
# Captura o nome do servidor automaticamente em Maiúsculas (Ex: MURIAE ou IRAJA)
$vServidorNome      = $(hostname).ToUpper() 
$vServidorQlik      = "localhost" 

# O caminho da rede adapta-se sozinho ao servidor onde está a rodar
$vPastaBackup       = "\\TROCAR\BACKUP\QLIK\QLIK_SENSE\$vServidorNome\Apps_QVF\" #Finalize com "\"
$vDiasBackup        = 60 
$vSemDados          = $true # $true para baixar SEM DADOS | $false para COM DADOS

# =================================================================
# CONFIGURAÇÕES DE E-MAIL (GMAIL)
# =================================================================
$vEnviarEmail       = $true 
$vSmtpServer        = "smtp.gmail.com"
$vSmtpPort          = 587
$vEmailRemetente    = "TROCAR@gmail.com"
$vSenhaAppGmail     = "TROCAR"
$vEmailDestino      = "TROCAR", "TROCAR"

# =================================================================
# FUNÇÃO DE ENVIO DE E-MAIL HTML (TEMPLATE CUBOTIMIZE)
# =================================================================
Function Send-CubotimizeEmail {
    param (
        [string]$Status,
        [string]$Mensagem,
        [string]$Anexo = "",
        [string]$CorBadge = "#4A5568"
    )

    $vServicoNome = "Dump de Apps Qlik Sense - $vServidorNome"
    
    $vHtmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 0; background-color: #f0f2f5; }
    .container { max-width: 650px; margin: 20px auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 8px 20px rgba(0,0,0,0.1); border: 1px solid #e1e4e8; }
    .cube-strip { height: 6px; width: 100%; display: flex; }
    .strip-blue { background-color: #2E63E6; width: 33.3%; }
    .strip-green { background-color: #3BE854; width: 33.3%; }
    .strip-red { background-color: #E83B3B; width: 33.3%; }
    .header { background-color: #4A5567; padding: 30px 20px; text-align: center; background-image: linear-gradient(135deg, #3BE854 0%, #2E63E6 100%); }
    .header h1 { color: #ffffff; margin: 0; font-size: 26px; font-weight: 800; letter-spacing: 3px; text-transform: uppercase; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    .content { padding: 40px 30px; color: #333333; text-align: center; }
    .status-badge { display: inline-block; padding: 12px 28px; border-radius: 50px; font-weight: 900; font-size: 20px; color: #ffffff; background-color: $CorBadge; box-shadow: 0 4px 6px rgba(0,0,0,0.15); margin-bottom: 25px; min-width: 140px; text-transform: uppercase; letter-spacing: 1px; }
    .info-card { background-color: #f8f9fa; border-radius: 8px; padding: 25px; margin-top: 20px; text-align: left; border: 1px solid #eaeaea; }
    .label { font-size: 11px; color: #888; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 4px; display: block;}
    .value { font-size: 16px; color: #222; font-weight: 500; margin-bottom: 15px; word-break: break-all; }
    .value-link { color: #2E63E6; text-decoration: none; font-weight: bold; }
    .footer { background-color: #f8f9fa; padding: 25px; text-align: center; font-size: 12px; color: #999; border-top: 1px solid #eaeaea; }
    .slogan { color: #2E63E6; font-weight: 700; font-size: 13px; margin-bottom: 8px; display: block; }
  </style>
</head>
<body>
  <div class="container">
    <div class="cube-strip"><div class="strip-blue"></div><div class="strip-green"></div><div class="strip-red"></div></div>
    <div class="header"><h1>STATUS DO PROCESSO</h1></div>
    <div class="content">
        <div class="status-badge">$Status</div>
        <div style="font-size: 16px; color: #555; margin-bottom: 20px;">Notificação de alteração de estado.</div>
        <div class="info-card">
            <span class="label">SERVIÇO</span>
            <div class="value">$vServicoNome</div>
            <span class="label">ENDEREÇO DE DESTINO</span>
            <div class="value"><a href="$vPastaDestino" class="value-link">$vPastaDestino</a></div>
            
            <div style="margin-top: 25px; font-size: 14px; color: #333; border-top: 1px solid #ddd; padding-top: 20px;">
                $Mensagem
            </div>
        </div>
    </div>
    <div class="footer"><span class="slogan">Soluções em Inteligência Tecnológica</span>&copy; Cubotimize - Monitoramento</div>
  </div>
</body>
</html>
"@

    $vAssunto = "[Cubotimize] $Status - Dump Apps Qlik ($vServidorNome)"
    
    try {
        $vSecurePassword = ConvertTo-SecureString $vSenhaAppGmail -AsPlainText -Force
        $vCredenciais = New-Object System.Management.Automation.PSCredential ($vEmailRemetente, $vSecurePassword)

        $mailParams = @{
            From       = $vEmailRemetente
            To         = $vEmailDestino
            Subject    = $vAssunto
            Body       = $vHtmlBody
            BodyAsHtml = $true
            SmtpServer = $vSmtpServer
            Port       = $vSmtpPort
            UseSsl     = $true
            Credential = $vCredenciais
            Encoding   = [System.Text.Encoding]::UTF8
        }
        
        if ($Anexo -ne "" -and (Test-Path $Anexo)) { $mailParams.Add("Attachments", $Anexo) }

        Send-MailMessage @mailParams
        Write-Host "Notificação '$Status' enviada com sucesso!" -ForegroundColor Green
    } catch {
        Write-Host "Falha ao enviar e-mail: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# =================================================================
# SETUP DE DIRETÓRIOS E LOGS
# =================================================================
$vTempoInicioScript = Get-Date
$vDataAgora = Get-Date -Format "yyyy-MM-dd"
$vPastaDestino = "$($vPastaBackup)$($vDataAgora)\"

If (!(Test-Path $vPastaDestino)) {
    New-Item -ItemType Directory -Force -Path $vPastaDestino | Out-Null
}

# DISPARO 1: E-MAIL DE INÍCIO
if ($vEnviarEmail) {
    $vModoDados = if ($vSemDados) { "SEM DADOS (SkipData)" } else { "COM DADOS" }
    Send-CubotimizeEmail -Status "▶️ INICIADO" -Mensagem "<div style='background-color:#fff3cd; color:#856404; padding:15px; border-radius:6px; font-family:Consolas,monospace;'>O processo de dump de aplicativos (.qvf) começou.<br>Modo: <b>$vModoDados</b>. Aguarde o relatório de conclusão.</div>" -CorBadge "#2E63E6"
}

$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -path "$($vPastaDestino)backup.log" -append

Echo "================================================================="
Echo " Cubotimize - Backup de Aplicações Qlik Sense"
Echo " Versão: 2.4.0"
Echo " Versão Data: 16/03/2026"
Echo "================================================================="
Echo ""

# =================================================================
# CONEXÃO (CERTIFICADO)
# =================================================================
echo "Buscando certificado interno do Qlik Sense..."
$vCert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Subject -eq "CN=QlikClient" } | Select-Object -First 1
if (-not $vCert) {
    $vCert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -eq "CN=QlikClient" } | Select-Object -First 1
}

if ($vCert) {
    echo "Certificado encontrado! Conectando..."
    Connect-Qlik -ComputerName $vServidorQlik -TrustAllCerts -Certificate $vCert
} else {
    Write-Error "FALHA CRÍTICA: Certificado QlikClient não encontrado!"
    Stop-Transcript
    exit
}

# =================================================================
# CONTADORES PARA O RELATÓRIO
# =================================================================
$vCountPublicados = 0
$vCountTrabalho   = 0
$vCountErros      = 0
$vListaErros      = @()
$vContagemFluxo   = @{}
$vContagemPessoa  = @{}

echo "------Iniciando backup de Apps (Publicados e Área de Trabalho)-------"

$vCaracteresInvalidos = '[\\/:*?"<>|\[\]]'
$vTodosOsApps = Invoke-QlikGet -path "/qrs/app/full"
$vQtdTotalDeApps = $vTodosOsApps.Count

foreach ($vQvf in $vTodosOsApps) {
    
    $vNomePastaDestino = ""
    $vNomeLimpoParaTabela = ""

    if ($vQvf.published -and $vQvf.stream.name) {
        $vCountPublicados++
        $vNomeLimpoParaTabela = $vQvf.stream.name -replace $vCaracteresInvalidos, ""
        $vNomeLimpoParaTabela = $vNomeLimpoParaTabela.Trim()
        $vNomePastaDestino = $vNomeLimpoParaTabela
        
        # Soma +1 neste Fluxo
        $vContagemFluxo[$vNomeLimpoParaTabela] = [int]$vContagemFluxo[$vNomeLimpoParaTabela] + 1
    } 
    else {
        $vCountTrabalho++
        $vID_Usuario = $vQvf.owner.userId
        if ([string]::IsNullOrWhiteSpace($vID_Usuario)) { $vID_Usuario = $vQvf.owner.name }
        if ([string]::IsNullOrWhiteSpace($vID_Usuario)) { $vID_Usuario = [string]$vQvf.owner }
        
        if ([string]::IsNullOrWhiteSpace($vID_Usuario) -or $vID_Usuario -match "System.Object") { 
            $vID_Usuario = "SemDono" 
        } 
        
        $vID_Usuario = $vID_Usuario -replace $vCaracteresInvalidos, ""
        $vNomeLimpoParaTabela = $vID_Usuario.Trim()
        $vNomePastaDestino = "__Trabalho\$vNomeLimpoParaTabela"
        
        # Soma +1 nesta Pessoa
        $vContagemPessoa[$vNomeLimpoParaTabela] = [int]$vContagemPessoa[$vNomeLimpoParaTabela] + 1
    }

    echo "Carregando App: $($vQvf.name) | Destino: $vNomePastaDestino"

    $vCaminhoCompletoPasta = "$vPastaDestino$vNomePastaDestino"
    If (!(Test-Path $vCaminhoCompletoPasta)) {
        New-Item -ItemType Directory -Force -Path $vCaminhoCompletoPasta | Out-Null
    }
    
    $vAppFileName = $vQvf.name -replace $vCaracteresInvalidos, ""
    $vAppFileName = $vAppFileName.Trim()

    try {
        $vCaminhoArquivo = "$vCaminhoCompletoPasta\$vAppFileName.qvf"
        
        if ($vSemDados) {
            Export-QlikApp -id $vQvf.id -filename $vCaminhoArquivo -SkipData -ErrorAction Stop
        } else {
            Export-QlikApp -id $vQvf.id -filename $vCaminhoArquivo -ErrorAction Stop
        }
    }
    catch {
        $vCountErros++
        $vMsgErroFormatada = "<b>App:</b> $($vQvf.name) <br><b>Falha:</b> $($_.Exception.Message)"
        $vListaErros += $vMsgErroFormatada
        Write-Warning "FALHA ao exportar o App: '$($vQvf.name)' (ID: $($vQvf.id)). Erro: $($_.Exception.Message)"
    }
}

# =================================================================
# LIMPEZA DE RETENÇÃO D-60
# =================================================================
echo ""
echo "------Iniciando exclusão Backup antigo D-$($vDiasBackup)-------"

$vDataCorte = (Get-Date).AddDays(-$vDiasBackup)
$vPastas_Deletar = Get-ChildItem -Path $vPastaBackup -Directory | Where-Object { 
    $_.Name -match "^\d{4}-\d{2}-\d{2}$" -and $_.CreationTime -lt $vDataCorte
}

if ($vPastas_Deletar) {
    foreach ($vPasta in $vPastas_Deletar) {
        echo "Excluindo backup antigo: $($vPasta.FullName)"
        Get-ChildItem -Path $vPasta.FullName -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { if ($_.Attributes -match "ReadOnly") { $_.Attributes = "Normal" } }
        Get-ChildItem -Path $vPasta.FullName -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -Path $vPasta.FullName -Force -Recurse -ErrorAction Continue
    }
} else {
    echo "Nenhum backup com mais de $vDiasBackup dias encontrado para exclusão."
}

echo "------Finalizado-------"

$vTempoFimScript = Get-Date
$vDuracaoScript = $vTempoFimScript - $vTempoInicioScript

echo ""
Echo ">>>>>>>>>>> Tempo de execução: $([math]::Round($vDuracaoScript.TotalMinutes, 2)) minutos"
echo ""

Stop-Transcript

# =================================================================
# DISPARO 2: E-MAIL DE CONCLUÍDO (DINÂMICO)
# =================================================================
if ($vEnviarEmail) {
    $vDuracaoArredondada = [math]::Round($vDuracaoScript.TotalMinutes, 2)
    
    $vCorErro = if ($vCountErros -gt 0) { "#E83B3B" } else { "#3BE854" }
    $vModoDadosTexto = if ($vSemDados) { "SEM DADOS (SkipData)" } else { "COM DADOS" }

    # Lógica Dinâmica de Status Geral com Ícone de Alto Contraste
    $vStatusFinal = "✅ CONCLUÍDO COM SUCESSO"
    $vCorBadgeFinal = "#3BE854"

    if ($vCountErros -gt 0) {
        $vStatusFinal = "⚠️ CONCLUÍDO COM FALHAS/ALERTAS"
        $vCorBadgeFinal = "#E83B3B"
    }

    # Bloco 1: Resumo Geral
    $vHtmlRelatorio = @"
    <h3 style='color: #2E63E6; margin-bottom: 10px; font-size: 18px;'>📊 Resumo da Execução</h3>
    <table style='width: 100%; border-collapse: collapse; text-align: left; background-color: #fff; border: 1px solid #eee;'>
        <tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>Modo de Backup:</td><td style='padding: 8px; font-weight: bold; color: #4A5567; border-bottom: 1px solid #eee;'>$vModoDadosTexto</td></tr>
        <tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>Tempo de Retenção:</td><td style='padding: 8px; font-weight: bold; color: #4A5567; border-bottom: 1px solid #eee;'>$vDiasBackup dias</td></tr>
        <tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>Total de Apps no Servidor:</td><td style='padding: 8px; font-weight: bold; border-bottom: 1px solid #eee;'>$vQtdTotalDeApps</td></tr>
        <tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>Apps Publicados:</td><td style='padding: 8px; font-weight: bold; color: #2E63E6; border-bottom: 1px solid #eee;'>$vCountPublicados</td></tr>
        <tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>Apps em Área de Trabalho:</td><td style='padding: 8px; font-weight: bold; border-bottom: 1px solid #eee;'>$vCountTrabalho</td></tr>
        <tr><td style='padding: 8px;'>Falhas de Exportação:</td><td style='padding: 8px; font-weight: bold; color: $vCorErro;'>$vCountErros</td></tr>
    </table>
    <br>
"@

    # Bloco 2: Alertas (Só aparece se houver falhas)
    if ($vCountErros -gt 0) {
        $vHtmlRelatorio += @"
        <h3 style='color: #E83B3B; margin-bottom: 10px; font-size: 16px;'>⚠️ Alertas e Falhas</h3>
        <div style='background-color: #fce8e6; padding: 15px; border-radius: 6px; border-left: 4px solid #E83B3B; margin-bottom: 20px;'>
            <ul style='color: #d93025; margin: 0; padding-left: 20px; font-size: 13px; line-height: 1.6;'>
"@
        foreach ($erro in $vListaErros) { $vHtmlRelatorio += "<li style='margin-bottom: 8px;'>$erro</li>" }
        $vHtmlRelatorio += "</ul></div>"
    }

    # Bloco 3: Agregado por Fluxo
    $vHtmlRelatorio += @"
    <h3 style='color: #4A5567; margin-top: 15px; margin-bottom: 10px; font-size: 16px;'>📂 Quantidade por Fluxo (Publicados)</h3>
    <table style='width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; border: 1px solid #ddd;'>
        <tr style='background-color: #f1f3f4;'>
            <th style='padding: 8px; border: 1px solid #ddd;'>Nome do Fluxo</th>
            <th style='padding: 8px; border: 1px solid #ddd; width: 80px; text-align: center;'>Qtd Apps</th>
        </tr>
"@
    foreach ($key in ($vContagemFluxo.Keys | Sort-Object)) {
        $vHtmlRelatorio += "<tr><td style='padding: 6px 8px; border: 1px solid #ddd;'>$key</td><td style='padding: 6px 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;'>$($vContagemFluxo[$key])</td></tr>"
    }
    $vHtmlRelatorio += "</table><br>"

    # Bloco 4: Agregado por Pessoa
    $vHtmlRelatorio += @"
    <h3 style='color: #4A5567; margin-top: 15px; margin-bottom: 10px; font-size: 16px;'>👤 Quantidade por Usuário (Área de Trabalho)</h3>
    <table style='width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; border: 1px solid #ddd;'>
        <tr style='background-color: #f1f3f4;'>
            <th style='padding: 8px; border: 1px solid #ddd;'>Nome do Usuário</th>
            <th style='padding: 8px; border: 1px solid #ddd; width: 80px; text-align: center;'>Qtd Apps</th>
        </tr>
"@
    foreach ($key in ($vContagemPessoa.Keys | Sort-Object)) {
        $vHtmlRelatorio += "<tr><td style='padding: 6px 8px; border: 1px solid #ddd;'>$key</td><td style='padding: 6px 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;'>$($vContagemPessoa[$key])</td></tr>"
    }
    $vHtmlRelatorio += "</table><br>"

    $vTextoTopo = "O dump dos aplicativos foi concluído em <b>$vDuracaoArredondada minutos</b>. Segue abaixo o relatório executivo da extração:"
    $vMensagemFinal = $vTextoTopo + $vHtmlRelatorio
    
    Send-CubotimizeEmail -Status $vStatusFinal -Mensagem $vMensagemFinal -Anexo "$vPastaDestino\backup.log" -CorBadge $vCorBadgeFinal
}
