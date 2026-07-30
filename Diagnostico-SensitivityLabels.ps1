#Requires -Version 7.0

<#
.SYNOPSIS
    Diagnóstico de Rótulos de Sensibilidade no SharePoint Online (Microsoft 365)

.DESCRIPTION
    Verifica se a integração de rótulos de sensibilidade (AIP/MIP) está habilitada
    no nível de tenant e diagnostica sites individuais de forma interativa.

.NOTES
    Módulos : Microsoft.Online.SharePoint.PowerShell
              Microsoft.Graph.Authentication (para listar labels via Graph /beta)
    Permissão necessária: SharePoint Administrator
              Escopo Graph: InformationProtectionPolicy.Read (delegado)
    Autenticação: Moderna (MFA via browser)

.NOTES
    O endpoint de labels do Microsoft Graph usado aqui
    (security/informationProtection/sensitivityLabels) ainda está em /beta
    e sujeito a mudanças pela Microsoft. Em modo delegado só há consulta
    via /me (não existe endpoint organizacional delegado).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────

function Write-Banner {
    $linhas = @(
        '╔══════════════════════════════════════════════════════════════════╗',
        '║   Diagnóstico de Rótulos de Sensibilidade – SharePoint Online   ║',
        '║   Microsoft 365                                                  ║',
        '╚══════════════════════════════════════════════════════════════════╝'
    )
    Write-Host ''
    foreach ($linha in $linhas) { Write-Host $linha -ForegroundColor Cyan }
    Write-Host ''
}

function Write-Section ([string]$Titulo) {
    Write-Host ''
    Write-Host ('─' * 64) -ForegroundColor DarkGray
    Write-Host "  $Titulo" -ForegroundColor Yellow
    Write-Host ('─' * 64) -ForegroundColor DarkGray
}

function Write-Ok    ([string]$Msg) { Write-Host "  [OK] $Msg" -ForegroundColor Green  }
function Write-Warn  ([string]$Msg) { Write-Host "  [!!] $Msg" -ForegroundColor Yellow }
function Write-Fail  ([string]$Msg) { Write-Host "  [XX] $Msg" -ForegroundColor Red    }
function Write-Info  ([string]$Msg) { Write-Host "       $Msg" -ForegroundColor Gray   }

# ─────────────────────────────────────────────────────────────────────────────
# VERIFICAÇÃO E INSTALAÇÃO DO MÓDULO
# ─────────────────────────────────────────────────────────────────────────────

function Assert-SPOModule {
    $moduleName = 'Microsoft.Online.SharePoint.PowerShell'

    Write-Section "Verificando módulo PowerShell"

    $modInstalado = Get-Module -ListAvailable -Name $moduleName |
                    Sort-Object Version -Descending |
                    Select-Object -First 1

    if (-not $modInstalado) {
        Write-Warn "Módulo '$moduleName' não encontrado."
        Write-Info  "Instalando no escopo do usuário atual..."

        try {
            Install-Module -Name $moduleName `
                           -Scope CurrentUser `
                           -Force `
                           -AllowClobber `
                           -Repository PSGallery `
                           -ErrorAction Stop

            $modInstalado = Get-Module -ListAvailable -Name $moduleName |
                            Sort-Object Version -Descending |
                            Select-Object -First 1

            Write-Ok "Módulo instalado (versão $($modInstalado.Version))."
        }
        catch {
            Write-Fail "Falha ao instalar o módulo: $_"
            exit 1
        }
    }
    else {
        Write-Ok "Módulo encontrado (versão $($modInstalado.Version))."
    }

    # Importar — em PS7 pode carregar via modo de compatibilidade automaticamente
    try {
        Import-Module $moduleName -ErrorAction Stop
    }
    catch {
        # Tenta o modo de compatibilidade Windows PowerShell (PS7 fallback)
        Write-Warn "Carregamento direto falhou. Tentando modo de compatibilidade..."
        Import-Module $moduleName -UseWindowsPowerShell -ErrorAction Stop
        Write-Ok "Módulo carregado via modo de compatibilidade."
    }
}

function Assert-GraphModule {
    $moduleName = 'Microsoft.Graph.Authentication'

    $modInstalado = Get-Module -ListAvailable -Name $moduleName |
                    Sort-Object Version -Descending |
                    Select-Object -First 1

    if (-not $modInstalado) {
        Write-Warn "Módulo '$moduleName' não encontrado."
        Write-Info  "Instalando no escopo do usuário atual..."
        try {
            Install-Module -Name $moduleName `
                           -Scope CurrentUser `
                           -Force `
                           -AllowClobber `
                           -Repository PSGallery `
                           -ErrorAction Stop
            $modInstalado = Get-Module -ListAvailable -Name $moduleName |
                            Sort-Object Version -Descending |
                            Select-Object -First 1
            Write-Ok "Módulo instalado (versão $($modInstalado.Version))."
        }
        catch {
            Write-Fail "Falha ao instalar o módulo: $_"
            exit 1
        }
    }
    else {
        Write-Ok "Módulo '$moduleName' encontrado (versão $($modInstalado.Version))."
    }

    Import-Module $moduleName -ErrorAction Stop
}

# ─────────────────────────────────────────────────────────────────────────────
# CONEXÃO AO SHAREPOINT ONLINE
# ─────────────────────────────────────────────────────────────────────────────

function Connect-AoSPO {
    Write-Section "Conexão ao SharePoint Online"

    $tenantNome = (Read-Host "`n  Nome do tenant (ex: contoso)").Trim().ToLower()
    $adminUrl   = "https://$tenantNome-admin.sharepoint.com"

    Write-Info "Conectando a: $adminUrl"
    Write-Info 'Abrindo autenticação moderna no browser (MFA habilitado)...'

    try {
        Connect-SPOService -Url $adminUrl -ErrorAction Stop
        Write-Ok "Conectado com sucesso."
        return $adminUrl
    }
    catch {
        Write-Fail "Falha na conexão: $_"
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONEXÃO AO MICROSOFT GRAPH E LISTAGEM DE LABELS
# ─────────────────────────────────────────────────────────────────────────────

function Connect-AoGraph {
    Write-Section "Conexão ao Microsoft Graph"
    Write-Info 'Abrindo autenticação moderna no browser (MFA habilitado)...'

    try {
        Connect-MgGraph -Scopes 'InformationProtectionPolicy.Read' -NoWelcome -ErrorAction Stop
        Write-Ok 'Conectado ao Microsoft Graph.'
    }
    catch {
        Write-Fail "Falha na conexão ao Microsoft Graph: $_"
        exit 1
    }
}

function Get-LabelsDoTenant {
    Write-Section "Rótulos de Sensibilidade definidos (Microsoft Graph)"

    # Endpoint em /beta (sujeito a mudanças). Em modo delegado só existe consulta
    # via /me — não há endpoint organizacional disponível para permissões delegadas.
    $uri = 'https://graph.microsoft.com/beta/me/security/informationProtection/sensitivityLabels'

    try {
        $resposta = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    }
    catch {
        Write-Warn "Não foi possível listar os rótulos via Graph: $_"
        return @{}
    }

    $labels = $resposta.value

    if (-not $labels -or $labels.Count -eq 0) {
        Write-Warn 'Nenhum rótulo de sensibilidade encontrado para o usuário atual.'
        return @{}
    }

    # Monta hashtable GUID(lower) → Nome para uso na resolução de sites
    $mapa = @{}

    Write-Host ''
    Write-Host "  {'Prioridade',-4}  {'Nome',-40}  {'Escopo'}" -ForegroundColor White
    Write-Host "  $('-' * 4)  $('-' * 40)  $('-' * 30)" -ForegroundColor DarkGray

    foreach ($label in ($labels | Sort-Object sensitivity)) {
        $guid       = $label.id.ToString().ToLower()
        $nome       = $label.name
        $escopo     = ($label.contentFormats -join ', ')
        $prioridade = $label.sensitivity

        $mapa[$guid] = $nome

        $cor     = if ($label.isActive) { 'Green' } else { 'DarkYellow' }
        $inativo = if (-not $label.isActive) { '  [inativo]' } else { '' }
        Write-Host ("  {0,-6}  {1,-40}  {2}{3}" -f $prioridade, $nome, $escopo, $inativo) -ForegroundColor $cor
    }

    Write-Host ''
    Write-Ok "$($labels.Count) rótulo(s) encontrado(s)."

    return $mapa
}

# ─────────────────────────────────────────────────────────────────────────────
# DIAGNÓSTICO — NÍVEL TENANT
# ─────────────────────────────────────────────────────────────────────────────

function Test-TenantLabels {
    Write-Section "Diagnóstico — Nível de Tenant"

    try {
        $tenant = Get-SPOTenant -ErrorAction Stop
    }
    catch {
        Write-Fail "Não foi possível obter as configurações do tenant: $_"
        Disconnect-SPOService -ErrorAction SilentlyContinue
        exit 1
    }

    $aipAtivo = $tenant.EnableAIPIntegration

    Write-Host ''

    if ($aipAtivo) {
        Write-Ok  "EnableAIPIntegration = TRUE"
        Write-Info "A integração com rótulos de sensibilidade (MIP/AIP) está ATIVA no SharePoint."
    }
    else {
        Write-Fail "EnableAIPIntegration = FALSE"
        Write-Info "A integração com rótulos de sensibilidade está DESATIVADA no tenant SPO."
    }

    Write-Host ''
    Write-Host '  ┌─ NOTA sobre Entra ID ──────────────────────────────────────┐' -ForegroundColor DarkCyan
    Write-Host '  │ A configuração "EnableMIPLabels" no Entra ID (necessária    │' -ForegroundColor DarkCyan
    Write-Host '  │ para rótulos em Grupos Microsoft 365) deve ser verificada   │' -ForegroundColor DarkCyan
    Write-Host '  │ manualmente:                                                │' -ForegroundColor DarkCyan
    Write-Host '  │ Portal Entra ID > Grupos > Configurações > EnableMIPLabels  │' -ForegroundColor DarkCyan
    Write-Host '  └────────────────────────────────────────────────────────────┘' -ForegroundColor DarkCyan

    return $aipAtivo
}

# ─────────────────────────────────────────────────────────────────────────────
# DIAGNÓSTICO — SITE INDIVIDUAL
# ─────────────────────────────────────────────────────────────────────────────

function Get-SiteLabelInfo ([string]$SiteUrl, [hashtable]$LabelsMap = @{}) {
    try {
        $site = Get-SPOSite -Identity $SiteUrl -ErrorAction Stop

        # SensitivityLabel retorna um Guid; vazio/zeros = sem rótulo
        $labelGuid = $site.SensitivityLabel
        $semLabel  = [string]::IsNullOrWhiteSpace($labelGuid) -or
                     $labelGuid -eq '00000000-0000-0000-0000-000000000000'

        # Resolve o nome: primeiro pelo mapa Purview (mais confiável), depois pela propriedade SPO
        $labelNome = $null
        if (-not $semLabel) {
            $labelNome = $LabelsMap[$labelGuid.ToString().ToLower()]
        }
        if (-not $labelNome) {
            try { $labelNome = $site.SensitivityLabelName } catch { }
        }

        return [PSCustomObject]@{
            Url       = $site.Url
            Titulo    = $site.Title
            Template  = $site.Template
            LabelGuid = if (-not $semLabel) { $labelGuid } else { $null }
            LabelNome = if (-not $semLabel) { $labelNome } else { $null }
            ComLabel  = -not $semLabel
            Erro      = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Url       = $SiteUrl
            Titulo    = $null
            Template  = $null
            LabelGuid = $null
            LabelNome = $null
            ComLabel  = $false
            Erro      = $_.Exception.Message
        }
    }
}

function Show-SiteResult ([PSCustomObject]$Resultado) {
    Write-Host ''

    if ($Resultado.Erro) {
        Write-Fail "Erro ao consultar o site:"
        Write-Info  "URL  : $($Resultado.Url)"
        Write-Info  "Msg  : $($Resultado.Erro)"
        Write-Info  "Verifique se a URL está correta e se você tem permissão."
        return
    }

    Write-Host "  URL      : $($Resultado.Url)"     -ForegroundColor White
    Write-Host "  Título   : $($Resultado.Titulo)"  -ForegroundColor White
    Write-Host "  Template : $($Resultado.Template)" -ForegroundColor White

    if ($Resultado.ComLabel) {
        $nomeDisplay = if ($Resultado.LabelNome) { $Resultado.LabelNome } else { '(nome não disponível)' }
        Write-Host "  Rótulo   : $nomeDisplay" -ForegroundColor Green
        Write-Host "  GUID     : $($Resultado.LabelGuid)" -ForegroundColor DarkGreen
        Write-Ok   "Rótulo de sensibilidade APLICADO neste site."
    }
    else {
        Write-Host "  Rótulo   : (nenhum)" -ForegroundColor Yellow
        Write-Warn "Nenhum rótulo de sensibilidade aplicado neste site."
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# LOOP INTERATIVO DE SITES
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SiteLoop ([hashtable]$LabelsMap = @{}) {
    Write-Section "Diagnóstico — Sites Individuais"

    Write-Host @"

  Informe a URL completa de cada site para diagnóstico.
  Pressione ENTER sem digitar nada para encerrar.

  Exemplos:
    https://contoso.sharepoint.com/sites/MeuSite
    https://contoso.sharepoint.com/teams/MinhaEquipe

"@ -ForegroundColor Gray

    $resultados = [System.Collections.Generic.List[PSCustomObject]]::new()

    while ($true) {
        $entrada = (Read-Host '  URL do site').Trim()

        if ([string]::IsNullOrEmpty($entrada)) {
            break
        }

        if ($entrada -notmatch '^https://') {
            Write-Fail "URL inválida — deve começar com 'https://'"
            continue
        }

        $resultado = Get-SiteLabelInfo -SiteUrl $entrada -LabelsMap $LabelsMap
        Show-SiteResult -Resultado $resultado
        $resultados.Add($resultado)
    }

    return $resultados
}

# ─────────────────────────────────────────────────────────────────────────────
# RESUMO FINAL
# ─────────────────────────────────────────────────────────────────────────────

function Show-Resumo ([System.Collections.Generic.List[PSCustomObject]]$Resultados) {
    if ($Resultados.Count -eq 0) { return }

    Write-Section "Resumo dos Sites Verificados"
    Write-Host ''

    $Resultados | ForEach-Object {
        $status = if ($_.Erro)     { 'ERRO'        }
                  elseif ($_.ComLabel) { 'Com rótulo'  }
                  else              { 'Sem rótulo'  }

        $cor = switch ($status) {
            'Com rótulo'  { 'Green'  }
            'Sem rótulo'  { 'Yellow' }
            default        { 'Red'    }
        }

        $nome = if ($_.LabelNome) { $_.LabelNome }
                elseif ($_.LabelGuid) { $_.LabelGuid }
                else { '-' }

        $linha = "  [{0,-12}]  {1,-50}  {2}" -f $status, $_.Url, $nome
        Write-Host $linha -ForegroundColor $cor
    }

    $total    = $Resultados.Count
    $comLabel = ($Resultados | Where-Object { $_.ComLabel }).Count
    $semLabel = ($Resultados | Where-Object { -not $_.ComLabel -and -not $_.Erro }).Count
    $erros    = ($Resultados | Where-Object { $_.Erro }).Count

    Write-Host ''
    Write-Host "  Total verificado : $total" -ForegroundColor White
    Write-Host "  Com rótulo       : $comLabel" -ForegroundColor Green
    Write-Host "  Sem rótulo       : $semLabel" -ForegroundColor Yellow
    Write-Host "  Com erro         : $erros"    -ForegroundColor Red
}

# ─────────────────────────────────────────────────────────────────────────────
# EXECUÇÃO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

try {
    Write-Banner
    Assert-SPOModule
    Assert-GraphModule
    Connect-AoSPO | Out-Null

    $tenantHabilitado = Test-TenantLabels

    if (-not $tenantHabilitado) {
        Write-Host ''
        Write-Warn "Diagnóstico de sites encerrado — tenant não está habilitado."
        Write-Info  "Habilite com:  Set-SPOTenant -EnableAIPIntegration `$true"
        Write-Host ''
    }
    else {
        Connect-AoGraph
        $labelsMap = Get-LabelsDoTenant

        $resultados = Invoke-SiteLoop -LabelsMap $labelsMap
        Show-Resumo -Resultados $resultados
    }
}
finally {
    Write-Host ''
    Write-Host '  Desconectando...' -ForegroundColor Gray
    Disconnect-SPOService -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host '  Sessão encerrada.' -ForegroundColor Cyan
    Write-Host ''
}
