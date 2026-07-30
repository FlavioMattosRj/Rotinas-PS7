#Requires -Version 7.0

<#
.SYNOPSIS
    Diagnóstico de Rótulos de Sensibilidade no SharePoint Online (Microsoft 365)

.DESCRIPTION
    Verifica se a integração de rótulos de sensibilidade (AIP/MIP) está habilitada
    no nível de tenant e diagnostica sites individuais de forma interativa.

.NOTES
    Módulos : Microsoft.Online.SharePoint.PowerShell
              ExchangeOnlineManagement (para listar labels via Purview)
    Permissão necessária: SharePoint Administrator + Compliance Administrator
    Autenticação: Moderna (MFA via browser)
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

function Assert-EXOModule {
    $moduleName = 'ExchangeOnlineManagement'

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
# CONEXÃO AO PURVIEW (COMPLIANCE) E LISTAGEM DE LABELS
# ─────────────────────────────────────────────────────────────────────────────

function Connect-AoPurview ([string]$TenantNome) {
    Write-Section "Conexão ao Purview / Compliance (ExchangeOnlineManagement)"
    Write-Info 'Conectando via Connect-IPPSSession (MFA habilitado)...'

    try {
        # UserPrincipalName é opcional; se omitido, o browser pede credenciais
        Connect-IPPSSession -ShowBanner:$false -ErrorAction Stop
        Write-Ok 'Conectado ao centro de Conformidade/Purview.'
    }
    catch {
        Write-Fail "Falha na conexão ao Purview: $_"
        exit 1
    }
}

function Get-LabelsDoTenant {
    Write-Section "Rótulos de Sensibilidade definidos no Tenant (Purview)"

    try {
        $labels = Get-Label -ErrorAction Stop
    }
    catch {
        Write-Warn "Não foi possível listar os rótulos: $_"
        return @{}
    }

    if ($labels.Count -eq 0) {
        Write-Warn 'Nenhum rótulo de sensibilidade publicado encontrado no tenant.'
        return @{}
    }

    # Monta hashtable GUID(lower) → DisplayName para uso na resolução de sites
    $mapa = @{}

    Write-Host ''
    Write-Host "  {'Prioridade',-4}  {'Nome',-40}  {'Escopo'}" -ForegroundColor White
    Write-Host "  $('-' * 4)  $('-' * 40)  $('-' * 30)" -ForegroundColor DarkGray

    foreach ($label in ($labels | Sort-Object Priority)) {
        $guid    = $label.ImmutableId.ToString().ToLower()
        $nome    = $label.DisplayName
        $escopo  = $label.ContentType -join ', '
        $prioridade = $label.Priority

        $mapa[$guid] = $nome

        $cor = if ($label.IsActive) { 'Green' } else { 'DarkYellow' }
        $inativo = if (-not $label.IsActive) { '  [inativo]' } else { '' }
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
    Assert-EXOModule
    Connect-AoSPO | Out-Null

    $tenantHabilitado = Test-TenantLabels

    if (-not $tenantHabilitado) {
        Write-Host ''
        Write-Warn "Diagnóstico de sites encerrado — tenant não está habilitado."
        Write-Info  "Habilite com:  Set-SPOTenant -EnableAIPIntegration `$true"
        Write-Host ''
    }
    else {
        Connect-AoPurview
        $labelsMap = Get-LabelsDoTenant

        $resultados = Invoke-SiteLoop -LabelsMap $labelsMap
        Show-Resumo -Resultados $resultados
    }
}
finally {
    Write-Host ''
    Write-Host '  Desconectando...' -ForegroundColor Gray
    Disconnect-SPOService -ErrorAction SilentlyContinue
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host '  Sessão encerrada.' -ForegroundColor Cyan
    Write-Host ''
}
