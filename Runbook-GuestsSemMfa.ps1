<#
.SYNOPSIS
    Bloqueia convidados sem MFA após X horas e envia relatório (Excel ou CSV) por e-mail via Microsoft Graph.

.DESCRIPTION
    - Conecta ao Microsoft Graph usando Managed Identity (Azure Automation).
    - Identifica usuários Guest ativos (accountEnabled = true) membros de um grupo específico,
      criados há >= janela definida.
    - Verifica métodos de autenticação individualmente via REST API para identificar quem está sem MFA.
    - Opcionalmente ignora convidados com convite pendente (ExternalUserState <> Accepted).
    - Bloqueia (accountEnabled = false), gera relatório (.xlsx com ImportExcel, ou .csv como fallback)
      e envia via Microsoft Graph (Invoke-MgGraphRequest).

.PARAMETER HoursWithoutMfa
    Janela em horas (padrão 24).

.PARAMETER SkipPending
    Ignora convidados com convite pendente (padrão: $false).

.PARAMETER GuestGroupId
    Object ID do grupo do Entra ID contendo os usuários Guest a serem avaliados.

.PARAMETER SenderUpn
    UPN da mailbox remetente (usuário ou caixa compartilhada).

.PARAMETER To
    Destinatário(s) do relatório (separar múltiplos por ',' ou ';').

.PARAMETER Cc
    (Opcional) Cópias (separar múltiplos por ',' ou ';').

.PARAMETER Subject
    Assunto do e-mail.

.PARAMETER DryRun
    Simulação sem aplicar bloqueio (padrão: $false).

.EXAMPLE
    .\Runbook-GuestsSemMfa.ps1 -HoursWithoutMfa 24 -SkipPending $true `
       -GuestGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
       -SenderUpn "noreply@contoso.com" -To "secops@contoso.com;gestao@contoso.com" `
       -Subject "Guests bloqueados - diário" -DryRun $false

.NOTES
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                                                                              ║
    ║  Autor    : Erick Medeiros - MVP - Microsoft Azure                           ║
    ║  Versão   : 2.0                                                              ║
    ║  Data     : 2026-06-06                                                       ║
    ║  Licença  : MIT                                                              ║
    ║                                                                              ║
    ╚══════════════════════════════════════════════════════════════════════════════╝

    REQUISITOS
    ----------
    - Azure Automation Account com System-assigned Managed Identity habilitada
    - Runtime PowerShell 5.1
    - Módulos: Microsoft.Graph.Authentication, Microsoft.Graph.Users,
               Microsoft.Graph.Groups, Microsoft.Graph.Reports,
               Microsoft.Graph.Users.Actions
    - Módulo opcional: ImportExcel (para gerar XLSX em vez de CSV)

    PERMISSÕES GRAPH (Application) na Managed Identity
    --------------------------------------------------
    - GroupMember.Read.All
    - User.ReadWrite.All
    - UserAuthenticationMethod.Read.All
    - AuditLog.Read.All
    - Mail.Send

    CHANGELOG
    ---------
    v2.0 (2026-06-06) - Erick Medeiros
      [BREAKING] Filtro de Guests alterado de sufixo de UPN para membership de grupo
                 (suporta grupos dinâmicos e atribuídos).
      [BREAKING] Parâmetros [switch] $SkipPending e $DryRun migrados para [bool]
                 — necessário para compatibilidade com o Test Pane do Azure Automation.
      [ADD]      Novo parâmetro $GuestGroupId (obrigatório).
      [ADD]      Parametrização de emails (SenderUpn, To, Cc) removendo hardcoded.
      [ADD]      Adicionado módulo Microsoft.Graph.Groups às dependências.
      [CHANGE]   Verificação de MFA agora é per-user via Invoke-MgGraphRequest direto
                 ao endpoint /users/{id}/authentication/methods, em substituição ao
                 relatório agregado Get-MgReportAuthenticationMethodUserRegistrationDetail
                 (que não cobria todos os Guests de forma confiável).
      [CHANGE]   Envio de e-mail migrado de Send-MgUserMail para Invoke-MgGraphRequest
                 com ConvertTo-Json -Depth 20 para evitar problemas de serialização
                 de hashtables aninhadas.
      [ADD]      Output detalhado linha-a-linha indicando "SEM MFA" ou "COM MFA" por
                 usuário, facilitando troubleshooting.
      [ADD]      Permissão UserAuthenticationMethod.Read.All adicionada à lista
                 obrigatória da Managed Identity.

    v1.0 - Versão inicial (modelo base)
      - Filtro de Guests por sufixo de UPN (hardcoded).
      - Verificação de MFA via relatório agregado.
      - Envio de e-mail via Send-MgUserMail (SDK cmdlet).
      - Destinatários e remetente hardcoded no script.

    REPOSITÓRIO
    -----------
    Para documentação completa de implantação, consulte os arquivos:
      - README.md
      - IMPLANTACAO.md
#>

param(
    [int]    $HoursWithoutMfa = 24,
    [bool]   $SkipPending      = $false,
    [string] $GuestGroupId     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",   # <<< Object ID do grupo de Guests no Entra ID
    [string] $SenderUpn        = "usuario.remetente@dominio.com.br",        # <<< UPN da caixa remetente
    [string] $To               = "destino1@dominio.com.br, destino2@dominio.com.br", # <<< Destinatários do relatório
    [string] $Cc,
    [string] $Subject          = "Relatório - Guests bloqueados sem MFA",
    [bool]   $DryRun           = $false
)

# Tornar erros "terminais" para facilitar troubleshooting em Runbook
#$ErrorActionPreference = 'Stop'

Write-Output "=== Início do runbook ($(Get-Date)) ==="
Write-Output "Janela sem MFA: $HoursWithoutMfa h | SkipPending: $SkipPending | DryRun: $DryRun"

# ---------------------------------------------------------------------
# Módulos (Graph + ImportExcel se disponível)
# ---------------------------------------------------------------------
$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Users.Actions'
)

#foreach ($m in $graphModules) {
#    try { Import-Module $m -ErrorAction Stop } catch {
#        throw "Módulo obrigatório não encontrado: $m. Garanta que está importado no Automation Account."
#    }
#}

# ImportExcel é opcional. Se não existir, usamos CSV.
$UseCsv = $false
try {
    if (Get-Module -ListAvailable -Name 'ImportExcel') {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Output "ImportExcel disponível: relatório será gerado em .xlsx."
    } else {
        $UseCsv = $true
        Write-Output "ImportExcel NÃO disponível: relatório será gerado em .csv (fallback)."
    }
} catch {
    $UseCsv = $true
    Write-Output "Falha ao carregar ImportExcel: relatório será gerado em .csv (fallback)."
}

# ---------------------------------------------------------------------
# Conexão Graph (Managed Identity)
# ---------------------------------------------------------------------
Write-Output "Conectando ao Microsoft Graph com Managed Identity..."
Connect-MgGraph -Identity | Out-Null
# Opcional: garantir perfil v1.0
try { Select-MgProfile -Name 'v1.0' } catch { }

# ---------------------------------------------------------------------
# Janela de corte
# ---------------------------------------------------------------------
$cutoff = (Get-Date).AddHours(-1 * $HoursWithoutMfa)
Write-Output "Data/hora de corte: $cutoff"

# Propriedades que usaremos
$selectProps = @('Id','DisplayName','UserPrincipalName','CreatedDateTime','AccountEnabled','ExternalUserState')

# ---------------------------------------------------------------------
# 1) Obter membros do grupo de Guests especificado
# ---------------------------------------------------------------------
Write-Output "Listando convidados (Guest) do tenant..."

$selectProps = @(
  "id",
  "displayName",
  "userPrincipalName",
  "mail",
  "createdDateTime",
  "accountEnabled",
  "externalUserState"
)

Write-Output "Buscando membros do grupo '$GuestGroupId'..."

# Busca todos os membros do grupo (suporta grupos dinâmicos e atribuídos)
# Requer GroupMember.Read.All
$groupMembers = Get-MgGroupMember -GroupId $GuestGroupId -All

# Busca os detalhes completos apenas dos membros que são usuários (filtra ServicePrincipals, etc.)
$guests = foreach ($member in $groupMembers) {
    if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
        Get-MgUser -UserId $member.Id -Property $selectProps |
            Select-Object $selectProps
    }
}

# Filtrar habilitados e criados antes do cutoff
$guests = $guests | Where-Object { $_.CreatedDateTime -le $cutoff -and $_.AccountEnabled -eq $true }

if ($SkipPending) {
    # ExternalUserState: 'Accepted' vs 'PendingAcceptance'
    $guests = $guests | Where-Object { $_.ExternalUserState -eq 'Accepted' }
    Write-Output "Ignorando convidados com convite pendente (somente Accepted)."
}

Write-Output "Total de convidados considerados após filtro: $($guests.Count)"

# ---------------------------------------------------------------------
# 2) Verificar MFA individualmente para cada Guest
# ---------------------------------------------------------------------
Write-Output "Verificando métodos de autenticação de cada convidado..."
# Usa Invoke-MgGraphRequest (REST direto) — não depende de sub-módulo extra

# Tipos de método que NÃO contam como MFA (apenas senha = sem MFA)
$nonMfaMethods = @(
    '#microsoft.graph.passwordAuthenticationMethod'
)

$targets = New-Object System.Collections.Generic.List[object]

foreach ($guest in $guests) {
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$($guest.Id)/authentication/methods"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $methods = $response.value

        # Filtra os tipos de método: se só tiver senha (ou nenhum), não tem MFA
        $mfaMethods = $methods | Where-Object {
            $_.'@odata.type' -notin $nonMfaMethods
        }
        if ($null -eq $mfaMethods -or @($mfaMethods).Count -eq 0) {
            $targets.Add($guest)
            Write-Output "  SEM MFA: $($guest.UserPrincipalName)"
        } else {
            Write-Output "  COM MFA: $($guest.UserPrincipalName) ($(@($mfaMethods).Count) método(s))"
        }
    } catch {
        Write-Output "  ERRO ao consultar métodos de $($guest.UserPrincipalName): $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------
# 3) Resumo dos candidatos a bloqueio
# ---------------------------------------------------------------------
Write-Output "Candidatos a bloqueio (sem MFA e dentro da janela): $($targets.Count)"

# ---------------------------------------------------------------------
# 4) Aplicar bloqueio (se não for DryRun)
# ---------------------------------------------------------------------
$blocked = New-Object System.Collections.Generic.List[object]

foreach ($u in $targets) {
    $row = [PSCustomObject]@{
        DisplayName        = $u.DisplayName
        UserPrincipalName  = $u.UserPrincipalName
        CreatedDateTime    = $u.CreatedDateTime
        ExternalUserState  = $u.ExternalUserState
        Action             = $DryRun ? 'Simulado' : 'Bloqueado'
        ActionDateTimeUtc  = (Get-Date).ToUniversalTime()
    }

    if (-not $DryRun) {
        try {
            # Requer User.EnableDisableAccount.All (ou User.ReadWrite.All)
            Update-MgUser -UserId $u.Id -BodyParameter @{ accountEnabled = $false }
            $blocked.Add($row)
        } catch {
            $row | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception.Message
            $blocked.Add($row)
            Write-Output "Falha ao bloquear $($u.UserPrincipalName): $($_.Exception.Message)"
        }
    } else {
        $blocked.Add($row)
    }
}

Write-Output "Total processado: $($targets.Count) | Bloqueados/Simulados: $($blocked.Count)"

# ---------------------------------------------------------------------
# 5) Gerar relatório (Excel se possível; caso contrário, CSV)
# ---------------------------------------------------------------------
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if ($UseCsv) {
    $reportPath = Join-Path $env:TEMP ("GuestsBloqueados_{0}.csv" -f $ts)
    $blocked | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
    $contentType = 'text/csv'
    Write-Output "Relatório gerado (CSV): $reportPath"
} else {
    $reportPath = Join-Path $env:TEMP ("GuestsBloqueados_{0}.xlsx" -f $ts)
    $blocked | Export-Excel -Path $reportPath -WorksheetName 'GuestsBloqueados' -AutoSize -FreezeTopRow -BoldTopRow
    $contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    Write-Output "Relatório gerado (Excel): $reportPath"
}

# ---------------------------------------------------------------------
# 6) Enviar e-mail com anexo via Graph (Send-MgUserMail)
# ---------------------------------------------------------------------

function ConvertTo-RecipientObjects {
    param([string]$Addresses)
    $arr = @()
    if ([string]::IsNullOrWhiteSpace($Addresses)) { return $arr }
    $split = $Addresses -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($addr in $split) { $arr += @{ emailAddress = @{ address = $addr } } }
    return ,$arr
}

$toRecipients = ConvertTo-RecipientObjects -Addresses $To
$ccRecipients = ConvertTo-RecipientObjects -Addresses $Cc

if ($SenderUpn -and $toRecipients.Count -gt 0) {
    Write-Output "Preparando e-mail para $($toRecipients.Count) destinatário(s)..."

$bodyHtml = @"
<p>Olá,</p>
<p>Segue em anexo o relatório de <b>Guests bloqueados sem MFA</b> na última janela de $HoursWithoutMfa horas.</p>
<p>Total de afetados: <b>$($blocked.Count)</b></p>
<p>Atenciosamente,<br/>Automação Entra ID</p>
"@

    $fileBytes     = [System.IO.File]::ReadAllBytes($reportPath)
    $base64Content = [System.Convert]::ToBase64String($fileBytes)

    $mail = @{
        message = @{
            subject      = $Subject
            body         = @{ contentType = 'HTML'; content = $bodyHtml }
            toRecipients = $toRecipients
            attachments  = @(
                @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name          = [System.IO.Path]::GetFileName($reportPath)
                    contentBytes  = $base64Content
                    contentType   = $contentType
                }
            )
        }
        saveToSentItems = $true
    }

    if ($ccRecipients.Count -gt 0) {
        $mail.message['ccRecipients'] = $ccRecipients
        Write-Output "CC incluído: $($ccRecipients.Count) destinatário(s)."
    }

    # Envio como /users/{SenderUpn}/sendMail (app-only, requer Mail.Send)
    $sendUri  = "https://graph.microsoft.com/v1.0/users/$SenderUpn/sendMail"
    $mailJson = $mail | ConvertTo-Json -Depth 20
    Invoke-MgGraphRequest -Method POST -Uri $sendUri -Body $mailJson -ContentType "application/json" | Out-Null
    Write-Output "E-mail enviado a partir de '$SenderUpn'."
} else {
    Write-Output "E-mail NÃO enviado: verifique SenderUpn e To."
}

# ---------------------------------------------------------------------
# Encerrar sessão Graph
# ---------------------------------------------------------------------
Disconnect-MgGraph | Out-Null
Write-Output "=== Fim do runbook ($(Get-Date)) ==="
