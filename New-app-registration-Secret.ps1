<#
New-app-registration.ps1
- Creates a default Entra ID App Registration (like in the portal)
- Creates the Service Principal (with tags so it shows under Enterprise Applications)
- Creates ONE client secret named: <service>-<workload>-<businessunit>-<environment>
- Appends CSV: DisplayName, ApplicationId, ServicePrincipalId, ClientSecret
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9-]+$')] [string]$Prefix,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9-]+$')] [string]$Service,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9-]+$')] [string]$Workload,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9]+$')]  [string]$BusinessUnit,
  [Parameter(Mandatory)][ValidateSet('prod','test','dev','stage')] [string]$Environment,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9-]+$')] [string]$Function,

  [string]$OutCsv = ".\appreg-output.csv",
  [ValidateRange(1,60)][int]$SecretExpiryInMonths = 12,
  [string]$TenantId
)

function Initialize-Graph {
  $modules = @('Microsoft.Graph.Authentication','Microsoft.Graph.Applications')
  foreach ($m in $modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
      Install-Module $m -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $m -Force -ErrorAction Stop
  }

  try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }

  $requiredScopes = @('Application.ReadWrite.All')
  if ($TenantId) {
    Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
  } else {
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
  }
}

function New-DisplayName {
  param([string]$Prefix,[string]$Service,[string]$Workload,[string]$BusinessUnit,[string]$Environment,[string]$Function)
  $name = ("{0}-{1}-{2}-{3}-{4}-{5}" -f $Prefix,$Service,$Workload,$BusinessUnit,$Environment,$Function).ToLower()
  if ($name.Length -gt 120) { throw "DisplayName too long ($($name.Length)). Keep it ≤120 chars." }
  if ($name -notmatch '^[a-z0-9-]+$') { throw "DisplayName contains invalid chars. Allowed: [a-z0-9-]. Name: $name" }
  return $name
}

function New-SecretName {
  param([string]$Service,[string]$Workload,[string]$BusinessUnit,[string]$Environment)
  $s = ("{0}-{1}-{2}-{3}" -f $Service,$Workload,$BusinessUnit,$Environment).ToLower()
  if ($s.Length -gt 64) { $s = $s.Substring(0,64) }
  return $s
}

# --- Main ---
Initialize-Graph

$displayName = New-DisplayName -Prefix $Prefix -Service $Service -Workload $Workload -BusinessUnit $BusinessUnit -Environment $Environment -Function $Function
$secretName  = New-SecretName  -Service $Service -Workload $Workload -BusinessUnit $BusinessUnit -Environment $Environment

# Guard against duplicates
$existing = Get-MgApplication -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
if ($existing) { throw "An application with displayName '$displayName' already exists (AppId: $($existing.AppId))." }

# 1) Create App Registration
$app = New-MgApplication -DisplayName $displayName -SignInAudience AzureADMyOrg -ErrorAction Stop
Write-Host "Application created. AppId: $($app.AppId)" -ForegroundColor Green

# 2) Create Service Principal WITH TAGS
$sp = New-MgServicePrincipal `
  -AppId $app.AppId `
  -Tags @('HideApp','WindowsAzureActiveDirectoryIntegratedApp') `
  -ErrorAction Stop
Write-Host "Service principal created. ObjectId: $($sp.Id)" -ForegroundColor Green

# Verify tags applied
$spCheck = Get-MgServicePrincipal -ServicePrincipalId $sp.Id
Write-Host "Tags on SP: $($spCheck.Tags -join ', ')" -ForegroundColor Cyan

# 3) Create Client Secret
$endIso = (Get-Date).ToUniversalTime().AddMonths($SecretExpiryInMonths).ToString("o")
$AppPassword = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
  displayName = $secretName
  endDateTime = $endIso
} -ErrorAction Stop

$secretValue = $AppPassword.SecretText
if (-not $secretValue) { throw "Client secret created, but no SecretText returned." }

# 4) Output to CSV
$row = [PSCustomObject]@{
  DisplayName        = $displayName
  ApplicationId      = $app.AppId
  ServicePrincipalId = $sp.Id
  ClientSecret       = $secretValue
}

if (Test-Path $OutCsv) {
  $row | Export-Csv -NoTypeInformation -Append -Path $OutCsv
} else {
  $row | Export-Csv -NoTypeInformation -Path $OutCsv
}

$row

Disconnect-MgGraph -ErrorAction SilentlyContinue

