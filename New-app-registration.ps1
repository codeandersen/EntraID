<#
New-app-registration.ps1
- Creates a default Entra ID App Registration (like in the portal)
- Creates the Service Principal
- Creates ONE client secret named: <service>-<workload>-<businessunit>-<environment>
- Appends CSV: DisplayName, ApplicationId, ServicePrincipalId, ClientSecret

Example:
.\New-app-registration.ps1 `
  -Prefix app `
  -Service acs `
  -Workload schilowsky `
  -BusinessUnit starkat `
  -Environment prod `
  -Function smtp-credentials `
  -OutCsv C:\Github\EntraID\appreg-csv.csv
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

  # Optional: force a specific tenant if you work across multiple
  [string]$TenantId
)

function Initialize-Graph {

  # Import only the lightweight Graph pieces we need (avoid meta-module overflow)
  $modules = @('Microsoft.Graph.Authentication','Microsoft.Graph.Applications')
  foreach ($m in $modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
      Install-Module $m -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $m -Force -ErrorAction Stop
  }

  # Force disconnect any existing sessions to ensure clean state
  try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "Cleared existing Graph sessions" -ForegroundColor Yellow
  } catch { }

  # Connect 
  $requiredScopes = @('Application.ReadWrite.All')
  Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
  
  try {
    if ($TenantId) {
      Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
    } 
    
    # Verify connection worked
    $testContext = Get-MgContext -ErrorAction Stop
    Write-Host "Successfully connected to Graph. Account: $($testContext.Account)" -ForegroundColor Green
    Write-Host "Available scopes: $($testContext.Scopes -join ', ')" -ForegroundColor Cyan
    

  } catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    throw "Unable to establish Graph connection. Please check your credentials and network connectivity."
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
  if ($s.Length -gt 64) { Write-Warning "Secret name is long ($($s.Length)). Truncating to 64 chars."; $s = $s.Substring(0,64) }
  return $s
}

# --- Main ---
Initialize-Graph

$displayName = New-DisplayName -Prefix $Prefix -Service $Service -Workload $Workload -BusinessUnit $BusinessUnit -Environment $Environment -Function $Function
$secretName  = New-SecretName  -Service $Service -Workload $Workload -BusinessUnit $BusinessUnit -Environment $Environment

# Guard against duplicate displayName

Write-Host "Checking for existing application with displayName: $displayName" -ForegroundColor Yellow
try {
  $existing = Get-MgApplication -Filter "displayName eq '$displayName'" -ErrorAction Stop
  if ($existing) {
    throw "An application with displayName '$displayName' already exists (AppId: $($existing.AppId))."
  }
  Write-Host "No existing application found. Proceeding..." -ForegroundColor Green
} catch {
  Write-Error "Failed to check for existing applications: $($_.Exception.Message)"
  throw
}

# 1) Create default App Registration (portal-like defaults)
Write-Host "Creating application registration..." -ForegroundColor Yellow
$app = New-MgApplication -DisplayName $displayName -SignInAudience AzureADMyOrg -ErrorAction Stop
Write-Host "Application created successfully. AppId: $($app.AppId)" -ForegroundColor Green

# 2) Create Service Principal
Write-Host "Creating service principal..." -ForegroundColor Yellow
$sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
Write-Host "Service principal created successfully. ObjectId: $($sp.Id)" -ForegroundColor Green

# 3) Create Client Secret
Write-Host "Creating client secret..." -ForegroundColor Yellow
$endIso = (Get-Date).ToUniversalTime().AddMonths($SecretExpiryInMonths).ToString("o")
$AppPassword = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
  displayName = $secretName
  endDateTime = $endIso
} -ErrorAction Stop

$secretValue = $AppPassword.SecretText
if (-not $secretValue) { 
  throw "Client secret created, but no SecretText returned. Check Graph SDK/scopes." 
}
Write-Host "Client secret created successfully" -ForegroundColor Green

# 4) CSV output (plaintext secret – handle securely)
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

Write-Host "Disconnecting from Graph sessions" -ForegroundColor Yellow
Disconnect-MgGraph -ErrorAction SilentlyContinue
    