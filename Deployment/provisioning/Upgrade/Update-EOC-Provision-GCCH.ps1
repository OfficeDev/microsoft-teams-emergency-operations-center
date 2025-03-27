param([string]$TenantName)

$FilePath = Read-Host "Please enter site template XML schema file path";
$FilePath = $FilePath.Trim();

Write-Host $FilePath

$TenantName = Read-Host "Please enter your tenant name: (contoso)";
$TenantName = $TenantName.Trim();

$SiteName = Read-Host "Enter your existing TEOC site name";
$SiteName = $SiteName.Trim();

$ClientID = Read-Host "Enter AzureAppId/ClientId";
$ClientID = $ClientID.Trim();

# verify the PnP.PowerShell module we need is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell )) {
  Write-Warning "Could not find the PnP.PowerShell module, installing it"
  Install-Module -Name "PnP.PowerShell"
}
else {
  Write-Host "PnP.PowerShell module found"
}

$TenantURL = "https://$TenantName.sharepoint.us"
$EOCSiteURL = "/sites/$SiteName"
$Tenant = "$TenantName.onmicrosoft.us"

Connect-PnPOnline -Url $TenantURL -Interactive -AzureEnvironment USGovernmentHigh -ClientId $ClientID -Tenant $Tenant

try {
  Write-Host "Checking if site exists at $EOCSiteURL"
  $site = Get-PnPTenantSite -Url $TenantURL$EOCSiteURL -ErrorAction SilentlyContinue

  if ($null -ne $site) {
    Write-Host "TEOC Site exists. Updating the site template..."        
  }
  else {
    Write-Host "TEOC Site doesn't exist. Update aborted."
    return;
  }

  Connect-PnPOnline -Url $TenantURL$EOCSiteURL -Interactive -ClientId $ClientID
    
  Invoke-PnPSiteTemplate -Path $FilePath -ErrorAction Stop -WarningAction SilentlyContinue
          
  Write-Host "TEOC App Upgrade complete."
  pause
}
catch {
  Write-Host "`nError Message: " $_.Exception.Message
  Write-Host "`nTEOC App Upgrade failed."
  pause
}