param([string]$AdminEmail,
  [string]$TenantName)

$FilePath = Read-Host "Please enter site template XML schema file path";
$FilePath = $FilePath.Trim();

Write-Host $FilePath

$TenantName = Read-Host "Please enter your tenant name: (contoso)";
$TenantName = $TenantName.Trim();

$AdminEmail = Read-Host "Please enter your tenant admin email";
$AdminEmail = $AdminEmail.Trim();

# verify the PnP.PowerShell module we need is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell )) {
  Write-Warning "Could not find the PnP.PowerShell module, installing it"
  Install-Module -Name "PnP.PowerShell"
} else {
  Write-Host "PnP.PowerShell module found"
}


$TenantURL = "https://$TenantName.sharepoint.com"
#Do not change
$EOCSiteURL = "/sites/TEOCSite"

Connect-PnPOnline -Url $TenantURL -Interactive

try {
    Write-Host "Checking if site exists at $EOCSiteURL"
    $site = Get-PnPTenantSite -Url $TenantURL$EOCSiteURL -ErrorAction SilentlyContinue

    if ($site -ne $null)
    {
        Write-Host "TEOC Site exists. Updating the site template..."
        
    }
    else
    {
        Write-Host "TEOC Site doesn't exist. Update aborted."
        return;
    }
    

    Connect-PnPOnline -Url $TenantURL$EOCSiteURL -Interactive

    
    Invoke-PnPSiteTemplate -Path $FilePath -ErrorAction Stop -WarningAction SilentlyContinue
          
    Write-Host "TEOC App Upgrade complete."
}
catch{

    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nTEOC App Upgrade failed."
}