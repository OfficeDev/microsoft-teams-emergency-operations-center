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
	Write-Host "Checking if site already exists at $EOCSiteURL"
	$site = Get-PnPTenantSite -Url $TenantURL$EOCSiteURL -ErrorAction SilentlyContinue
}
catch{
	
}

if ($site -ne $null)
{
    Write-Host "Site already exists, exiting the PowerShell script"
	return;
}
else
{
    Write-Host "Site doesn't exist, creating new site at $EOCSiteURL"
}

New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title TEOC -Url $TenantURL$EOCSiteURL -Owner $AdminEmail

Connect-PnPOnline -Url $TenantURL$EOCSiteURL -Interactive

Write-Host "Creating lists in the EOC site"

Invoke-PnPSiteTemplate -Path $FilePath

Write-Host "EOC App Provision complete."

