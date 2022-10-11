param([string]$AdminEmail,
  [string]$TenantName)

$FilePath = Read-Host "Enter site template XML schema file path";
$FilePath = $FilePath.Trim();

Write-Host $FilePath

$TenantName = Read-Host "Enter tenant name: (contoso)";
$TenantName = $TenantName.Trim();

$AdminEmail = Read-Host "Enter tenant admin email";
$AdminEmail = $AdminEmail.Trim();

$SiteName = Read-Host "Enter site name. Allowed characters for site name are underscore, dashes, single quotes, and periods (_,-,',.), and can't start or end with a period.";
$SiteName = $SiteName.Trim();

$SiteURL = $SiteName -replace " ",""
# verify the PnP.PowerShell module we need is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell )) {
  Write-Warning "Could not find the PnP.PowerShell module, installing it"
  Install-Module -Name "PnP.PowerShell"
} else {
  Write-Host "PnP.PowerShell module found"
}

$TenantURL = "https://$TenantName.sharepoint.com"
$EOCSiteURL = "/sites/$SiteURL"

Connect-PnPOnline -Url $TenantURL -Interactive

try {
    Write-Host "Checking if site already exists at $EOCSiteURL"
    $site = Get-PnPTenantSite -Url $TenantURL$EOCSiteURL -ErrorAction SilentlyContinue

    if ($site -ne $null)
    {
        Write-Host "Site already exists, exiting the PowerShell script"
        return;
    }
    else
    {
        Write-Host "Site doesn't exist, creating new site at $EOCSiteURL"
    }

    try{
        if(($TenantURL+$EOCSiteURL).Length -lt 128){
            New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $SiteName -Url $TenantURL$EOCSiteURL -Owner $AdminEmail -ErrorAction Stop -WarningAction SilentlyContinue
        }
        else{
            Write-Host "Site creation failed. Site URL cannot have more than 128 characters." -ErrorAction Stop
            return
        }
    }
    catch{
        Write-Host "Site creation failed. Site name cannot contain symbols other than underscore, dashes, single quotes, and periods (_,-,',.), and can't start or end with a period."
        return;

    }
    Connect-PnPOnline -Url $TenantURL$EOCSiteURL -Interactive

    Write-Host "Creating lists in $SiteName site"

    Invoke-PnPSiteTemplate -Path $FilePath -ErrorAction Stop -WarningAction SilentlyContinue
          
    Write-Host "App Provision complete."
}
catch{
    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nApp Provisioning failed."
}
