param([string]$SiteURL)

$SiteURL = Read-Host "Enter the TEOC Site URL: (https://<tenantname>.sharepoint.com/sites/<sitename>)";
$SiteURL = $SiteURL.Trim();

Write-Host $SiteURL

#Parameters
$ParentListName = "TEOC-IncidentStatus"
$ChildListName = "TEOC-IncidentTransaction"
$LookupColumnName = "Status"
 
try {
    #Connect to SharePoint Online site
    Connect-PnPOnline $SiteURL -Interactive

    #Get all items from TEOC-IncidentStatus list
    $items = Get-PnPListItem -List $ParentListName

    $hash = $null
    $hash = @{}

    #Store all the items in hash array
    foreach ($item in $items) {
        $hash.add($item.FieldValues.Title, $item.FieldValues.ID)
    }

    #Get all items from TEOC-IncidentTransaction list
    $listItems = Get-PnPListItem -List $ChildListName
 
    #Update the lookup column for all items in TEOC-IncidentTransaction list
    foreach ($listItem in $listItems) {
        try {
            $status = $listItem.FieldValues.IncidentStatus
            $value = $hash.$status

            $lookupColValue = $listItem.FieldValues.Status.LookupValue

            if (($null -eq $lookupColValue)) {
                #Update List Item Lookup column with the value from Status column
                Set-PnPListItem -List $ChildListName -Identity $listItem.Id -Values @{$LookupColumnName = $value } | Out-Null
                Write-host  "Updated Item -" $listItem.ID
            }
            else {
                Write-host  "Skipped Item -" $listItem.ID " since the lookup column already has a value"
            }
        }
        catch {
            Write-Host "`nError Message: " $_.Exception.Message
            continue
        }
    }
}
catch {
    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nFailed to update the lookup column."
}
