<#
    .SYNOPSIS
        This script loads Azure VMs into a Sharepoint Online list. 

    .DESCRIPTION
        A SharePoint Online custom list is required with at least the following columns to load the Virtual Machines found in
        the provided subscription.
            - VMName {Renamed Title column} - Single line of text
                This will be the Name of your VM in Azure
	        - SubscriptionGUID - 'Single line of text'
                The Azure Subscription GUID where the VM is situated
            - ResourceGroup - 'Single line of text'
                The Resource Group Name of the VM
	        - TimeZone - 'Choice'
                A drop down list of the available Timezones for the shutdown times to be run against
                Run the following to generate the choice options.
                ForEach ($TZ in ([system.timezoneinfo]::GetSystemTimeZones())) {Write-Output "$($TZ.DisplayName) [$($TZ.Id)]"}
    .MODULES
        AzureRM 3.6.0
        SharePointSDK 2.1.0

    .INPUTS
        $InParam_SubscriptionID
            The Azure Subscription GUID to operate against
        $SPOURL
            The URL of your Sharepoint online site i.e. https://xxxx.sharepoint.com
        $SPOListName
            The name of the SharePoint List to add VM details to

    .OUTPUTS
        Records to specified SharePoint Online list
        General information for progress of script to Host Output
#>

Login-AzureRmAccount

$InParam_SubscriptionID = Read-Host "Enter Subscription ID"

$AvailableSubs = Get-AzureRmSubscription
if ($InParam_SubscriptionID -notin $AvailableSubs.SubscriptionId) {
    Write-Error "Access to selected Subscription is denied"
} else {
    Select-AzureRmSubscription -SubscriptionID $InParam_SubscriptionID 
}

$SPOCreds = Get-Credential -Message "Enter Sharepoint Online Credentials"

$SPOURL = Read-Host "Enter SharePoint Online Site URL i.e. https://xxxx.sharepoint.com"
$SPOListName = Read-Host "Enter the name of the SharePoint List to add VM details to"

$SPOVMs = Get-SPListItem -SiteUrl $SPOURL -Credential $SPOCreds -IsSharePointOnlineSite $true -ListName $SPOListName 
#Remove deleted VMs
$SPOVMs = $SPOVMs | Where-Object {$_.DateRemoved -le "1 Jan 2000"}

$VMs = Get-AzureRmVM -Status
Write-Output "Found $($VMs.Count) VMs"
$SPIDs = @()
ForEach ($VM in $VMs) {
    $SPORec = $null
    $SPORec = $SPOVMs | Where-Object {($_.Title -eq $($VM.Name)) -and ($_.SubscriptionGUID -eq $InParam_SubscriptionID)}
    
    $SPData = @{
        Title = $VM.Name
        SubscriptionGUID = $Subscrip.SubscriptionId
        ResourceGroup = $VM.ResourceGroupName
        TimeZone = "(UTC+10:00) Canberra, Melbourne, Sydney [AUS Eastern Standard Time]"
    }
    
    if ($SPORec -eq $null) {
        Write-Output "Added $($VM.Name)"
        $SPIDs += Add-SPListItem -SPConnection $SPOConn -ListName $SPOListName -ListFieldsValues $SPData
    }
}

$SubRecs = $SPOVMs | Where-Object {$_.SubscriptionGUID -eq $InParam_SubscriptionID}
foreach ($SPOVM in $SPOVMs) {
    $VMIndex = $VMs.name.IndexOf($SPOVM.Title)
    if ($VMIndex -eq -1) {
        $SPData = @{
            Title = $SPOVM.Title
            SubscriptionGUID = $InParam_SubscriptionID
            DateRemoved = Get-Date -Format 'dd MMM yyyy HH:mm'
        }
        Write-Output "Deleted $($SPOVM.Title)"
        $SPOUpd = Update-SPListItem -SPConnection $SPOConn -ListName $SPOListName -ListItemId $SPOVM.Id -ListFieldsValues $SPData
        if ($SPOUpd) {
            Write-Output "Successfully marked $($SPOVM.Title) as deleted"
        } else {
            Write-Output "Failed to update Sharepoint for $($SPOVM.Title)"
        }
    }
}

