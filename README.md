Load Azure VMs into a SharePoint Online List
============================================


  *  
SYNOPSIS
This script loads Azure VMs into a Sharepoint Online list.


  *  
DESCRIPTION
A SharePoint Online custom list is required with at least the following columns to load the Virtual Machines found in the provided subscription.


  *  
VMName {Renamed Title column} - Single line of text

This will be the Name of your VM in Azure

  *  
SubscriptionGUID - 'Single line of text' 
The Azure Subscription GUID where the VM is situated

  *  
ResourceGroup - 'Single line of text' 
The Resource Group Name of the VM

  *  
TimeZone - 'Choice' 
A drop down list of the available Timezones for the shutdown times to be run against
Run the following to generate the choice options. 
ForEach ($TZ in ([system.timezoneinfo]::GetSystemTimeZones())) {Write-Output '$($TZ.DisplayName) [$($TZ.Id)]'}





  *  
MODULES

  *  
AzureRM 3.6.0

  *  
SharePointSDK 2.1.0


  *  
INPUTS

  *  
$InParam_SubscriptionID 
The Azure Subscription GUID to operate against

  *  
$SPOURL 
The URL of your Sharepoint online site i.e. [https://xxxx.sharepoint.com](https://xxxx.sharepoint.com)

  *  
$SPOListName 
The name of the SharePoint List to add VM details to OUTPUTS Records to specified SharePoint Online list General information for progress of script to Host Output

See this post for an Azure Automation Runbook version -> Load
 Azure VMs into a SharePoint Online List - Automation Runbook



 


        
    
TechNet gallery is retiring! This script was migrated from TechNet script center to GitHub by Microsoft Azure Automation product group. All the Script Center fields like Rating, RatingCount and DownloadCount have been carried over to Github as-is for the migrated scripts only. Note : The Script Center fields will not be applicable for the new repositories created in Github & hence those fields will not show up for new Github repositories.
