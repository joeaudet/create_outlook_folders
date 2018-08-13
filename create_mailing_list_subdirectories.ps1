$comments = @'
Script name: Create-Folder.ps1
Created on: Wednesday, September 26, 2007
Author: Kent Finkle
Original Purpose: How can I use Windows Powershell to Create a New Folder in Microsoft Outlook?
Modified: 2018AUG010 Joe Audet
Current Purpose: Create a subfolder structure for mailing lists, ignoring folders that are already created
'@
#-----------------------------------------------------
#Set the Inbox folder
$olFolderInbox = 6

#Load the Outlook COM object and set the namespace
$o = new-object -comobject outlook.application
$n = $o.GetNamespace("MAPI")

#Get the Inbox and store it as an object
$inbox = $n.GetDefaultFolder($olFolderInbox)

#Initial revision uses an array of folder names to create subfolders under the 'Mailing Lists' folder within the Inbox
#Example: $mailing_lists_array = @("SUB_FOLDER_1","SUB_FOLDER_2")
$mailing_lists_array = @("15K_23K_APPLIANCES","61K_SUPPORT","ADVANCED_ROUTING","API_TECH","APP_TECH","BRIDGE_TECH","CLOUD_GW","COMPETITIVE-ANALYSIS","DDOS_TECH","DEMOPOINT","DLP_TECH","DOCUMENTSECURITY","EA_UPDATES","ENDPOINTSECURITY","ENDPOINT_CONNECT","EVENTIA","GAIA_TECH","ICS_TECH","IDACCESS","IPS","IPV6_TECH","MOBILE-ENTERPRISE","MSP_TECH","MTP","MTP_CUSTOMER_UPDATES","MTP_SALES","POC_TECH","R80_MGMT","SBCloudO365","SECACCEL_TECH","SE_INSTRUCTOR","SMB_TECH","SOCIAL_MEDIA","SSLVPN","US_TECH","VE","VOIP_TECH","VPN_S2S","VSEC","VSXSUPPORT","PRICE_LIST_UPDATES","NA_SALES_SE_TEAM")

#Check if the 'Mailing Lists' folder exists - If yes, send a message and do nothing, if not create the folder
if ([bool]($mailing_list_folder = $inbox.Folders | where-object { $_.name -eq "Mailing Lists" }) ) {
	Write-Host "Folder: $($inbox.name)\$($mailing_list_folder.name) Already Exists - Skipping"
} else {
	$newfolder = $inbox.Folders.Add("Mailing Lists")
	Write-Host "$($newfolder.name) Created"
	$mailing_list_folder = $inbox.Folders | where-object { $_.name -eq "Mailing Lists" }
}

#Function to iterate through the array, checking to see if each folder exists - if so send a message and do nothing, if not create folder
function create_subfolders {
	foreach ($array_folder in $mailing_lists_array) {
		if ( [bool]( $targetfolder = $mailing_list_folder.Folders | where-object { $_.name -eq $array_folder } ) ) {
			Write-Host "Folder: $($array_folder) exists under $($inbox.name)\$($mailing_list_folder.name) - Skipping"
		} else {
			$newfolder = $mailing_list_folder.Folders.Add($array_folder)
			Write-Host "Created folder: $($array_folder) under $($inbox.name)\$($mailing_list_folder.name)"
		}
	}
}

#Call function to create folders
create_subfolders

#Send message to let user know folders were created
Write-Host "Create / Update Mailing List folder script complete"
