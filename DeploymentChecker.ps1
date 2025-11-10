#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '05/13/2024 8:32:09 PM'.

# Site configuration
$SiteCode = "TRP" # Site code 
$ProviderMachineName = "opcs00620.corp.troweprice.net" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

add-type -assemblyName System.Windows.forms

#run from ISE so that the console window remains open so you can view the results after selecting the deployment(s), hold CTRL to select multiple or SHIFT to select a contiguous group

$deployments=get-cmdeployment | where-object {$_.deploymentTime -gt (get-date) -or $_.EnforcementDeadline -gt (Get-date) } | Select ApplicationName,CollectionName,DeploymentTime,EnforcementDeadline,FeatureType | sort-object EnforcementDeadline -Descending
# Passthru enables the script to continue executing after the grid-view is closed or OK is clicked
# assigning a variable to the output we push to it enables us to select either one or multiple objects that we can then retrieve via the variable
if (!($deployments)) { [System.Windows.Forms.MessageBox]::Show("No Future deployments found, script will end");[System.Environment]::Exit(0)}
$selected=$deployments | out-gridview -passthru

foreach ($select in $selected)
{
    switch ($($Select.FeatureType)) {
    '1' { $type="Application" }
    '2' { $type="Package" }
    '5' { $Type="Software Update" }
    
    }
    # I could probably just re-use $var to find the match since my above query already has the data but for a quick test I wasn't thinking too clearly and just re-looked up per each selected
    get-cmdeployment -CollectionName "$($select.CollectionName)" -SoftwareName "$($select.applicationName)" -featureType $select.FeatureType | where-object { $_.deploymentTime -gt (get-date) -or $_.EnforcementDeadline -gt (get-date) } | select ApplicationName,DeploymentID,collectionID,CollectionName,DeploymentTime,EnforcementDeadline,PackageID, FeatureType
    
}
