#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '01/24/2024 4:14:52 PM'.

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

# started with 30 days but stuff can sit in UAT for longer period so bumping up to 60 days
# 3/5/24 KLB - Just realized I need to subtract NumberUnknown from TotalTargeted to make the errorpct a little more accurate
# Requires 'configurationmanager' module.  Machine has sccm installed or have the .psd1 module file in the psmodule path.
$st=(get-date).addDays(-45)
$st1=(get-date).addDays(-1)

# right now not adding calculated field for errors for the deployments done for previous day as its intent is just to know what was deployed basically in the last 24 hours

$result=get-cmdeployment | where-object { $_.DeploymentTime -ge $st -and $_.NumberTargeted -gt 0 -and $_.NumberErrors -gt 0} | where-object { $_.NumberErrors/($_.NumberTargeted - $_.NumberUnknown) -gt .05} | where-object { $_.FeatureType -ne 5 -and $_.FeatureType -ne 6 } | select @{Name="ErrorPct";Expression={$pct=(($_.numberErrors/($_.NumberTargeted - $_.NumberUnknown))*100);'{0:N}' -f $pct}},ApplicationName,CollectionName,NumberTargeted,NumberSuccess,NumberErrors,NumberUnknown,NumberInProgress,DeploymentTime,FeatureType | sort-object @{e={[int]$_.ErrorPct}} -desc
$result1=get-cmdeployment | where-object { $_.DeploymentTime -gt $st1 -and $_.DeploymentTime -lt (get-date)} | select ApplicationName,CollectionName,NumberTargeted,NumberSuccess,NumberErrors,NumberUnknown,NumberInProgress,DeploymentTime,FeatureType | sort-object ApplicationName

$outlook=new-object -comobject Outlook.Application
$mail=$outlook.CreateItem(0)
$mail.to="TechOps.Core.Packaging@troweprice.com;Chris.Koutek@troweprice.com"
$mail.subject="Daily Deployments >5% error rate (Automated)"
# needing to format as strict string rather than object to properly display in body of email for some reason
$Sresult=$result | out-string
$mail.body=">5% Error Rates`nDeployments started within last 45 days:`nSince $st`nWe are excluding 3rd Party software updates [Patch my PC data]`nAlso Excluding Configuration and baseline deployments`nErrorPCT calculation=(NumberErrors/NumberTargeted - NumberUnknown) * 100`nAnything with > 5% error rate included`n=======================================================================`n$sResult`n=======================================================================================`n`n1=application`n2=package`n5=Software Updates`n6=Configuration Item`n7=Task Sequences`n11=Firewall Setting"
$mail.send()

$outlook=new-object -comobject Outlook.Application
$mail=$outlook.CreateItem(0)
$mail.to="TechOps.Core.Packaging@troweprice.com;Chris.Koutek@troweprice.com"
$mail.subject="Yesterday/Today Deployments (Automated)"
$S1result=$result1 | out-string
$mail.body="Deployments that started yesterday after $st1`n`n$S1result"
$mail.send()



