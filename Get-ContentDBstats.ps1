                                         #####################    Content Database Status Script    #####################
                                         
# This script gets the status of the Content Databases and related info on the current farm and saves the result in a CSV file on UKFIL812WIN
# This script is executed via a scheduled task

### Variables and Arguments -------------------------------------------------------------------------------------------------------------------------------

$FormattedDate = $(Get-Date).ToString("dd-MM-yyyy")
$HostServer = hostname

# Path where the CSV file is saved on the remote server
$SavePath = "\\UKFIL812WIN\D$\Scripts\Nizar\SharePoint Health\Content Databases Reports\ContentDatabases $FormattedDate.csv"
#----------------------------------------------------------------------------------------------------------------------------------------------------------

### Functions ---------------------------------------------------------------------------------------------------------------------------------------------

### Function to format size in GB 
function FormatBytes ($bytes){ 
    switch ($bytes) 
    { 
        {$bytes -ge 1TB} {"{0:n$sigDigits}" -f ($bytes/1TB) + " TB" ; break} 
        {$bytes -ge 1GB} {"{0:n$sigDigits}" -f ($bytes/1GB) + " GB" ; break} 
        {$bytes -ge 1MB} {"{0:n$sigDigits}" -f ($bytes/1MB) + " MB" ; break} 
        {$bytes -ge 1KB} {"{0:n$sigDigits}" -f ($bytes/1KB) + " KB" ; break} 
        Default { "{0:n$sigDigits}" -f $bytes + " Bytes" } 
    } 
}

### Function to get Content Database Info
function Get-ContentDBInfo(){ 
# Load SharePoint assemblies	
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
	
# Get the Content Service from SharePoint Web Services
$WebService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService

# Get the propertie types for the content database
$DBFarm  = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("Farm")
$DBServer= [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("ServiceInstance")
$DBName = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("Name") 
$DBStatus  = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("Status")
$DBSiteCount = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("CurrentSiteCount")
$DBWarningSiteCount = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("WarningSiteCount")
$DBMaximumSiteCount = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("MaximumSiteCount")
$DBDiskSize  = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("DiskSizeRequired")
$DBIsReadOnly = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("IsReadOnly")
$DBNeedsUpgrade  = [Microsoft.SharePoint.Administration.SPContentDatabase].GetProperty("NeedsUpgrade")

        # Enumerate through all Web applications in the farm -------------------------------------------
            	foreach($WebApplication in $WebService.WebApplications){

            		$ContentDBCollection = $WebApplication.ContentDatabases

            		$webAppName = $WebApplication.name	
            		
            		# Enumerate through all content databases attached to the Web application
            		foreach($ContentDB in $ContentDBCollection){
            			
            			# Add variables here with any new properties and then add them to the output lines
                        $CurrentDBFarm = $DBFarm.GetValue($ContentDB, $null)
                        $CurrentDBFarm = [string]$CurrentDBFarm
                        $CurrentDBFarm = $CurrentDBFarm.split("=")[1]
                        $CurrentDBServer = ($DBServer.GetValue($ContentDB, $null)).NormalizedDataSource
            			$CurrentDBName = $DBName.GetValue($ContentDB, $null)
                        $CurrentDBStatus = $DBStatus.GetValue($ContentDB, $null)
            			$CurrentDBCurrentSiteCount = $DBSiteCount.GetValue($ContentDB, $null)
                        $CurrentDBDWarningSiteCount = $DBWarningSiteCount.GetValue($ContentDB, $null)
            			$CurrentDBDBMaximumSiteCount = $DBMaximumSiteCount.GetValue($ContentDB, $null)
            			$DiskSize = $DBDiskSize.GetValue($ContentDB, $null)
                        $DiskSizeGB = [Math]::Round($DiskSize/1GB, 2)
                        $DiskSizeFormatted = FormatBytes $DiskSize
            			$CurrentDBIsReadOnly = $DBIsReadOnly.GetValue($ContentDB, $null)
                        $CurrentDBNeedsUpgrade = $DBNeedsUpgrade.GetValue($ContentDB, $null)
                        
                        # Update CSV file	
            	        "$HostServer,$CurrentDBFarm,$webAppName,$CurrentDBServer,$CurrentDBName,$CurrentDBStatus,$CurrentDBCurrentSiteCount,$CurrentDBDWarningSiteCount,$CurrentDBDBMaximumSiteCount,$DiskSizeGB,$DiskSizeFormatted,$CurrentDBIsReadOnly,$CurrentDBNeedsUpgrade" | out-file "$SavePath" -append	
            	   }
               }
        #-----------------------------------------------------------------------------------------------	
}        
#----------------------------------------------------------------------------------------------------------------------------------------------------------


#---------------------------------------------------------------- MAIN SCRIPT -----------------------------------------------------------------------------

# Check if CSV file already exists, if not then create the header again	
    if (Test-Path $SavePath){
       Get-ContentDBInfo
    } 
    Else{
       # Create CSV file header
       "HostServer,Farm,webAppName,DBServer,DBName,DBStatus,DBCurrentSiteCount,DBDWarningSiteCount,DBDBMaximumSiteCount,DiskSizeGB,DiskSizeFormatted,DBIsReadOnly,NeedsUpgrade" | out-file "$SavePath" -append	
       Get-ContentDBInfo
    }       
#---------------------------------------------------------------- MAIN SCRIPT END -------------------------------------------------------------------------
