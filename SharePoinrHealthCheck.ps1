
                                         #####################    SharePoint Health Script    #####################

### Add Ins -----------------------------------------------------------------------------------------------------------------------------------------------
Add-PsSnapin Microsoft.SharePoint.PowerShell
Start-SPAssignment -global
#----------------------------------------------------------------------------------------------------------------------------------------------------------

### Variables and Arguments -------------------------------------------------------------------------------------------------------------------------------

# SharePoint site and list
$SPsite = "http://"
$SPlist = "SharePoint Health"

# Send Mail Configuration
$users = "nizar.ghribi@mail.com" # List of users to email your report to (separate by comma)
$fromemail = "SharePointHealth@mail.com"
$server = "mailgateway.com" 

# Path of the CSV file with servers list
$ServersList = "SPServersList.csv" 
$MonitoredServices = "ServicesToMonitor.csv"

# Variables intialization
$StartTime = Get-Date
$FormatedDate = $StartTime.ToString("dd/MM/yyyy")
$ListOfAttachments = @()
$countErrors = 0
$countWarnings = 0
#----------------------------------------------------------------------------------------------------------------------------------------------------------

### Functions ---------------------------------------------------------------------------------------------------------------------------------------------

# Function to create pie chart
Function Create-PieChart() {
       param([string]$FileName)
             
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
      
       #Create our chart object
       $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
       $Chart.Width = 40
       $Chart.Height = 40
       $Chart.Left = 0
       $Chart.Top = 0
       $Chart.BackColor = [System.Drawing.Color]::Transparent
       $Chart.Palette = "None"
       $Chart.PaletteCustomColors = "#98CA44", "#007CC1"

       #Create a chart area to draw on and add this to the chart
       $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
       $Chart.ChartAreas.Add($ChartArea)
       [void]$Chart.Series.Add("Data")

       #Add a datapoint for each value specified in the arguments (args)
       foreach ($value in $args[0]) {
           $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $value)
           $Chart.Series["Data"].Points.Add($datapoint)
       }

       $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::doughnut
       $Chart.Series["Data"]["PieLabelStyle"] = "Inside"
       $Chart.Series["Data"]["PieLineColor"] = "Black"

       #Save the chart to a file
       $Chart.SaveImage($FileName + ".png","png")
}

# Function to get for how long the server has been running
Function Get-HostUptime ($ComputerName){
       $Uptime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
       $LastBootUpTime = $Uptime.ConvertToDateTime($Uptime.LastBootUpTime)
       $Time = (Get-Date) - $LastBootUpTime
       Return '{0:00} Days, {1:00} Hours, {2:00} Minutes, {3:00} Seconds' -f $Time.Days, $Time.Hours, $Time.Minutes, $Time.Seconds
}

# Function to get logical disk info
Function Get-DiskInfo ($ComputerName){
      $DiskInfo = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName | Where-Object{$_.DriveType -eq 3} | Select-Object SystemName, DriveType, VolumeName, Name, @{n='SizeGB';e={"{0:n2}" -f ($_.size/1gb)}}, @{n='FreeSpaceGB';e={"{0:n2}" -f ($_.freespace/1gb)}}, @{n='PercentFree';e={"{0:n2}" -f ($_.freespace/$_.size*100)}}
      return $DiskInfo
}

Function AddAttachment($item, $filePath)
{
      $bytes = [System.IO.File]::ReadAllBytes($filePath)
      $item.Attachments.Add([System.IO.Path]::GetFileName($filePath), $bytes)
}
#----------------------------------------------------------------------------------------------------------------------------------------------------------


#---------------------------------------------------------------- MAIN SCRIPT -----------------------------------------------------------------------------

# Open site and list to store health stats
$Web = Get-SPWeb $SPsite
$List = $Web.Lists[$SPlist]

# Update List Descrition
$List.Description = "Script is now running... (Started @ $StartTime)"
$List.Update()

# Start loop on all items in the CSV file
Import-csv -path $ServersList | foreach-object { 

                $ProcessStart = Get-Date

        # Affect each entry from the CSV file to a variable --------------------------------------------
                $Server = $_.Server
                $Envirenment = $_.Envirenment
                $SP_Version = $_.SP_Version
                $ServerRole = $_.Role
        Write-Host $Server $Envirenment $SP_Version $ServerRole -ForegroundColor Blue
        #-----------------------------------------------------------------------------------------------
        
        # System Uptime --------------------------------------------------------------------------------
                $SystemUptime = Get-HostUptime $Server
        #-----------------------------------------------------------------------------------------------
        
        # Disk Info ------------------------------------------------------------------------------------ 
                $DiskInfo = Get-DiskInfo $Server
                $DiskInfo | foreach-object {
                $SystemName = $_.SystemName
                $VolumeName = $_.VolumeName
                $Name = $_.Name
                $Name2 = $Name.split(":")[0]
                $SizeGB = $_.SizeGB
                $FreeSpaceGB = $_.FreeSpaceGB
                $UsedSpaceGB = $SizeGB - $FreeSpaceGB
                $PercentFree = $_.PercentFree
                    Create-PieChart -FileName ((Get-Location).Path + "\chart-$Name2 $Server") $FreeSpaceGB, $UsedSpaceGB 
                }      
        #-----------------------------------------------------------------------------------------------

        # RAM Info -------------------------------------------------------------------------------------
                $OS = (Get-WmiObject Win32_OperatingSystem -computername $Server).caption
                $SystemInfo = Get-WmiObject -Class Win32_OperatingSystem -computername $Server | Select-Object Name, TotalVisibleMemorySize, FreePhysicalMemory
                $TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB
                $FreeRAM = $SystemInfo.FreePhysicalMemory/1MB
                $UsedRAM = $TotalRAM - $FreeRAM
                $RAMPercentFree = ($FreeRAM / $TotalRAM) * 100
                $TotalRAM = [Math]::Round($TotalRAM, 2)
                $FreeRAM = [Math]::Round($FreeRAM, 2)
                $UsedRAM = [Math]::Round($UsedRAM, 2)
                $RAMPercentFree = [Math]::Round($RAMPercentFree, 2)
        #-----------------------------------------------------------------------------------------------
        
        # Event Logs Report ----------------------------------------------------------------------------
#                $SystemEvents = Get-EventLog -ComputerName $Server -LogName System -EntryType Error, Warning -After (Get-Date).AddHours(-12)
#                $ApplicationEvents = Get-EventLog -ComputerName $Server -LogName Application -EntryType Error, Warning -After (Get-Date).AddHours(-12)
#               
#                $SystemEventsErrors = $SystemEvents | Where {$_.EntryType -eq "Error"}
#                $SystemEventsWarnings = $SystemEvents | Where {$_.EntryType -eq "Warning"}
#                $ApplicationEventsErrors = $ApplicationEvents | Where {$_.EntryType -eq "Error"}
#                $ApplicationEventsWarnings = $ApplicationEvents | Where {$_.EntryType -eq "Warning"}
#          
#                $countSystemErrors = $SystemEventsErrors.count
#                $countSystemWarnings = $SystemEventsWarnings.count
#                $countApplicationErrors = $ApplicationEventsErrors.count
#                $countApplicationWarnings = $ApplicationEventsWarnings.count
#                
#        write-host "System Errors: "$countSystemErrors -nonewline
#        write-host "  System Warning: "$countSystemWarnings -nonewline
#        write-host "  Application Errors: "$countApplicationErrors -nonewline
#        write-host "  Application Warning: "$countApplicationWarnings
               #foreach ($event in $LatestEvents) {               
                 #     $TimeGenerated = $event.TimeGenerated
                 #     $EntryType = $event.EntryType
                 #     $Source = $event.Source
                      #$Message = $event.Message

        #write-host $TimeGenerated $EntryType $Source

          #     }
        #-----------------------------------------------------------------------------------------------
        
        # Services Report ------------------------------------------------------------------------------
                $ServicesReport = ""
                Import-csv -path $MonitoredServices | foreach-object {
                $ServiceName = $_.Service
                    $Services = Get-WmiObject -Class Win32_Service -ComputerName $Server | where {$_.Name -eq "$ServiceName"}
                       foreach ($Service in $Services){
                           if ($Service.State -eq "Stopped"){$color = "#EE3124"} else {$color = "#98CA44"}
                           $ServicesReport += "<b>" + $Service.Name + ": <font color=$color>" + $Service.State + "</font><br>"
                       }
                }
        #-----------------------------------------------------------------------------------------------
        
        # Top Processes Report -------------------------------------------------------------------------
                $ProcessReport = ""
                $TopProcesses = Get-Process -ComputerName $Server | Sort WS -Descending | Select ProcessName, Id, WS -First 5
                   foreach ($Process in $TopProcesses){
                     $ProcessReport += "<b>" + $Process.ProcessName + "</b> " + [Math]::Round($Process.WS/1MB, 2) + "MB<br>"
                   }
        #-----------------------------------------------------------------------------------------------
        
        # Calculate processing time for each item ------------------------------------------------------                    
                $ProcessEnd = Get-Date
                $ProcessTime = $ProcessEnd - $ProcessStart
                $ProcessMinutes = $ProcessTime.minutes
                $ProcessSeconds = $ProcessTime.seconds
        #-----------------------------------------------------------------------------------------------
       
        # Update SharePoint List -----------------------------------------------------------------------
                $DiskCharts = ""
                $Item = $List.Items.Add()
                $Item.update()
                $ItemID = $Item.id
                    $item["Title"] =  "$Server - $FormatedDate"
                    $item["Date"] = $CurrentTime
                    $item["Server"] = $Server   
                    $item["Envirenment"] = $Envirenment
                    $item["SharePoint"] = $SP_Version
                    $item["ServerRole"] = $ServerRole
                    $item["ServerUptime"] = $SystemUptime
                    $item["TotalRAM"] = $TotalRAM   
                       foreach ($Disk in $DiskInfo){
                        $Name = $Disk.Name
                        $SizeGB = $Disk.SizeGB
                        $FreeSpaceGB = $Disk.FreeSpaceGB
                        $UsedSpaceGB = [Math]::Round($SizeGB - $FreeSpaceGB, 2)
                        $PercentFree = $Disk.PercentFree
                         $item[$Name + " Free (%)"] = $PercentFree 
                         $item[$Name + " Free (GB)"] = $FreeSpaceGB
                         $item[$Name + " Size (GB)"] = $SizeGB
                             if ($Disk.PercentFree -lt 30){
                             $Name2 = $Disk.Name.split(":")[0]
                             $ItemID = $Item.ID
                             AddAttachment $item "chart-$Name2 $Server.png"
                             $DiskCharts += @"
                             <table style="font-family:Segoe UI Light; font-size:15px">
                              <tr>
                                <td  rowspan="2" style="font-size:30px; width:25px">$Name</td>
                                <td style="border-bottom:1px solid #007CC1; width:80px">$UsedSpaceGB GB<span style="font-size:10px">Used</span></td>
                                <td rowspan="2" style="font-size:20px"><img src="$SPsite/Lists/$SPlist/Attachments/$ItemID/chart-$Name2 $Server.png">$PercentFree%<span style="font-size:12px">Free</span></td>
                              </tr>
                              <tr>
                                <td style="border-bottom:1px solid #98CA44; width:80px">$FreeSpaceGB GB<span style="font-size:10px">Free</span></td>
                              </tr>
                            </table>
"@
                             $item["Disks < 30%"] = $DiskCharts
                             }
                       } 
                    $item["System Errors"] =  $countSystemErrors
                    $item["System Warnings"] = $countSystemWarnings
                    $item["Application Errors"] = $countApplicationErrors
                    $item["Application Warnings"] = $countApplicationWarnings
                    $item["Processing Time"] = "$ProcessMinutes Mins $ProcessSeconds Secs"
                    $item["Top Processes"] =  $ProcessReport
                    $item["Services Status"] =  $ServicesReport
                $Item.Update()
        #------------------------------------------------------------------------------------------------

# End loop on CSV file
}
#---------------------------------------------------------------- MAIN SCRIPT END -------------------------------------------------------------------------

### Calculate Total running time for script ---------------------------------------------------------------------------------------------------------------
$EndTime = Get-Date
$RunTime = $EndTime - $StartTime
$RunHours = $RunTime.hours
$RunMinutes = $RunTime.minutes
$RunSeconds = $RunTime.seconds

### Update List Descrition --------------------------------------------------------------------------------------------------------------------------------
$List.Description = "Last updated at $EndTime (Started @ $StartTime,  Duration $RunHours hrs $RunMinutes mins $RunSeconds secs)"
$List.Update()

### Free SharePoint site ----------------------------------------------------------------------------------------------------------------------------------
$web.dispose()

### Remove Add-ins ----------------------------------------------------------------------------------------------------------------------------------------
Stop-SPAssignment –global
Remove-PSSnapin Microsoft.SharePoint.PowerShell
