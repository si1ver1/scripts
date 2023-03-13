# This script will do the following:
# Create a full backup of the Windows Application and System Event Logs
# Create a CSV file limited to Errors and Warnings from Application and System Events since the last 7 days
# Create a zip file of all the above files and open the explorer window to the location of all these files for easy transfer

Set-Variable -Name EventAgeDays -Value 7 # we will take events for the latest 7 days
Set-Variable -Name LogNames -Value @("Application”, "System”) # Checking application and system logs
Set-Variable -Name EventTypes -Value @("Error”, "Warning”) # Loading only Errors and Warnings
Set-Variable -Name ExportFolder -Value "C:\templogs\” # temp folder to save files

$el_c = @() #consolidated error log
$now=get-date
$startdate=$now.adddays(-$EventAgeDays)
$ExportFile=$ExportFolder + $(Hostname) + $now.ToString("_yyyy-MM-dd_hh-mm-ss”) + ".csv”

# Create our temp folder if it doesn't exist
If(!(test-path $ExportFolder))
{
      New-Item -ItemType Directory -Force -Path $ExportFolder
}


foreach($log in $LogNames)
{
    # Export full application and system logs
    Write-Host Processing $log
    $exportFileName = $(Hostname) + "_" + $log + $now.ToString("_yyyy-MM-dd_hh-mm-ss”) + ".evt"
    $logFile = Get-WmiObject Win32_NTEventlogFile | Where-Object {$_.logfilename -eq $log}
    $logFile.backupeventlog($ExportFolder + $exportFileName)

    Write-Host Processing $log
    $el = get-eventlog -log $log -After $startdate -EntryType $EventTypes
    $el_c += $el # consolidating
}

$el_sorted = $el_c | Sort-Object TimeGenerated # sort by time
Write-Host Exporting to $ExportFile
$el_sorted|Select EntryType, TimeGenerated, Source, EventID, Message | Export-CSV $ExportFile -NoTypeInfo # EXPORT

$zip = $ExportFolder + $(Hostname) + '.zip'
Compress-Archive -Path $ExportFolder -DestinationPath $zip -Force

# Open explorer to saved location
Invoke-Item $ExportFolder
