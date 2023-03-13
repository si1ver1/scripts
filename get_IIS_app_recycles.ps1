# Export to csv the last weeks application pool recycles

Set-Variable -Name EventAgeDays -Value 7 # we will take events for the latest 7 days
Set-Variable -Name ExportFolder -Value "C:\templog\” # temp folder to save files

$el_c = @() #consolidated error log
$now=get-date
$startdate=$now.adddays(-$EventAgeDays)
$ExportFile=$ExportFolder + $(Hostname) + $now.ToString("_yyyy-MM-dd_hh-mm-ss”) + ".csv”

# Create our temp folder if it doesn't exist
If(!(test-path $ExportFolder))
{
      New-Item -ItemType Directory -Force -Path $ExportFolder
}

# Get all application pool recycles since our start date
$el_c = Get-WinEvent -LogName System | Where-Object {($_.Message -like "*recycle*") -and ($_.TimeCreated -ge $startdate)}

$el_sorted = $el_c | Sort-Object TimeCreated # sort by time
Write-Host Exporting to $ExportFile
$el_sorted|Select TimeCreated, Id, Message | Export-CSV $ExportFile -NoTypeInfo # EXPORT

# Open explorer to saved location
Invoke-Item $ExportFolder