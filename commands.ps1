# Robocopy /s includes subdirectories, /e includes empty folders, *.* includes all files and all extensions
Robocopy C:\directory D:\directory /s /e *.*

# Search files within a path for specific string
Get-ChildItem -PATH "C:\Program Files\xchangexec\xchangepoint" -Recurse | Select-String -Pattern 'search'

# Extract zip files
Get-ChildItem $foldername -Filter *.zip | ForEach-Object {Expand-Archive $_.FullName -Force}
# Recusively Search through log files for specific string
Get-ChildItem -r | Select-String "Search Term"
# Same as above but output filename
Get-ChildItem -r | Select-String 'Search Term' | Select Path

# Search for a pattern excluding another pattern on that line
Get-ChildItem -r | Select-String -pattern 'status code:' | where { $_.line -NotLike '*OK (200)*' }

# Search event logs by log type and date range
Get-EventLog -LogName Application -After (Get-Date -Date '1/26/2020') -Before (Get-Date -Date '1/27/2020')
# Filter event logs by log type and search phrase
Get-WinEvent -FilterHashtable @{LogName='Application'} | Where-Object -Property Message -Match 'XchangePoint.exe'
# Search event logs by log type, date range, and search phrase
Get-WinEvent -FilterHashtable @{LogName='Application';StartTime='1/20/2020';EndTime='1/21/2020'} | Where-Object -Property Message -Match 'XchangePoint.exe'

# View all Users
Get-LocalUser | Select *

# View last 10 reboots
get-eventlog system | where-object {$_.eventid -eq 6006} | select -first 10

# View IIS recycles
Get-WinEvent -LogName System | Where-Object {$_.Message -like "*recycle*"} | Out-File C:\recycles.txt

# Change domain password
Set-AdAccountPassword -Identity AccountName

# Unblock all files in a directory
gci C:\directory\ | Unblock-File

# View local network devices
Get-NetIPAddress

# Compare two directories
$fso = Get-ChildItem -Recurse -path 'C:\Program Files\xchangexec\xchangepoint'
$fsoBU = Get-ChildItem -Recurse -path 'C:\Program Files\xchangexec\xchangepointold'
Compare-Object -ReferenceObject $fso -DifferenceObject $fsoBU

# View app pools and their processID
Import-Module WebAdministration
dir "IIS:\AppPools" | % {dir "IIS:\AppPools\$($_.name)\WorkerProcesses"} | select processId, appPoolName | format-table -AutoSize
