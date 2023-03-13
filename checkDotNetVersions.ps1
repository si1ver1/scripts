# Allow script to run temporarily
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

$Values = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | Get-ItemProperty -Name version -EA 0 | Where { $_.PSChildName -Match '^(?!S)\p{L}'} | Select PSChildName, version
$Values |Format-Table | Out-String | Write-Host

Read-Host -Prompt "Press Enter to exit"