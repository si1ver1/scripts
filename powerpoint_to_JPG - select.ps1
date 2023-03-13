# Powershell script to export Powerpoint slides as jpg images using the Powerpoint COM API

function Get-FileName
{
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    #$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
}

function Export-Slide($inputFile)
{
	# Load Powerpoint Interop Assembly
	[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
	[Reflection.Assembly]::LoadWithPartialname("Office") > $null

	$msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue

	# start Powerpoint
	$application = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

	# Make sure inputFile is an absolte path
	$inputFile = Resolve-Path $inputFile
    $outputFile = Split-Path $inputFile -Parent
    #$outputFile = $outputFile + "\" + (get-date).ToString("yyyy-MM-dd")
    #write-host $outputFile

	$presentation = $application.Presentations.Open($inputFile, $msoTrue, $msoFalse, $msoFalse)
	$width = 1920
	$height = 1080

	$presentation.Export($outputFile, "JPG", $width, $height)
	
	$slide = $null
	
	$presentation.Close()
	$presentation = $null
	
	if($application.Windows.Count -eq 0)
	{
		$application.Quit()
	}
	
	$application = $null
	
	# Make sure references to COM objects are released, otherwise powerpoint might not close
	# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();       
}

$filename = Get-FileName
Export-Slide -inputFile