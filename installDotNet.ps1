function installDotNet {
    $DNVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 528040

    if (-Not $DNVersion) {
	    # Install latest .net 4.8
        Write-Host "Installing Dot Net 4.8" -ForegroundColor Yellow
	    $DNSourceURL = "https://go.microsoft.com/fwlink/?linkid=2088631"
	    $DNInstaller = $env:TEMP + "\dotnet.exe"
	    Invoke-WebRequest $DNSourceURL -OutFile $DNInstaller
	    Start-Process -FilePath $DNInstaller -Args "/s" -Verb RunAs -Wait
	    Remove-Item $DNInstaller
    }
    else {
        Write-Host "Dot Net 4.8 or higher already installed, skipping..." -ForegroundColor Yellow
    }
}

installDotNet