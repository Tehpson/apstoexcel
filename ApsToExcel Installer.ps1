# Kontrollera om användaren är administratör
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Scriptet kräver administratörsbehörigheter. Kör om som administratör."
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    return
}

Invoke-WebRequest -Uri "https://github.com/Tehpson/apstoexcel/releases/download/1.0.1.2/APS2XLSX.exe" -OutFile "APS2XLSX.exe"

$installFolder = "C:\Program Files\Aps2Excel"

Install-Module -Name ImportExcel -Force
Copy-Item -Path "APS2XLSX.exe" -Destination $installFolder

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut("$env:USERPROFILE\Desktop\APS2XLSX.lnk")
$shortcut.TargetPath = "$installFolder\APS2XLSX.exe"
$shortcut.Save()
