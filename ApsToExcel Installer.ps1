# -*- coding: utf-8 -*-
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Force" -Verb RunAs 
    exit
}
try {
    
    Write-Host "installerar..."
    $scriptUrl = "https://raw.githubusercontent.com/Tehpson/apstoexcel/main/Main.ps1"
    $scriptName = "\apsToExcel.ps1"

    Invoke-WebRequest -Uri $scriptUrl -OutFile "$env:USERPROFILE\$scriptName"

    $scriptFolder = [System.IO.Path]::GetDirectoryName("$env:USERPROFILE\$scriptName")
    [System.Environment]::SetEnvironmentVariable("PATH", $scriptFolder + ";" + [System.Environment]::GetEnvironmentVariable("PATH", [System.EnvironmentVariableTarget]::Machine), [System.EnvironmentVariableTarget]::Machine)

    Install-Module -Name ImportExcel -Force
    Clear-Host
    Read-Host "Installationen är klar. Tryck på Enter för att stänga..."

} catch {
    Write-Host "Fel: $($_.Exception.Message)"
    Read-Host "Installationen är klar. Tryck på Enter för att stänga..."
}