
Add-Type -AssemblyName System.Windows.Forms
Import-Module -Name ImportExcel

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('MyDocuments')
$openFileDialog.Filter = 'APS Files (*.aps)|*.aps|All Files (*.*)|*.*'
$openFileDialog.Title = 'Select a file'

if ($openFileDialog.ShowDialog() -eq 'OK') {
    $selectedFilePath = $openFileDialog.FileName
} else {
    Write-Host "No file selected"
    return
}

if (!$selectedFilePath.EndsWith(".aps")) {
    Write-Host "Detta är inte en .aps fil"
    return
}

$lines = Get-Content -Path $selectedFilePath
$exportData = @()
$radData = @()

foreach ($line in $lines) {
    if ($line -match '<Var\s+Name="([^"]*)"\s+Value="([^"]*)"(?:\s+Unit="([^"]*)")?\s+Text="([^"]*)".*') {

        $Name = $matches[1]
        $Value = $matches[2]
        $Unit = $matches[3]
        $Text = $matches[4]

        if ($Text -match '^Rad\s+(\d+)') {
            $RadNumber = [int]::Parse($matches[1])
            $PaddedRadNumber = $RadNumber.ToString("D3")
            $radData += [PSCustomObject]@{
                Name = $Name
                Value = $Value
                Unit = $Unit
                Text = $Text
                RadNumber = $PaddedRadNumber
            }
        } else {
            $exportData += [PSCustomObject]@{
                Name = $Name
                Value = $Value
                Unit = $Unit
                Text = $Text
            }
        }
    }
}

$radData = $radData | Sort-Object -Property RadNumber
$exportData += $radData

$fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($selectedFilePath)
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$exportPath = [System.Environment]::GetFolderPath('MyDocuments') + "\" + $fileNameWithoutExtension + "_$timestamp.xlsx"

$exportData | Export-Excel -Path $exportPath -AutoSize -Show -TableName "MyTable"