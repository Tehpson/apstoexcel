param (
    [Parameter(Position = 0, Mandatory = $true)]
    [string]$fileName
)

if(!$fileName.EndsWith(".aps")){
    Write-Host "Detta är inte en .aps fil"
    exit
}

Import-Module -Name ImportExcel

$currentDirectory = $PWD.Path
$fileName = $fileName.TrimStart(".\")
$filePath = $currentDirectory + "\" + $fileName
$exportPath = $filePath.TrimEnd(".aps") + ".xlsx"

$lines = Get-Content -Path $filePath
$exportData = @()
foreach ($line in $lines) {
    $match = $line | Select-String 'Name="([^"]+)"\s+Value="([^"]*)"\s+Text="([^"]+)"'
    if ($match) {
        $Name = $match.Matches.Groups[1].Value
        $Value = $match.Matches.Groups[2].Value
        $Text = $match.Matches.Groups[3].Value
        $Name = $Name -replace '\s+', ' '
        $Value = $Value -replace '\s+', ' '
        $Text = $Text -replace '\s+', ' '
        $csvObj = [PSCustomObject]@{
            Name = $Name
            Value = $Value
            Text = $Text
        }
        $exportData += $csvObj
    }
}

$exportData | Export-Excel -Path $exportPath -AutoSize -Show -IncludePivotTable -AutoFilter
