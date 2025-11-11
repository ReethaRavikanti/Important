<# 
.SYNOPSIS
    Reads a list of computer names from an Excel file, checks if they have a specific log file, 
    and exports those with the log file to a new Excel file.
#>

# --- Variables ---
$SourceExcel = "C:\Scripts\Computers.xlsx"     # Input Excel file path
$OutputExcel = "C:\Scripts\Computers_With_Logs.xlsx"  # Output Excel file path
$SheetName = "Sheet1"                          # Excel sheet name
$LogPath = "C$\ProgramData\App\Logs\example.log"  # Log file path to check

# --- Check for ImportExcel module ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

# --- Read Computer Names from Excel ---
$Computers = Import-Excel -Path $SourceExcel -WorksheetName $SheetName

if (-not $Computers) {
    Write-Host "No data found in Excel file!" -ForegroundColor Red
    exit
}

# Assume the column name is 'ComputerName' – adjust if needed
$Results = @()

foreach ($Computer in $Computers) {
    $CompName = $Computer.ComputerName

    if ([string]::IsNullOrWhiteSpace($CompName)) { continue }

    Write-Host "Checking $CompName ..." -ForegroundColor Cyan

    $PathToCheck = "\\$CompName\$LogPath"

    try {
        if (Test-Path $PathToCheck) {
            Write-Host "✔ Log file found on $CompName" -ForegroundColor Green
            $Results += [PSCustomObject]@{
                ComputerName = $CompName
                LogFileFound = "Yes"
                CheckedAt     = (Get-Date)
            }
        } else {
            Write-Host "✖ Log file NOT found on $CompName" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Host "⚠ Error connecting to $CompName: $_" -ForegroundColor Red
    }
}

# --- Export results to new Excel file ---
if ($Results.Count -gt 0) {
    $Results | Export-Excel -Path $OutputExcel -AutoSize -BoldTopRow -Title "Computers with Log File"
    Write-Host "`n✅ Exported $($Results.Count) computers to $OutputExcel" -ForegroundColor Green
} else {
    Write-Host "`nNo computers had the log file. Nothing exported." -ForegroundColor Yellow
}
