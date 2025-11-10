<#
.SYNOPSIS
  Read computer names from an Excel file, scan a log file on each computer for a string, output matched computer names.

.NOTES
  - Two ways to read Excel:
      * Recommended: ImportExcel module (Install-Module -Name ImportExcel)
      * Fallback: Excel COM interop (works on machines with Excel installed)
  - Two ways to access remote logs:
      * Remoting (Invoke-Command) - requires PSRemoting enabled and appropriate credentials.
      * Admin share (\\COMPUTER\C$\path\file.log) - requires admin rights and file-share access.
  - Results saved to CSV and copied to clipboard.

.EXAMPLE
  .\Scan-LogFromExcel.ps1 -ExcelPath "C:\temp\computers.xlsx" -LogPath "C:\Logs\app.log" -SearchString "ERROR" -UseRemoting
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ExcelPath = "C:\temp\computers.xlsx",

    [Parameter(Mandatory=$false)]
    [string]$SheetName = "",            # leave empty to use first sheet

    [Parameter(Mandatory=$false)]
    [string]$ColumnName = "Computer",   # column header that contains computer names

    [Parameter(Mandatory=$false)]
    [string]$LogPath = "C:\Logs\app.log",  # path as it's on the remote machine

    [Parameter(Mandatory=$true)]
    [string]$SearchString,               # string to search for

    [Parameter(Mandatory=$false)]
    [string]$OutputCsv = "C:\temp\matched-computers.csv",

    [switch]$UseRemoting,                # prefer Invoke-Command
    [switch]$UseAdminShare,              # prefer \\computer\C$\path\file.log

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)

function Get-ComputerListFromExcel {
    param(
        [string]$Path,
        [string]$Sheet,
        [string]$ColName
    )

    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
            if ($Sheet) {
                $rows = Import-Excel -Path $Path -WorkSheetname $Sheet
            } else {
                $rows = Import-Excel -Path $Path
            }
            $computers = $rows | Select-Object -ExpandProperty $ColName -ErrorAction Stop
            return $computers | Where-Object { $_ -and $_.Trim() -ne "" } | ForEach-Object { $_.Trim() } | Sort-Object -Unique
        } catch {
            Write-Warning "ImportExcel present but failed to read ($Path): $_. Falling back to COM method."
        }
    }

    # Fallback: COM Excel (requires Excel installed)
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).ProviderPath)
        if ($Sheet) {
            $ws = $wb.Worksheets | Where-Object { $_.Name -ieq $Sheet }
            if (-not $ws) { throw "Sheet '$Sheet' not found." }
        } else {
            $ws = $wb.Worksheets.Item(1)
        }

        # find header row (assume header in row 1)
        $used = $ws.UsedRange
        $headerRange = $used.Rows.Item(1)
        $colIndex = $null
        for ($c=1; $c -le $headerRange.Columns.Count; $c++) {
            $val = ($headerRange.Columns.Item($c)).Value2
            if ($val -and $val.ToString().Trim() -ieq $ColName) { $colIndex = $c; break }
        }
        if (-not $colIndex) { throw "Column '$ColName' not found in sheet." }

        $values = @()
        for ($r=2; $r -le $used.Rows.Count; $r++) {
            $cell = $ws.Cells.Item($r, $colIndex).Text
            if ($cell -and $cell.Trim() -ne "") { $values += $cell.Trim() }
        }

        $wb.Close($false)
        $excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        return $values | Sort-Object -Unique
    } catch {
        throw "Failed to read Excel file using ImportExcel or COM: $_"
    }
}

function Test-LogForString_AdminShare {
    param(
        [string]$Computer,
        [string]$RemotePath,
        [string]$Pattern
    )
    # convert c:\path\file to \\computer\C$\path\file
    if ($RemotePath -match "^([a-zA-Z]:)\\(.+)$") {
        $drive = $Matches[1].TrimEnd(':')
        $rest = $Matches[2]
        $unc = "\\$Computer\$($drive)`$$\" + $rest
    } else {
        # if UNC already or something else
        $unc = "\\$Computer\$RemotePath"
    }

    try {
        if (-not (Test-Path -LiteralPath $unc)) {
            Write-Verbose "File not found: $unc"
            return $false
        }
        # Use Select-String for streaming search
        $match = Select-String -Path $unc -Pattern $Pattern -SimpleMatch -Quiet -ErrorAction SilentlyContinue
        return [bool]$match
    } catch {
        Write-Verbose "Admin-share access failed for $Computer: $_"
        return $false
    }
}

function Test-LogForString_Remoting {
    param(
        [string]$Computer,
        [string]$RemotePath,
        [string]$Pattern,
        [System.Management.Automation.PSCredential]$Cred
    )
    try {
        $script = {
            param($p,$pat)
            if (-not (Test-Path -LiteralPath $p)) { return $false }
            Select-String -Path $p -Pattern $pat -SimpleMatch -Quiet
        }
        if ($Cred) {
            $res = Invoke-Command -ComputerName $Computer -ScriptBlock $script -ArgumentList $RemotePath, $Pattern -Credential $Cred -ErrorAction Stop -Authentication Negotiate
        } else {
            $res = Invoke-Command -ComputerName $Computer -ScriptBlock $script -ArgumentList $RemotePath, $Pattern -ErrorAction Stop
        }
        return [bool]$res
    } catch {
        Write-Verbose "Remoting failed for $Computer: $_"
        return $false
    }
}

# -------------------------
# Main
# -------------------------
try {
    if (-not (Test-Path -Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }

    Write-Host "Reading computer list from: $ExcelPath" -ForegroundColor Cyan
    $computers = Get-ComputerListFromExcel -Path $ExcelPath -Sheet $SheetName -ColName $ColumnName
    if (-not $computers -or $computers.Count -eq 0) { throw "No computer names found in Excel column '$ColumnName'." }

    Write-Host "Found $($computers.Count) computers. Starting scan..." -ForegroundColor Cyan

    $matched = [System.Collections.Generic.List[string]]::new()

    foreach ($c in $computers) {
        Write-Host -NoNewline "Checking $c ... "
        $found = $false

        if ($UseRemoting) {
            $found = Test-LogForString_Remoting -Computer $c -RemotePath $LogPath -Pattern $SearchString -Cred $Credential
        } elseif ($UseAdminShare) {
            $found = Test-LogForString_AdminShare -Computer $c -RemotePath $LogPath -Pattern $SearchString
        } else {
            # Try remoting first, then admin share as fallback
            $found = Test-LogForString_Remoting -Computer $c -RemotePath $LogPath -Pattern $SearchString -Cred $Credential
            if (-not $found) {
                $found = Test-LogForString_AdminShare -Computer $c -RemotePath $LogPath -Pattern $SearchString
            }
        }

        if ($found) {
            Write-Host "MATCH" -ForegroundColor Green
            $matched.Add($c)
        } else {
            Write-Host "no" -ForegroundColor DarkGray
        }
    }

    # Save results
    $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $out = [PSCustomObject]@{
        Timestamp = $timestamp
        MatchedComputer = $null
    }

    if ($matched.Count -gt 0) {
        $matched | Sort-Object | ForEach-Object {
            [PSCustomObject]@{ Computer = $_ }
        } | Export-Csv -Path $OutputCsv -NoTypeInformation -Force

        # Copy to clipboard (one-per-line)
        $matched -join "`r`n" | Set-Clipboard

        Write-Host "Found $($matched.Count) matching computers. Results saved to: $OutputCsv" -ForegroundColor Yellow
        Write-Host "Computer names copied to clipboard (one per line)." -ForegroundColor Yellow
    } else {
        Write-Host "No matches found." -ForegroundColor Yellow
        # write empty CSV
        @() | Export-Csv -Path $OutputCsv -NoTypeInformation -Force
    }

} catch {
    Write-Error "Script failed: $_"
}
