# Print-ExcelSheet-Enhanced.ps1
[CmdletBinding()]
param(
    [string]$Path  = 'C:\powershell\wip\wip sheet.xlsx',
    [string]$Sheet = 'wip',
    [string]$Cell  = 'C5',
    [string]$Value = 'PolyPebt',

    [ValidateRange(1,100)]
    [int]$Copies = 5,

    [switch]$SaveAfterChange,
    [switch]$ShowExcel,
    [switch]$LeaveOpen,
    [switch]$PreviewOnly,

    # Use this only if you know the exact Excel string, e.g. 'HP LaserJet on Ne05:'
    [string]$ActivePrinterRaw,

    # Recommended: pass the Windows printer display name (as seen in Settings/Control Panel)
    [string]$PrinterName = 'PRT-HP-LJ-IT',

    # Additional enhancements
    [hashtable]$CellUpdates = @{},  # Multiple cell updates: @{'A1'='Value1'; 'B2'='Value2'}
    [string]$PrintRange = '',       # Specific range to print (e.g., 'A1:E10')
    [switch]$Landscape,             # Print orientation
    [switch]$FitToPage,             # Fit to one page
    [switch]$Verbose
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Error' { 'Red' }
        'Warning' { 'Yellow' }
        'Success' { 'Green' }
        default { 'White' }
    }
    Write-Host "[$timestamp] $Message" -ForegroundColor $color
}

function Release-ComObject {
    param([Parameter(Mandatory=$true)][object]$ComObj)
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObj) } catch {}
}

function Set-DefaultPrinter {
    param([Parameter(Mandatory)][string]$Name)
    $p = Get-CimInstance Win32_Printer -Filter ("Name='{0}'" -f ($Name -replace "'", "''")) -ErrorAction SilentlyContinue
    if (-not $p) { throw "Printer '$Name' not found on this machine/user profile." }
    $null = $p.SetDefaultPrinter()
    Write-Log "Set Windows default printer to: $Name" -Level 'Success'
}

function Get-AvailablePrinters {
    Write-Log "Available printers:"
    Get-CimInstance Win32_Printer | ForEach-Object {
        $status = if ($_.Default) { " (DEFAULT)" } else { "" }
        Write-Log "  - $($_.Name)$status"
    }
}

function Invoke-PrintWorksheet {
    param(
        [Parameter(Mandatory)][__ComObject]$Worksheet,
        [int]$Copies = 1,
        [string]$PrintRange = '',
        [switch]$PreviewOnly
    )

    if ($PreviewOnly) {
        Write-Log "Opening print preview..." -Level 'Success'
        $Worksheet.PrintPreview()
    } else {
        Write-Log "Printing $Copies copies..." -Level 'Success'

        if ($PrintRange) {
            # Print specific range
            $rangeObj = $Worksheet.Range($PrintRange)
            $null = $rangeObj.PrintOut($null, $null, [int]$Copies, $false, $null, $false, $true, $null)
        } else {
            # Print entire worksheet
            $null = $Worksheet.PrintOut($null, $null, [int]$Copies, $false, $null, $false, $true, $null)
        }

        Write-Log "Print job completed successfully!" -Level 'Success'
    }
}

function Try-SetExcelActivePrinter {
    param(
        [Parameter(Mandatory)][__ComObject]$Excel,
        [string]$ActivePrinterRaw,
        [string]$PrinterName
    )

    if ($ActivePrinterRaw) {
        try {
            $Excel.ActivePrinter = $ActivePrinterRaw
            Write-Log "Set Excel ActivePrinter to: $ActivePrinterRaw" -Level 'Success'
            return $true
        } catch {}
        throw "Provided -ActivePrinterRaw ('$ActivePrinterRaw') not accepted by Excel."
    }

    if (-not $PrinterName) {
        Write-Log "No specific printer requested, using Excel default"
        return $true
    }

    # If Excel already picked it up, we're done
    try {
        if ($Excel.ActivePrinter -like "$PrinterName*") {
            Write-Log "Excel already using correct printer: $($Excel.ActivePrinter)" -Level 'Success'
            return $true
        }
    } catch {}

    Write-Log "Attempting to set Excel printer to: $PrinterName"

    # Build candidates and try them
    $prn = Get-CimInstance Win32_Printer -Filter ("Name='{0}'" -f ($PrinterName -replace "'", "''")) -ErrorAction SilentlyContinue
    $namesToTry = @()
    if ($prn) { $namesToTry += $prn.Name }
    $namesToTry = $namesToTry | Select-Object -Unique

    foreach ($nm in $namesToTry) {
        # Try network printer formats
        foreach ($i in 0..99) {
            $cand = '{0} on Ne{1:D2}:' -f $nm, $i
            try {
                $Excel.ActivePrinter = $cand
                Write-Log "Successfully set Excel printer to: $cand" -Level 'Success'
                return $true
            } catch {}
        }

        # Try port-based formats
        if ($prn -and $prn.PortName) {
            $port = $prn.PortName.TrimEnd(':')
            foreach ($suffix in @("$port:", "$port")) {
                $cand = '{0} on {1}' -f $nm, $suffix
                try {
                    $Excel.ActivePrinter = $cand
                    Write-Log "Successfully set Excel printer to: $cand" -Level 'Success'
                    return $true
                } catch {}
            }
        }

        # Try well-known ports
        foreach ($wellKnown in @('USB001:','XPSPort:','FILE:','PORTPROMPT:')) {
            $cand = '{0} on {1}' -f $nm, $wellKnown
            try {
                $Excel.ActivePrinter = $cand
                Write-Log "Successfully set Excel printer to: $cand" -Level 'Success'
                return $true
            } catch {}
        }
    }

    Write-Log "Failed to set Excel printer to: $PrinterName" -Level 'Warning'
    return $false
}

# Validate file exists
if (-not (Test-Path -LiteralPath $Path)) {
    throw "File not found: $Path"
}

Write-Log "Starting Excel print job for: $Path"

# Show available printers if verbose
if ($Verbose) {
    Get-AvailablePrinters
}

# 1) Set Windows default printer (most reliable)
if ($PrinterName -and -not $ActivePrinterRaw) {
    try {
        Set-DefaultPrinter -Name $PrinterName
    } catch {
        Write-Log "Failed to set Windows default printer to '$PrinterName'. $($_.Exception.Message)" -Level 'Warning'
        if ($Verbose) { Get-AvailablePrinters }
        throw
    }
}

$excel = $wb = $ws = $null
try {
    Write-Log "Opening Excel application..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$ShowExcel
    $excel.DisplayAlerts = $false

    # 2) Set Excel printer
    $ok = $true
    try {
        $ok = Try-SetExcelActivePrinter -Excel $excel -ActivePrinterRaw $ActivePrinterRaw -PrinterName $PrinterName
    } catch {
        throw "Printer setup failed. $($_.Exception.Message)"
    }

    if (-not $ok) {
        $winDefault = (Get-CimInstance Win32_Printer | Where-Object Default -eq $true | Select-Object -Expand Name -First 1)
        $excelAp = $null; try { $excelAp = $excel.ActivePrinter } catch {}
        throw "Printer setup failed. Could not set Excel.ActivePrinter for '$PrinterName'. Excel.ActivePrinter='$excelAp'. Windows Default='$winDefault'."
    }

    Write-Log "Opening workbook: $Path"
    $wb = $excel.Workbooks.Open($Path)

    try {
        $ws = $wb.Worksheets.Item($Sheet)
        Write-Log "Activated worksheet: $Sheet"
    } catch {
        throw "Worksheet not found: '$Sheet'."
    }

    $ws.Activate() | Out-Null

    # Update single cell (backward compatibility)
    if ($Cell -and $Value) {
        Write-Log "Updating cell $Cell with value: $Value"
        $ws.Range($Cell).Value2 = $Value
    }

    # Update multiple cells if provided
    if ($CellUpdates.Count -gt 0) {
        Write-Log "Updating multiple cells..."
        foreach ($cellAddr in $CellUpdates.Keys) {
            $cellValue = $CellUpdates[$cellAddr]
            Write-Log "  Setting $cellAddr = $cellValue"
            $ws.Range($cellAddr).Value2 = $cellValue
        }
    }

    # Set print orientation
    if ($Landscape) {
        Write-Log "Setting landscape orientation"
        $ws.PageSetup.Orientation = 2  # xlLandscape
    }

    # Fit to page
    if ($FitToPage) {
        Write-Log "Setting fit to page"
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = 1
    }

    if ($SaveAfterChange) {
        Write-Log "Saving workbook..."
        $wb.Save()
    }

    Write-Log "Excel.ActivePrinter -> $($excel.ActivePrinter)" -Level 'Success'

    # Call the print function
    # Invoke-PrintWorksheet -Worksheet $ws -Copies $Copies -PrintRange $PrintRange -PreviewOnly:$PreviewOnly
    Write-Log "Printing disabled for testing - cell update only" -Level 'Warning'

} catch {
    Write-Log "Error occurred: $($_.Exception.Message)" -Level 'Error'
    throw
} finally {
    Write-Log "Cleaning up COM objects..."

    if ($wb -and -not $LeaveOpen) {
        try { $wb.Close($false) | Out-Null } catch {}
    }
    if ($excel -and -not $LeaveOpen) {
        try { $excel.Quit() | Out-Null } catch {}
    }

    if ($ws) { Release-ComObject $ws }
    if ($wb) { Release-ComObject $wb }
    if ($excel) { Release-ComObject $excel }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Log "Cleanup completed"
}

Write-Log "Script execution completed"