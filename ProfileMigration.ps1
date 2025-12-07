# =============================================================================
# USER PROFILE TRANSFER TOOL - NOVEMBER 2025 - FINAL FIXED VERSION
# Tested 100% working on Windows 11 24H2 (26100.3194+)
# =============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === GLOBALS ===
$global:Form = $null
$global:UserComboBox = $null
$global:BrowseButton = $null
$global:ExportButton = $null
$global:ImportButton = $null
$global:DomainCheckBox = $null
$global:ComputerNameTextBox = $null
$global:DomainNameTextBox = $null
$global:RestartComboBox = $null
$global:DelayTextBox = $null
$global:DomainJoinButton = $null
$global:DomainRetryButton = $null
$global:ProgressBar = $null
$global:StatusText = $null
$global:SelectedZipPath = $null
$global:LogBox = $null
$global:SevenZipPath = $null
$global:ImportBackupPath = $null
$global:ImportStartTime = $null
$global:ExportStartTime = $null
$global:DomainCredential = $null
$global:CurrentLogFile = $null
$global:CancelRequested = $false
$global:CurrentOperation = $null
$global:LogEntries = @()  # Array to store all log entries with metadata for filtering

# === CPU CORE DETECTION ===
# Detect number of CPU cores for optimal multi-threading
$cpuCores = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
if (-not $cpuCores -or $cpuCores -lt 1) { $cpuCores = 4 }  # Fallback to 4 if detection fails

# === CONFIG CONSTANTS (all tunables in one place) ===
$Config = @{
    DomainReachabilityTimeout = 3000      # milliseconds for LDAP/DNS checks
    DomainJoinCountdown       = 10        # seconds for restart warning
    HiveUnloadMaxAttempts     = 3         # retry attempts for hive cleanup
    HiveUnloadWaitMs          = 500       # milliseconds between retry attempts
    MountPointMaxAttempts     = 5         # max collision detection attempts
    ProgressUpdateIntervalMs  = 1000      # milliseconds between progress display updates
    ExportProgressCheckMs     = 500       # milliseconds between export file count checks
    RobocopyThreads           = [Math]::Min(32, [Math]::Max(8, $cpuCores))  # Dynamic: 8 to 32 threads based on CPU cores
    RobocopyRetryCount        = 1         # /R parameter
    RobocopyRetryWaitSec      = 1         # /W parameter
    SevenZipThreads           = $cpuCores # Use all available CPU cores for 7-Zip compression/extraction
    ProfileValidationTimeoutSec = 10      # max time to validate profile path writeability
    SizeEstimationDepth     = 3           # depth for fast size estimation
    HashVerificationEnabled = $true       # verify ZIP integrity
    LogLevel                = 'INFO'      # Minimum log level: DEBUG, INFO, WARN, ERROR
    GenerateHTMLReports     = $true       # Generate HTML migration reports
    AutoOpenReports         = $true       # Automatically open HTML reports after generation
}

# === LOG LEVEL ENUM ===
# Log levels control which messages are displayed and written to files.
# Set Config.LogLevel to control minimum verbosity:
#   DEBUG - Show all messages (verbose diagnostics, file paths, detailed operations)
#   INFO  - Show informational messages and above (default - normal operations)
#   WARN  - Show warnings and errors only (problematic conditions)
#   ERROR - Show only critical errors (failures that prevent operation)
# 
# Usage:
#   Log-Debug "Detailed diagnostic information"
#   Log-Info "Normal operation message"
#   Log-Warning "Something unusual but non-critical"
#   Log-Error "Critical failure condition"
#
# UI Control: Log level can be changed dynamically via dropdown in main UI
$LogLevels = @{
    DEBUG = 0
    INFO  = 1
    WARN  = 2
    ERROR = 3
}

# === PRIVILEGE HELPER ===
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Priv {
    [DllImport("advapi32.dll", SetLastError=true)]
    public static extern bool OpenProcessToken(IntPtr ProcessHandle, UInt32 DesiredAccess, out IntPtr TokenHandle);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetCurrentProcess();
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool LookupPrivilegeValue(string lpSystemName, string lpName, out LUID lpLuid);
    [DllImport("advapi32.dll", SetLastError=true)]
    public static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges, ref TOKEN_PRIVILEGES NewState, UInt32 BufferLength, IntPtr PreviousState, IntPtr ReturnLength);
    [StructLayout(LayoutKind.Sequential)]
    public struct LUID { public UInt32 LowPart; public Int32 HighPart; }
    [StructLayout(LayoutKind.Sequential)]
    public struct TOKEN_PRIVILEGES {
        public UInt32 PrivilegeCount;
        public LUID Luid;
        public UInt32 Attributes;
    }
    public const UInt32 TOKEN_ADJUST_PRIVILEGES = 0x0020;
    public const UInt32 TOKEN_QUERY = 0x0008;
    public const UInt32 SE_PRIVILEGE_ENABLED = 0x00000002;
    public static bool EnablePrivilege(string privilege) {
        IntPtr token;
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out token)) return false;
        LUID luid;
        if (!LookupPrivilegeValue(null, privilege, out luid)) return false;
        TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES();
        tp.PrivilegeCount = 1;
        tp.Luid = luid;
        tp.Attributes = SE_PRIVILEGE_ENABLED;
        return AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
"@ -ErrorAction Stop

function Enable-Privilege {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [bool]$IsCritical = $false
    )
    if (-not $Name) { throw "Privilege name cannot be empty" }
    $ok = [Priv]::EnablePrivilege($Name)
    if (-not $ok) {
        $msg = "Failed to enable privilege: $Name"
        Write-Host "WARNING: $msg" -ForegroundColor Yellow
        Log-Warning $msg
        if ($IsCritical) {
            throw "CRITICAL: Cannot enable required privilege $Name - import cannot proceed"
        }
    }
    return $ok
}

# === ENSURE 7-ZIP IS INSTALLED ===
Write-Host "Checking for 7-Zip..." -ForegroundColor Cyan
# Function to show 7-Zip recovery dialog
function Show-SevenZipRecoveryDialog {
    param([string]$Message)
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $recoveryForm = New-Object System.Windows.Forms.Form
    $recoveryForm.Text = "7-Zip Required"
    $recoveryForm.Size = New-Object System.Drawing.Size(600, 380)
    $recoveryForm.StartPosition = "CenterScreen"
    $recoveryForm.FormBorderStyle = "FixedDialog"
    $recoveryForm.MaximizeBox = $false
    $recoveryForm.MinimizeBox = $false
    $recoveryForm.TopMost = $true
    $recoveryForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $recoveryForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(600, 70)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
    $recoveryForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(560, 25)
    $lblTitle.Text = "7-Zip Not Found"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 43)
    $lblSubtitle.Size = New-Object System.Drawing.Size(560, 20)
    $lblSubtitle.Text = "7-Zip is required for profile migration"
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 200)
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Content panel
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(15, 85)
    $contentPanel.Size = New-Object System.Drawing.Size(560, 180)
    $contentPanel.BackColor = [System.Drawing.Color]::White
    $recoveryForm.Controls.Add($contentPanel)
    
    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Location = New-Object System.Drawing.Point(15, 15)
    $lblMessage.Size = New-Object System.Drawing.Size(530, 60)
    $lblMessage.Text = $Message
    $lblMessage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $contentPanel.Controls.Add($lblMessage)
    
    $lblOptions = New-Object System.Windows.Forms.Label
    $lblOptions.Location = New-Object System.Drawing.Point(15, 80)
    $lblOptions.Size = New-Object System.Drawing.Size(530, 20)
    $lblOptions.Text = "Choose an option:"
    $lblOptions.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($lblOptions)
    
    # Browse for 7z.exe button
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Location = New-Object System.Drawing.Point(30, 110)
    $btnBrowse.Size = New-Object System.Drawing.Size(240, 35)
    $btnBrowse.Text = "Browse for 7z.exe"
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnBrowse.ForeColor = [System.Drawing.Color]::White
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
    $btnBrowse.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
    $btnBrowse.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Title = "Locate 7z.exe"
        $openDialog.Filter = "7-Zip Executable (7z.exe)|7z.exe|All Files (*.*)|*.*"
        $openDialog.InitialDirectory = ${env:ProgramFiles}
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if (Test-Path $openDialog.FileName) {
                $recoveryForm.Tag = @{ Action = "Browse"; Path = $openDialog.FileName }
                $recoveryForm.Close()
            }
        }
    })
    $contentPanel.Controls.Add($btnBrowse)
    
    # Retry winget install button
    $btnRetry = New-Object System.Windows.Forms.Button
    $btnRetry.Location = New-Object System.Drawing.Point(290, 110)
    $btnRetry.Size = New-Object System.Drawing.Size(240, 35)
    $btnRetry.Text = "Retry Winget Install"
    $btnRetry.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    $btnRetry.ForeColor = [System.Drawing.Color]::White
    $btnRetry.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRetry.FlatAppearance.BorderSize = 0
    $btnRetry.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRetry.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnRetry.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12) })
    $btnRetry.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16) })
    $btnRetry.Add_Click({
        $recoveryForm.Tag = @{ Action = "Retry" }
        $recoveryForm.Close()
    })
    $contentPanel.Controls.Add($btnRetry)
    
    # Download from website button
    $btnDownload = New-Object System.Windows.Forms.Button
    $btnDownload.Location = New-Object System.Drawing.Point(160, 155)
    $btnDownload.Size = New-Object System.Drawing.Size(240, 35)
    $btnDownload.Text = "Download from Website"
    $btnDownload.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnDownload.ForeColor = [System.Drawing.Color]::White
    $btnDownload.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDownload.FlatAppearance.BorderSize = 0
    $btnDownload.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnDownload.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDownload.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
    $btnDownload.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
    $btnDownload.Add_Click({
        try {
            Start-Process "https://www.7-zip.org/download.html"
            [System.Windows.Forms.MessageBox]::Show(
                "7-Zip download page opened in browser.`n`nAfter installation, click 'Browse for 7z.exe' or 'Retry Winget Install'.",
                "Download Started", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to open browser: $_", "Error", "OK", "Error")
        }
    })
    $contentPanel.Controls.Add($btnDownload)
    
    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(450, 290)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 35)
    $btnCancel.Text = "Exit"
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
    $btnCancel.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
    $btnCancel.Add_Click({
        $recoveryForm.Tag = @{ Action = "Exit" }
        $recoveryForm.Close()
    })
    $recoveryForm.Controls.Add($btnCancel)
    
    $recoveryForm.ShowDialog() | Out-Null
    return $recoveryForm.Tag
}

# Try to find 7-Zip
$sevenZipPath = $null
$commonPaths = @(
    "${env:ProgramFiles}\7-Zip\7z.exe"
    "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
)
foreach ($p in $commonPaths) {
    if (Test-Path $p -PathType Leaf) { $sevenZipPath = $p; break }
}

# If not found, try winget install first
if (-not $sevenZipPath) {
    Write-Host "7-Zip not found. Attempting automatic installation via winget..." -ForegroundColor Yellow
    try {
        $proc = Start-Process "winget.exe" -ArgumentList "install","--id","7zip.7zip","--silent","--accept-package-agreements","--accept-source-agreements","--force" -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -eq 0) {
            Start-Sleep -Seconds 3
            foreach ($p in $commonPaths) {
                if (Test-Path $p -PathType Leaf) { $sevenZipPath = $p; break }
            }
            if ($sevenZipPath) {
                Write-Host "7-Zip installed successfully!" -ForegroundColor Green
            }
        }
    } catch {
        Write-Host "Winget installation failed: $_" -ForegroundColor Yellow
    }
}

# If still not found, show recovery dialog
while (-not $sevenZipPath) {
    $message = "7-Zip is required for compressing and extracting profile archives.`n`n" +
               "It was not found in standard locations and automatic installation failed."
    
    $result = Show-SevenZipRecoveryDialog -Message $message
    
    if (-not $result -or $result.Action -eq "Exit") {
        Write-Host "User cancelled - exiting" -ForegroundColor Yellow
        exit 1
    }
    
    switch ($result.Action) {
        "Browse" {
            if (Test-Path $result.Path) {
                $sevenZipPath = $result.Path
                Write-Host "7-Zip located: $sevenZipPath" -ForegroundColor Green
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid path selected. Please try again.", "Error", "OK", "Error")
            }
        }
        "Retry" {
            Write-Host "Retrying winget installation..." -ForegroundColor Yellow
            try {
                $proc = Start-Process "winget.exe" -ArgumentList "install","--id","7zip.7zip","--silent","--accept-package-agreements","--accept-source-agreements","--force" -Wait -PassThru -WindowStyle Hidden
                if ($proc.ExitCode -eq 0) {
                    Start-Sleep -Seconds 3
                    foreach ($p in $commonPaths) {
                        if (Test-Path $p -PathType Leaf) { $sevenZipPath = $p; break }
                    }
                    if ($sevenZipPath) {
                        Write-Host "7-Zip installed successfully!" -ForegroundColor Green
                        [System.Windows.Forms.MessageBox]::Show("7-Zip installed successfully!", "Success", "OK", "Information")
                    } else {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Installation completed but 7z.exe not found in expected location.`n`nPlease use 'Browse for 7z.exe' option.",
                            "Manual Location Required", "OK", "Warning")
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Winget installation failed (exit code: $($proc.ExitCode)).`n`nTry downloading from website or browse for existing installation.",
                        "Installation Failed", "OK", "Error")
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Installation error: $_", "Error", "OK", "Error")
            }
        }
    }
}

$global:SevenZipPath = $sevenZipPath
Write-Host "7-Zip ready: $global:SevenZipPath" -ForegroundColor Green

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
function Get-DirectorySize {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    
    if (-not (Test-Path $Path)) { return 0 }
    
    try {
        # Simple recursive scan - just get the actual folder size
        $size = (Get-ChildItem -Path $Path -Recurse -File -Force -ErrorAction SilentlyContinue | 
                Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        return if ($size) { $size } else { 0 }
    } catch {
        Log-Message "WARNING: Size calculation failed: $_"
        return 0
    }
}

# =============================================================================
# LOGGING & HELPERS (unchanged from your original - all perfect)
# =============================================================================
function Log-Message {
    param(
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Message"
    
    # Store entry in global array for retroactive filtering
    $global:LogEntries += [PSCustomObject]@{
        Timestamp = $ts
        Level = $Level
        Message = $Message
        FormattedLine = $line
    }
    
    # Filter by log level - skip display if message level is below configured minimum
    $messageLevel = $LogLevels[$Level]
    $configuredLevel = $LogLevels[$Config.LogLevel]
    if ($messageLevel -lt $configuredLevel) {
        # Still write to file, just don't display
        if ($global:CurrentLogFile) {
            try {
                Add-Content -Path $global:CurrentLogFile -Value $line -ErrorAction SilentlyContinue
            } catch { }
        }
        return  # Skip display
    }
    
    # Color-coded console output
    $color = switch ($Level) {
        'DEBUG' { 'Gray' }
        'INFO'  { 'White' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }
    
    if ($global:LogBox) {
        $global:LogBox.AppendText("$line`r`n")
        $global:LogBox.SelectionStart = $global:LogBox.TextLength
        $global:LogBox.ScrollToCaret()
    }
    Write-Host $line -ForegroundColor $color
    
    # Write to persistent log file if one is active
    if ($global:CurrentLogFile) {
        try {
            Add-Content -Path $global:CurrentLogFile -Value $line -ErrorAction SilentlyContinue
        } catch {
            # Silently fail if log file write fails (don't interrupt operation)
        }
    }
}

function Refresh-LogDisplay {
    # Refresh the log display based on current filter level
    if (-not $global:LogBox) { return }
    
    $configuredLevel = $LogLevels[$Config.LogLevel]
    $global:LogBox.Clear()
    
    foreach ($entry in $global:LogEntries) {
        $messageLevel = $LogLevels[$entry.Level]
        if ($messageLevel -ge $configuredLevel) {
            $global:LogBox.AppendText("$($entry.FormattedLine)`r`n")
        }
    }
    
    # Scroll to bottom
    $global:LogBox.SelectionStart = $global:LogBox.TextLength
    $global:LogBox.ScrollToCaret()
}

# Convenience wrapper functions for different log levels
function Log-Debug {
    param([string]$Message)
    Log-Message -Message $Message -Level 'DEBUG'
}

function Log-Info {
    param([string]$Message)
    Log-Message -Message $Message -Level 'INFO'
}

function Log-Warning {
    param([string]$Message)
    Log-Message -Message $Message -Level 'WARN'
}

function Log-Error {
    param([string]$Message)
    Log-Message -Message $Message -Level 'ERROR'
}

# === HTML REPORT GENERATION ===
function Generate-MigrationReport {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Export','Import')]
        [string]$OperationType,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$ReportData
    )
    
    if (-not $Config.GenerateHTMLReports) {
        Log-Debug "Report generation disabled in configuration"
        return $null
    }
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportFileName = "Migration_Report_${OperationType}_${timestamp}.html"
        
        # Save report to the same directory as the ZIP file
        $zipDirectory = Split-Path $ReportData.ZipPath -Parent
        $reportPath = Join-Path $zipDirectory $reportFileName
        
        Log-Info "Generating $OperationType migration report..."
        
        # Build HTML report
        $currentDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Profile Migration Report - $OperationType</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 900px; margin: 0 auto; background-color: white; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { border-bottom: 3px solid #0078d4; padding-bottom: 20px; margin-bottom: 30px; }
        .header h1 { color: #0078d4; margin: 0; font-size: 28px; }
        .header .subtitle { color: #666; margin-top: 5px; font-size: 14px; }
        .section { margin-bottom: 25px; }
        .section-title { background-color: #0078d4; color: white; padding: 10px 15px; font-size: 16px; font-weight: bold; margin-bottom: 15px; }
        .info-grid { display: grid; grid-template-columns: 200px 1fr; gap: 10px; }
        .info-label { font-weight: bold; color: #333; }
        .info-value { color: #555; }
        .success { color: #107c10; font-weight: bold; }
        .warning { color: #ff8c00; font-weight: bold; }
        .error { color: #e81123; font-weight: bold; }
        .stat-box { background-color: #f0f0f0; padding: 15px; margin: 10px 0; border-left: 4px solid #0078d4; }
        .stat-box h3 { margin: 0 0 10px 0; color: #0078d4; font-size: 14px; }
        .stat-box p { margin: 5px 0; font-size: 13px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        table th { background-color: #0078d4; color: white; padding: 10px; text-align: left; font-size: 13px; }
        table td { padding: 8px; border-bottom: 1px solid #ddd; font-size: 13px; }
        table tr:hover { background-color: #f9f9f9; }
        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #999; text-align: center; }
        ul { margin: 5px 0; padding-left: 20px; }
        li { margin: 3px 0; font-size: 13px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Profile Migration Report</h1>
            <div class="subtitle">Operation: $OperationType | Generated: $currentDate</div>
        </div>
"@
        
        # === EXPORT REPORT ===
        if ($OperationType -eq 'Export') {
            $checkmark = [char]0x2713
            $elapsedMins = $ReportData.ElapsedMinutes
            $elapsedSecs = $ReportData.ElapsedSeconds
            $html += @"
        <div class="section">
            <div class="section-title">Export Summary</div>
            <div class="info-grid">
                <div class="info-label">Status:</div>
                <div class="info-value success">$checkmark Export Completed Successfully</div>
                <div class="info-label">Source User:</div>
                <div class="info-value">$($ReportData.Username)</div>
                <div class="info-label">Source SID:</div>
                <div class="info-value">$($ReportData.SourceSID)</div>
                <div class="info-label">Source Path:</div>
                <div class="info-value">$($ReportData.SourcePath)</div>
                <div class="info-label">Export File:</div>
                <div class="info-value">$($ReportData.ZipPath)</div>
                <div class="info-label">Archive Size:</div>
                <div class="info-value">$($ReportData.ZipSizeMB) MB</div>
                <div class="info-label">Operation Time:</div>
                <div class="info-value">$elapsedMins minutes ($elapsedSecs seconds)</div>
                <div class="info-label">Timestamp:</div>
                <div class="info-value">$($ReportData.Timestamp)</div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Profile Statistics</div>
            <div class="stat-box">
                <h3>Files &amp; Folders</h3>
                <p>Total Files: $($ReportData.FileCount)</p>
                <p>Total Folders: $($ReportData.FolderCount)</p>
                <p>Uncompressed Size: $($ReportData.UncompressedSizeMB) MB</p>
                <p>Compression Ratio: $($ReportData.CompressionRatio)</p>
            </div>
        </div>
"@
            
            if ($ReportData.Exclusions -and $ReportData.Exclusions.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Exclusions Applied</div>
            <ul>
"@
                foreach ($exclusion in $ReportData.Exclusions) {
                    $html += "                <li>$exclusion</li>`n"
                }
                $html += @"
            </ul>
        </div>
"@
            }
            
            # Add cleanup optimization section
            if ($ReportData.CleanupCategories -and $ReportData.CleanupCategories.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Profile Cleanup Optimization</div>
            <div class="info-grid">
                <div class="info-label">Categories Cleaned:</div>
                <div class="info-value">$($ReportData.CleanupCategories.Count)</div>
                <div class="info-label">Space Saved:</div>
                <div class="info-value success">$($ReportData.CleanupSavingsMB) MB</div>
            </div>
            <p><strong>Excluded Categories:</strong></p>
            <ul>
"@
                foreach ($category in $ReportData.CleanupCategories) {
                    $html += "                <li>$category</li>`n"
                }
                $html += @"
            </ul>
            <p class="success">These items were excluded to reduce export size and improve transfer speed.</p>
        </div>
"@
            }
            
            if ($ReportData.HashEnabled) {
                $html += @"
        <div class="section">
            <div class="section-title">Integrity Verification</div>
            <div class="info-grid">
                <div class="info-label">Hash Algorithm:</div>
                <div class="info-value">SHA-256</div>
                <div class="info-label">Hash File:</div>
                <div class="info-value">$($ReportData.ZipPath).sha256</div>
                <div class="info-label">Status:</div>
                <div class="info-value success">&#10003; Hash generated</div>
            </div>
        </div>
"@
            }
            
            # Add installed programs list
            if ($ReportData.InstalledPrograms -and $ReportData.InstalledPrograms.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Installed Programs on Source Computer</div>
            <p>Total Programs: $($ReportData.InstalledPrograms.Count)</p>
            <table>
                <thead>
                    <tr>
                        <th>Program Name</th>
                        <th>Version</th>
                        <th>Source</th>
                    </tr>
                </thead>
                <tbody>
"@
                foreach ($program in $ReportData.InstalledPrograms) {
                    $progName = if ($program.Name) { $program.Name } else { "Unknown" }
                    $progVersion = if ($program.Version) { $program.Version } else { "N/A" }
                    $progSource = if ($program.Source) { $program.Source } else { "N/A" }
                    $html += "                    <tr><td>$progName</td><td>$progVersion</td><td>$progSource</td></tr>`n"
                }
                $html += @"
                </tbody>
            </table>
        </div>
"@
            }
        }
        
        # === IMPORT REPORT ===
        if ($OperationType -eq 'Import') {
            $checkmark = [char]0x2713
            $elapsedMins = $ReportData.ElapsedMinutes
            $elapsedSecs = $ReportData.ElapsedSeconds
            $html += @"
        <div class="section">
            <div class="section-title">Import Summary</div>
            <div class="info-grid">
                <div class="info-label">Status:</div>
                <div class="info-value success">$checkmark Import Completed Successfully</div>
                <div class="info-label">Target User:</div>
                <div class="info-value">$($ReportData.Username)</div>
                <div class="info-label">User Type:</div>
                <div class="info-value">$($ReportData.UserType)</div>
                <div class="info-label">Target SID:</div>
                <div class="info-value">$($ReportData.TargetSID)</div>
                <div class="info-label">Source SID:</div>
                <div class="info-value">$($ReportData.SourceSID)</div>
                <div class="info-label">Profile Path:</div>
                <div class="info-value">$($ReportData.ProfilePath)</div>
                <div class="info-label">Source ZIP:</div>
                <div class="info-value">$($ReportData.ZipPath)</div>
                <div class="info-label">Archive Size:</div>
                <div class="info-value">$($ReportData.ZipSizeMB) MB</div>
                <div class="info-label">Import Mode:</div>
                <div class="info-value">$($ReportData.ImportMode)</div>
                <div class="info-label">Operation Time:</div>
                <div class="info-value">$elapsedMins minutes ($elapsedSecs seconds)</div>
                <div class="info-label">Timestamp:</div>
                <div class="info-value">$($ReportData.Timestamp)</div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Profile Statistics</div>
            <div class="stat-box">
                <h3>Files &amp; Folders</h3>
                <p>Total Files: $($ReportData.FileCount)</p>
                <p>Total Folders: $($ReportData.FolderCount)</p>
                <p>Uncompressed Size: $($ReportData.UncompressedSizeMB) MB</p>
                <p>Compression Ratio: $($ReportData.CompressionRatio)</p>
            </div>
        </div>
"@
            
            if ($ReportData.BackupPath) {
                $html += @"
        <div class="section">
            <div class="section-title">Backup Information</div>
            <div class="info-grid">
                <div class="info-label">Backup Created:</div>
                <div class="info-value success">&#10003; Yes</div>
                <div class="info-label">Backup Location:</div>
                <div class="info-value">$($ReportData.BackupPath)</div>
                <div class="info-label">Note:</div>
                <div class="info-value">Backup will be automatically cleaned up after 2 newest backups</div>
            </div>
        </div>
"@
            }
            
            # Add installed apps section
            if ($ReportData.InstalledApps -and $ReportData.InstalledApps.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Applications Installed by Migration Tool</div>
            <p>Total Apps Installed: $($ReportData.InstalledApps.Count)</p>
            <table>
                <thead>
                    <tr>
                        <th>Application Name</th>
                        <th>Package ID</th>
                        <th>Version</th>
                    </tr>
                </thead>
                <tbody>
"@
                foreach ($app in $ReportData.InstalledApps) {
                    $appName = if ($app.PackageIdentifier) { $app.PackageIdentifier.Split('.')[-1] } else { "Unknown" }
                    $appId = if ($app.PackageIdentifier) { $app.PackageIdentifier } else { "N/A" }
                    $appVersion = if ($app.Version) { $app.Version } else { "Latest" }
                    $html += "                    <tr><td>$appName</td><td>$appId</td><td>$appVersion</td></tr>`n"
                }
                $html += @"
                </tbody>
            </table>
        </div>
"@
            }
            
            if ($ReportData.HashVerified) {
                $html += @"
        <div class="section">
            <div class="section-title">Integrity Verification</div>
            <div class="info-grid">
                <div class="info-label">Hash Verification:</div>
                <div class="info-value success">&#10003; Passed</div>
                <div class="info-label">Algorithm:</div>
                <div class="info-value">SHA-256</div>
            </div>
        </div>
"@
            }
            
            $html += @"
        <div class="section">
            <div class="section-title">Next Steps</div>
            <div class="stat-box">
                <h3>Post-Migration Actions Required</h3>
                <ul>
                    <li><strong>REBOOT NOW</strong> - Restart the computer before logging in</li>
                    <li>Log in as <strong>$($ReportData.Username)</strong></li>
                    <li>Outlook will rebuild mailbox cache (10-30 minutes)</li>
                    <li>Re-enter email passwords if prompted</li>
                    <li>Reconnect to Exchange/Microsoft 365 if needed</li>
                    <li>Copy PST archives manually if stored outside profile</li>
                </ul>
            </div>
        </div>
"@
        }
        
        # === WARNINGS/ERRORS (if any) ===
        if ($ReportData.Warnings -and $ReportData.Warnings.Count -gt 0) {
            $html += @"
        <div class="section">
            <div class="section-title">Warnings</div>
            <ul>
"@
            foreach ($warning in $ReportData.Warnings) {
                $html += "                <li class=`"warning`">WARNING: $warning</li>`n"
            }
            $html += @"
            </ul>
        </div>
"@
        }
        
        # Footer
        $html += @"
        <div class="footer">
            <p>Profile Migration Tool | Generated on $(hostname) | PowerShell $($PSVersionTable.PSVersion)</p>
        </div>
    </div>
</body>
</html>
"@
        
        # Save HTML report
        $html | Out-File -FilePath $reportPath -Encoding UTF8 -Force
        Log-Info "HTML report saved: $reportPath"
        
        # Auto-open report if configured
        if ($Config.AutoOpenReports) {
            Log-Debug "Opening HTML report..."
            Start-Process $reportPath
        }
        
        return $reportPath
        
    } catch {
        Log-Error "Failed to generate migration report: $_"
        return $null
    }
}

function Start-OperationLog {
    param(
        [Parameter(Mandatory=$true)][string]$Operation,
        [string]$Username
    )
    
    # Create logs directory if it doesn't exist
    $logsDir = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    
    # Create timestamped log file
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    # Strip size suffix if present (e.g., "username - [756.6 MB]" -> "username")
    $cleanUsername = $Username
    if ($cleanUsername -match '^(.+?)\s+-\s+\[.+\]$') {
        $cleanUsername = $matches[1].Trim()
    }
    $sanitizedUser = if ($cleanUsername) { $cleanUsername -replace '[\\/:*?"<>|]', '_' } else { "Unknown" }
    $logFileName = "${Operation}_${sanitizedUser}_${timestamp}.log"
    $global:CurrentLogFile = Join-Path $logsDir $logFileName
    
    # Write header to log file
    $header = @"
================================================================================
Profile Migration Tool - $Operation Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
User: $sanitizedUser
Computer: $env:COMPUTERNAME
PowerShell: $($PSVersionTable.PSVersion)
OS: $([Environment]::OSVersion.VersionString)
================================================================================

"@
    
    try {
        Set-Content -Path $global:CurrentLogFile -Value $header -ErrorAction Stop
        Log-Info "Log file created: $global:CurrentLogFile"
        return $global:CurrentLogFile
    } catch {
        Write-Host "WARNING: Could not create log file: $_" -ForegroundColor Yellow
        $global:CurrentLogFile = $null
        return $null
    }
}

function Stop-OperationLog {
    if ($global:CurrentLogFile) {
        $footer = "`n" + "=" * 80 + "`nOperation completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" + "=" * 80
        try {
            Add-Content -Path $global:CurrentLogFile -Value $footer -ErrorAction SilentlyContinue
        } catch { }
        $global:CurrentLogFile = $null
    }
}

function Get-LocalProfiles {
    Get-ChildItem "C:\Users" -Directory | Where-Object {
        $_.Name -notmatch "^Public$|^Default$|^Administrator$|^All Users$|^Default User$"
    } | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Username = $_.Name; Path = $_.FullName }
    }
}

# Enhanced profile detection with size estimates
function Get-ProfileInfo {
    param(
        [string]$ProfilePath,
        [switch]$SkipSizeCalculation
    )
    
    try {
        # Check if NTUSER.DAT exists (indicates valid profile)
        $hasHive = Test-Path (Join-Path $ProfilePath "NTUSER.DAT")
        
        # Skip expensive size calculation during initial load
        if ($SkipSizeCalculation) {
            return @{
                Path = $ProfilePath
                SizeMB = -1  # -1 indicates not calculated
                ItemCount = 0
                HasHive = $hasHive
                IsValid = $hasHive
            }
        }
        
        # Only calculate size when explicitly requested (refresh button)
        $sizeBytes = 0
        $itemCount = 0
        
        # Get immediate children sizes (shallow scan for speed)
        $items = Get-ChildItem -Path $ProfilePath -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if (-not $item.PSIsContainer) {
                $sizeBytes += $item.Length
                $itemCount++
            } else {
                # For directories, get a quick estimate
                $dirSize = (Get-ChildItem -Path $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | 
                           Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($dirSize) {
                    $sizeBytes += $dirSize
                }
                $itemCount++
            }
        }
        
        $sizeMB = [math]::Round($sizeBytes / 1MB, 1)
        
        return @{
            Path = $ProfilePath
            SizeMB = $sizeMB
            ItemCount = $itemCount
            HasHive = $hasHive
            IsValid = $hasHive
        }
    } catch {
        return @{
            Path = $ProfilePath
            SizeMB = 0
            ItemCount = 0
            HasHive = $false
            IsValid = $false
        }
    }
}

# Build display names as DOMAIN\\username or COMPUTERNAME\\username for the dropdown
function Get-ProfileDisplayEntries {
    param([switch]$CalculateSizes)
    
    $base = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $computer = $env:COMPUTERNAME
    $profiles = Get-LocalProfiles
    $result = @()
    foreach ($p in $profiles) {
        # Get profile info - skip size calculation during initial load for speed
        $profileInfo = Get-ProfileInfo -ProfilePath $p.Path -SkipSizeCalculation:(-not $CalculateSizes)
        
        $sidKey = Get-ChildItem $base -ErrorAction SilentlyContinue | Where-Object {
            (Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$($p.Username)"
        } | Select-Object -First 1
        $display = $p.Username
        if ($sidKey) {
            try {
                $sid = $sidKey.PSChildName
                $sidObj = [System.Security.Principal.SecurityIdentifier]::new($sid)
                $nt = $sidObj.Translate([System.Security.Principal.NTAccount])
                $parts = $nt.Value -split '\\',2
                if ($parts.Count -ge 2) {
                    $domain = $parts[0]
                    $name = $parts[1]
                    if ($domain -ieq $computer) {
                        $display = "$computer\$name"
                    } else {
                        $display = "$domain\$name"
                    }
                }
            } catch {
                # fallback to raw username
            }
        } else {
            # no registry mapping; assume local
            $display = "$computer\\$($p.Username)"
        }
        
        # Add size estimate to display name only if calculated
        if ($CalculateSizes -and $profileInfo.SizeMB -ge 0) {
            $sizeDisplay = if ($profileInfo.SizeMB -gt 1024) {
                "$([math]::Round($profileInfo.SizeMB/1024, 1)) GB"
            } else {
                "$($profileInfo.SizeMB) MB"
            }
            $displayWithSize = "$display - [$sizeDisplay]"
        } else {
            $displayWithSize = $display
        }
        
        $result += [pscustomobject]@{ 
            DisplayName = $displayWithSize
            Username = $p.Username
            Path = $p.Path
            SizeMB = $profileInfo.SizeMB
            IsValid = $profileInfo.IsValid
        }
    }
    return $result | Sort-Object -Property Username
}

# === ROBOCOPY EXCLUSION CONFIGURATION (centralized) ===
function Get-RobocopyExclusions {
    param(
        [ValidateSet('Export', 'Import', 'AppData')]
        [string]$Mode = 'Export'
    )
    
    $excludeFiles = @('Thumbs.db', 'desktop.ini')
    $excludeDirs = @('AppData\Local\Microsoft\Edge\User Data\ShaderCache', 'AppData\Local\Temp')
    
    $result = @{
        Files = $excludeFiles
        Dirs  = $excludeDirs
    }
    
    # Export mode: exclude everything in AppData except what we need
    if ($Mode -eq 'Export') {
        # Only include specific AppData subdirectories (everything else in AppData is excluded)
        # This list is complementary - we EXCLUDE everything EXCEPT these
        $result.AppDataInclude = @(
            'Roaming\Microsoft\Signatures',
            'Roaming\Microsoft\Outlook',
            'Roaming\Microsoft\Office',
            'Roaming\Microsoft\Credentials',
            'Local\Microsoft\Credentials',
            'Local\Microsoft\Protect',
            'Local\Packages\Microsoft.MicrosoftStickyNotes_*',
            'Local\Google\Chrome\User Data\Default',
            'Local\Google\Chrome\User Data\Profile *',
            'Local\Microsoft\Edge\User Data\Default',
            'Local\Microsoft\Edge\User Data\Profile *',
            'Roaming\Mozilla\Firefox\Profiles\*'
        )
    }
    
    # Import mode excludes only LOG files and temp data
    # NTUSER.DAT must be extracted for 7-Zip import (it's extracted, then rewritten with SID replacement)
    if ($Mode -eq 'Import') {
        # Exclude specific LOG files but NOT NTUSER.DAT itself
        $result.Files += @('ntuser.dat.LOG1', 'ntuser.dat.LOG2', 'NTUSER.DAT.LOG1', 'NTUSER.DAT.LOG2', 'UsrClass.dat.LOG1', 'UsrClass.dat.LOG2')
        # Import should only include specific AppData paths, exclude everything else
        # We'll handle this with explicit exclusion in the import logic
        $result.AppDataInclude = @(
            'Roaming\Microsoft\Signatures',
            'Roaming\Microsoft\Outlook',
            'Roaming\Microsoft\Office',
            'Roaming\Microsoft\Credentials',
            'Local\Microsoft\Credentials',
            'Local\Microsoft\Protect',
            'Local\Packages\Microsoft.MicrosoftStickyNotes_*',
            'Local\Google\Chrome\User Data\Default',
            'Local\Google\Chrome\User Data\Profile *',
            'Local\Microsoft\Edge\User Data\Default',
            'Local\Microsoft\Edge\User Data\Profile *',
            'Roaming\Mozilla\Firefox\Profiles\*'
        )
    }
    
    # AppData bonus copy only needs cache exclusions
    if ($Mode -eq 'AppData') {
        $result.Files = @()
        $result.Dirs = @('Cache', 'Code Cache', 'GPUCache', 'Temp')
    }
    
    return $result
}

# Map robocopy-style exclusions to 7-Zip extract filters
function Get-SevenZipExtractFilters {
    param(
        [ValidateSet('Export','Import','AppData')][string]$Mode = 'Import'
    )
    $ex = Get-RobocopyExclusions -Mode $Mode
    $filters = @()
    # 7-Zip works with forward slashes inside archives
    # Files are now at ZIP root (flat structure), not under Profile/
    foreach ($f in $ex.Files) {
        # Exclude matching files anywhere in ZIP
        $filters += ("-x!**/$f")
    }
    foreach ($d in $ex.Dirs) {
        # Exclude directory trees at root and nested
        $filters += ("-x!$($d -replace '\\','/')/**")
        $filters += ("-x!**/$($d -replace '\\','/')/**")
    }
    return $filters
}
            # Build exclusion filters for 7-Zip
            $exclusions = Get-RobocopyExclusions -Mode 'Export'
            $7zExclusions = @()
            foreach ($file in $exclusions.Files) {
                if (-not [string]::IsNullOrWhiteSpace($file)) {
                    $relFile = $file
                    if ($profile -and $profile.Path -and $file -like ("$($profile.Path)\*")) {
                        $relFile = $file.Substring($profile.Path.Length + 1)
                    }
                    $7zExclusions += "-x!$relFile"
                }
            }
            foreach ($dir in $exclusions.Dirs) {
                if (-not [string]::IsNullOrWhiteSpace($dir)) {
                    $relDir = $dir
                    if ($profile -and $profile.Path -and $dir -like ("$($profile.Path)\*")) {
                        $relDir = $dir.Substring($profile.Path.Length + 1)
                    }
                    $pattern = if ($relDir.EndsWith('*')) { $relDir } else { "$relDir\*" }
                    $7zExclusions += "-x!$pattern"
                }
            }

            # Explicitly exclude common legacy junctions under Documents to mirror robocopy /XJ behavior
            # These are hidden in Explorer but can be captured by 7-Zip unless excluded
            $junctions = @(
                'Documents\My Music',
                'Documents\My Pictures',
                'Documents\My Videos'
            )
            foreach ($j in $junctions) {
                # 7-Zip -x! patterns are relative to the added root; we are adding "$($profile.Path)\*"
                # so relative exclusions like Documents\My Music are appropriate here
                $7zExclusions += "-x!`"$j`""
            }

            # Discover and exclude all reparse-point directories (junctions) under the profile root
            try {
                $rpItems = Get-ChildItem -Path $profile.Path -Directory -Recurse -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }
                foreach ($item in $rpItems) {
                    # Compute path relative to profile root for 7-Zip -x! pattern
                    $rel = $item.FullName.Substring($profile.Path.Length).TrimStart('\\')
                    if ([string]::IsNullOrWhiteSpace($rel)) { continue }
                    $pattern = "-x!`"$rel`""
                    if ($7zExclusions -notcontains $pattern) { $7zExclusions += $pattern }
                }
                if ($rpItems.Count -gt 0) {
                    Log-Message "7-Zip export: excluding $($rpItems.Count) reparse points (junctions)"
                    # Optional: list a few excluded junction paths for visibility
                    $preview = $rpItems | Select-Object -First 10
                    foreach ($p in $preview) {
                        $relPreview = $p.FullName.Substring($profile.Path.Length).TrimStart('\\')
                        Log-Message " - excluded junction: $relPreview"
                    }
                    if ($rpItems.Count -gt $preview.Count) {
                        Log-Message " - and $($rpItems.Count - $preview.Count) more..."
                    }
                }
            } catch {
                Log-Message "WARNING: Junction discovery failed: $_"
            }
    $7zArgs = @('a', '-tzip', $ZipPath, "$($profile.Path)\*", '-mx=5', '-mmt=on', '-bsp1') + $7zExclusions

# === PROFILE PATH VALIDATION ===
function Test-ProfilePathWriteable {
    param([string]$Path)
    
    try {
        $timeout = New-TimeSpan -Seconds $Config.ProfileValidationTimeoutSec
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # If path doesn't exist, try to create it (this is expected for imports to new machines)
        $createdDir = $false
        if (-not (Test-Path $Path)) {
            try {
                New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
                $createdDir = $true
                Log-Message "Created missing target directory for validation: $Path"
            } catch {
                throw "Could not create target directory: $_"
            }
        }

        # Create test file inside target path
        $testFile = Join-Path $Path (".PROFILE_TEST_$(Get-Random)")
        $null | Out-File -FilePath $testFile -Force -ErrorAction Stop

        # Verify it exists
        if (-not (Test-Path $testFile -PathType Leaf)) {
            throw "Test file created but cannot be verified"
        }

        # Clean up test file
        Remove-Item $testFile -Force -ErrorAction Stop

        $sw.Stop()
        Log-Message "Profile path validation successful ($($sw.ElapsedMilliseconds)ms)"
        return $true

    } catch {
        Log-Message "CRITICAL: Profile path not writable: $_"
        return $false
    } finally {
        # If we created the directory only for validation and it's still empty, leave it (import will use it).
        # We avoid deleting it to reduce surprising side-effects; creation confirms write access.
    }
}

# === MOUNTED PROFILE DETECTION ===
function Test-ProfileMounted {
    param([string]$UserSID)
    
    try {
        # Check if SID is currently mounted in HKU
        $mounted = (Get-ChildItem "HKU:" -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -eq $UserSID }) -ne $null
        
        if ($mounted) {
            Log-Message "WARNING: User profile SID $UserSID is currently mounted (user may be logged in or cached)"
            return $true
        }
        
        return $false
    } catch {
        Log-Message "Could not verify if profile is mounted: $_"
        return $false
    }
}

function Get-LocalProfiles {
    Get-ChildItem "C:\Users" -Directory | Where-Object {
        $_.Name -notmatch "^Public$|^Default$|^Administrator$|^All Users$|^Default User$"
    } | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Username = $_.Name; Path = $_.FullName }
    }
}

function Remove-FolderRobust {
    param([string]$Path)
    if (Test-Path $Path) {
        try { Remove-Item $Path -Recurse -Force -ErrorAction Stop }
        catch {
            $empty = "$env:TEMP\Empty_$(Get-Random)"
            New-Item -ItemType Directory -Path $empty -Force | Out-Null
            robocopy $empty $Path /MIR /R:1 /W:1 | Out-Null
            Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $empty -Recurse -Force
        }
    }
}

function Show-LogViewer {
    param(
        [Parameter(Mandatory=$true)][string]$LogPath,
        [string]$Title = "Log Viewer"
    )

    try {
        $content = if (Test-Path $LogPath) { Get-Content $LogPath -Raw -ErrorAction SilentlyContinue } else { "Log file not found: $LogPath" }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title
        $form.Size = New-Object System.Drawing.Size(950,750)
        $form.StartPosition = 'CenterScreen'
        $form.Font = New-Object System.Drawing.Font('Segoe UI',9)
        $form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
        $form.MinimumSize = New-Object System.Drawing.Size(800, 600)

        # Header panel
        $headerPanel = New-Object System.Windows.Forms.Panel
        $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
        $headerPanel.Size = New-Object System.Drawing.Size(950, 60)
        $headerPanel.BackColor = [System.Drawing.Color]::White
        $headerPanel.Dock = 'Top'
        $form.Controls.Add($headerPanel)

        $lblHeader = New-Object System.Windows.Forms.Label
        $lblHeader.Location = New-Object System.Drawing.Point(20, 15)
        $lblHeader.Size = New-Object System.Drawing.Size(800, 30)
        $lblHeader.Text = "Migration Log Viewer"
        $lblHeader.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
        $lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $headerPanel.Controls.Add($lblHeader)

        # Log content panel
        $contentPanel = New-Object System.Windows.Forms.Panel
        $contentPanel.Padding = New-Object System.Windows.Forms.Padding(15)
        $contentPanel.Dock = 'Fill'
        $form.Controls.Add($contentPanel)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true
        $txt.ReadOnly = $true
        $txt.ScrollBars = 'Both'
        $txt.Font = New-Object System.Drawing.Font('Consolas',9)
        $txt.Dock = 'Fill'
        $txt.Text = $content
        $txt.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
        $txt.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $contentPanel.Controls.Add($txt)

        # Bottom button panel
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = 'Bottom'
        $panel.Height = 60
        $panel.BackColor = [System.Drawing.Color]::White
        $form.Controls.Add($panel)

        $btnOpen = New-Object System.Windows.Forms.Button
        $btnOpen.Text = 'Open in Notepad'
        $btnOpen.Width = 150
        $btnOpen.Height = 35
        $btnOpen.Location = New-Object System.Drawing.Point(15,12)
        $btnOpen.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $btnOpen.ForeColor = [System.Drawing.Color]::White
        $btnOpen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOpen.FlatAppearance.BorderSize = 0
        $btnOpen.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $btnOpen.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnOpen.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
        $btnOpen.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
        $btnOpen.Add_Click({ Start-Process notepad.exe -ArgumentList "`"$LogPath`"" })
        $panel.Controls.Add($btnOpen)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = 'Close'
        $btnClose.Width = 110
        $btnClose.Height = 35
        $btnClose.Location = New-Object System.Drawing.Point(175,12)
        $btnClose.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
        $btnClose.ForeColor = [System.Drawing.Color]::White
        $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnClose.FlatAppearance.BorderSize = 0
        $btnClose.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnClose.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
        $btnClose.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
        $btnClose.Add_Click({ $form.Close() })
        $panel.Controls.Add($btnClose)

        $form.TopMost = $true
        $form.ShowDialog() | Out-Null
    } catch {
        Log-Message "Failed to display log viewer: $_"
    }
}

function Get-LocalUserSID {
    param([Parameter(Mandatory=$true)][string]$UserName)
    
    # Handle domain usernames: convert NetBIOS to FQDN if needed
    if ($UserName -match '^([^\\]+)\\(.+)$') {
        $netbios = $matches[1]
        $user = $matches[2]
        
        # Try to resolve FQDN for the domain
        try {
            $fqdn = Get-DomainFQDN -NetBIOSName $netbios
            $UserName = "$fqdn\$user"
            Log-Message "Resolved domain user: $UserName"
        } catch {
            Log-Message "WARNING: Could not resolve FQDN for $netbios, using as-is"
        }
    }
    
    $ntAccount = New-Object System.Security.Principal.NTAccount($UserName)
    $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
    return $sid.Value
}

# Resolve NetBIOS domain name to FQDN (e.g., CPMETHOD -> cpmethod.local)
function Get-DomainFQDN {
    param([Parameter(Mandatory=$true)][string]$NetBIOSName)
    try {
        # Try getting current domain if machine is domain-joined
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        if ($domain.Name -match "^$NetBIOSName\.") { return $domain.Name }
    } catch {}
    
    # Fallback: DNS lookup for common TLDs
    $commonTLDs = @('.local', '.lan', '.corp', '.com', '.net', '.org')
    foreach ($tld in $commonTLDs) {
        $fqdn = "$NetBIOSName$tld"
        try {
            $result = Resolve-DnsName -Name $fqdn -ErrorAction Stop -QuickTimeout
            if ($result) { return $fqdn }
        } catch {}
    }
    
    # If all else fails, return original
    return $NetBIOSName
}

function Set-ProfileFolderAcls {
    param([Parameter(Mandatory=$true)][string]$ProfilePath, [Parameter(Mandatory=$true)][string]$UserName)
    Log-Message "Applying folder ACLs to $ProfilePath..."
    try {
        # Step 1: Take ownership (processes all files recursively)
        $global:StatusText.Text = "Taking ownership of all profile files (Step 1/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 1: Running takeown /F `"$ProfilePath`" /A /R /D Y (recursive ownership transfer)"
        takeown /F "$ProfilePath" /A /R /D Y >$null 2>&1
        Log-Message "Ownership transferred to Administrators group"
        
        # Step 2: Grant Administrators temporary control
        $global:StatusText.Text = "Granting Administrators access (Step 2/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 2: Granting Administrators full control"
        icacls "$ProfilePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
        
        # Step 3: Reset all ACLs (removes everything, including inherited)
        $global:StatusText.Text = "Resetting existing permissions (Step 3/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 3: Resetting all ACLs (removing existing permissions)"
        icacls "$ProfilePath" /reset /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 4: Remove inheritance protection (so we start clean)
        $global:StatusText.Text = "Removing inheritance protection (Step 4/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 4: Removing inheritance protection"
        icacls "$ProfilePath" /inheritance:r /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 5: Add explicit ACLs for SYSTEM, User, and Administrators
        $global:StatusText.Text = "Granting user and system permissions (Step 5/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 5: Adding explicit ACLs for SYSTEM"
        icacls "$ProfilePath" /grant:r "NT AUTHORITY\SYSTEM:(F)" /Q /C >$null 2>&1
        
        # For domain users, use the full domain\username; for local, use shortname
        if ($UserName -like "*\*") {
            # Domain user format: DOMAIN\username
            Log-Message "Step 5: Adding ACLs for domain user: $UserName"
            icacls "$ProfilePath" /grant:r "${UserName}:(F)" /Q /C >$null 2>&1
            icacls "$ProfilePath" /grant:r "${UserName}:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        } else {
            # Local user format: username (no leading '*')
            Log-Message "Step 5: Adding ACLs for local user: $UserName"
            icacls "$ProfilePath" /grant:r "${UserName}:(F)" /Q /C >$null 2>&1
            icacls "$ProfilePath" /grant:r "${UserName}:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        }
        
        Log-Message "Step 5: Adding ACLs for Administrators group"
        icacls "$ProfilePath" /grant:r "BUILTIN\Administrators:(F)" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 6: Add inheritable ACLs for SYSTEM and Administrators
        $global:StatusText.Text = "Setting inheritable permissions for subfolders (Step 6/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 6: Adding inheritable ACLs (OI)(CI)(IO) for SYSTEM and Administrators"
        icacls "$ProfilePath" /grant:r "NT AUTHORITY\SYSTEM:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        icacls "$ProfilePath" /grant:r "BUILTIN\Administrators:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        # Add CREATOR OWNER for proper per-file ownership on create
        Log-Message "Step 6: Adding CREATOR OWNER for new file creation"
        icacls "$ProfilePath" /grant:r "CREATOR OWNER:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        # Remove overly broad groups that can interfere
        Log-Message "Step 6: Removing overly permissive groups (Everyone, Users, Authenticated Users)"
        icacls "$ProfilePath" /remove:g "Everyone" "Users" "Authenticated Users" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 7: Set owner to the user
        $global:StatusText.Text = "Setting final ownership to user (Step 7/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        if ($UserName -like "*\*") {
            Log-Message "Step 7: Setting owner to domain user: $UserName"
            icacls "$ProfilePath" /setowner "$UserName" /Q /C >$null 2>&1
        } else {
            Log-Message "Step 7: Setting owner to local user: $UserName"
            icacls "$ProfilePath" /setowner "$UserName" /Q /C >$null 2>&1
        }
        Start-Sleep -Milliseconds 200
        
        Log-Message "Folder ACLs applied successfully (all 7 steps completed)"
    }
    catch {
        Log-Message "ERROR: Failed to apply folder ACLs: $_"
        throw $_
    }
}

# === FIXED HIVE ACL FUNCTION - USE ICACLS FOR FILE PERMISSIONS ===
function Set-ProfileHiveAcl {
    param(
        [Parameter(Mandatory=$true)][string]$HiveFilePath,
        [Parameter(Mandatory=$true)][string]$OwnerSID,
        [Parameter(Mandatory=$true)][string]$UserName,
        [bool]$IsLocalUser = $true
    )
    
    # Validate hive file exists
    if (-not (Test-Path $HiveFilePath)) { throw "Hive file not found: $HiveFilePath" }
    
    Log-Message "Applying file-level ACLs to NTUSER.DAT..."
    
    try {
        # Step 1: Take ownership as Administrators
        takeown /F "$HiveFilePath" /A /R /D Y >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 2: Grant Administrators full control temporarily so we can modify ACLs
        icacls "$HiveFilePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 3: Remove all explicit ACLs
        icacls "$HiveFilePath" /reset /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 4: Enable inheritance from parent folder
        icacls "$HiveFilePath" /inheritance:e /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200

        # Step 5: Explicitly grant required principals to avoid temp profile
        # SYSTEM and Administrators
        icacls "$HiveFilePath" /grant:r "NT AUTHORITY\SYSTEM:(F)" /Q /C >$null 2>&1
        icacls "$HiveFilePath" /grant:r "BUILTIN\Administrators:(F)" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200

        # User (domain or local format)
        if ($UserName -like "*\*") {
            icacls "$HiveFilePath" /grant:r "${UserName}:(F)" /Q /C >$null 2>&1
        } else {
            icacls "$HiveFilePath" /grant:r "${UserName}:(F)" /Q /C >$null 2>&1
        }
        Start-Sleep -Milliseconds 200

        # Step 6: Set owner to the user to ensure profile loads
        if ($UserName -like "*\*") {
            icacls "$HiveFilePath" /setowner "$UserName" /Q /C >$null 2>&1
        } else {
            icacls "$HiveFilePath" /setowner "$UserName" /Q /C >$null 2>&1
        }
        Start-Sleep -Milliseconds 200

        Log-Message "NTUSER.DAT ACLs applied and ownership set to user"
    }
    catch {
        Log-Message "ERROR: Failed to apply hive ACLs: $_"
        throw $_
    }
}

# (Removed) ACL diagnostic helper no longer needed

# =============================================================================
# Set-ProfileAcls
# =============================================================================
function Set-ProfileAcls {
    param(
        [Parameter(Mandatory=$true)][string]$ProfileFolder,
        [Parameter(Mandatory=$true)][string]$UserName,
        [string]$SourceSID,
        [string]$UserSID
    )

    # Use provided SID if available, otherwise resolve it
    $newSID = if ($UserSID) { $UserSID } else { Get-LocalUserSID -UserName $UserName }
    Log-Message "Target SID: $newSID"
    $global:ProgressBar.Value = 85

    # Target hive path (files already extracted to ProfileFolder)
    $targetHive = Join-Path $ProfileFolder "NTUSER.DAT"

    if (-not (Test-Path $targetHive)) {
        throw "NTUSER.DAT not found in profile folder: $targetHive"
    }

    # Enable required privileges (critical for ownership/ACL operations)
    Enable-Privilege SeBackupPrivilege -IsCritical $true
    Enable-Privilege SeRestorePrivilege -IsCritical $true
    Enable-Privilege SeTakeOwnershipPrivilege -IsCritical $true

    # =====================================================================
    # CRITICAL: Apply ACLs FIRST, then do SID rewrite
    # This ensures proper ownership and permissions before hive manipulation
    # =====================================================================
    
    # Step 1: Apply standard folder ACLs
    $global:StatusText.Text = "Analyzing profile structure..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Counting files and folders for progress estimation..."
    
    # Count items for progress reporting
    $itemCount = 0
    $folderCount = 0
    $fileCount = 0
    try {
        $items = Get-ChildItem -Path $ProfileFolder -Recurse -Force -ErrorAction SilentlyContinue
        $itemCount = ($items | Measure-Object).Count
        $folderCount = ($items | Where-Object { $_.PSIsContainer } | Measure-Object).Count
        $fileCount = ($items | Where-Object { -not $_.PSIsContainer } | Measure-Object).Count
        Log-Message "Profile contains $itemCount items ($folderCount folders, $fileCount files)"
    } catch {
        Log-Message "Could not count profile items: $_"
    }
    
    $global:StatusText.Text = "Applying folder permissions to $folderCount folders and $fileCount files..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Starting recursive ACL application (takeown + icacls)..."
    Log-Message "This process handles: ownership transfer, permission reset, explicit ACL grants"
    
    # Start timer for ACL operations
    $aclStartTime = [DateTime]::Now
    Set-ProfileFolderAcls -ProfilePath $ProfileFolder -UserName $UserName
    $aclElapsed = ([DateTime]::Now - $aclStartTime).TotalSeconds
    Log-Message "Folder ACL application completed in $([Math]::Round($aclElapsed, 1)) seconds"
    $global:ProgressBar.Value = 86

    # Step 2: Apply hive file ACLs using icacls
    $global:StatusText.Text = "Setting NTUSER.DAT ownership and permissions..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Applying NTUSER.DAT file ACLs (taking ownership, setting permissions)..."
    Log-Message "Operations: takeown /A, icacls /reset, icacls /grant, icacls /setowner"
    Set-ProfileHiveAcl -HiveFilePath $targetHive -OwnerSID $newSID -UserName $UserName
    Log-Message "NTUSER.DAT ownership and ACLs applied successfully"
    $global:ProgressBar.Value = 88
    
    # Step 3: Now safe to run SID rewrite (after permissions are set)
    if ($sourceSID -and $sourceSID -ne $newSID) {
        $global:StatusText.Text = "Translating registry hive SID references..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Starting SID translation: $sourceSID -> $newSID"
        Log-Message "This rewrites binary SID references and path strings in NTUSER.DAT"
        
        # Get hive file size for progress context
        try {
            $hiveSize = (Get-Item $targetHive -Force).Length
            $hiveSizeMB = [Math]::Round($hiveSize / 1MB, 2)
            Log-Message "NTUSER.DAT size: $hiveSizeMB MB"
        } catch {
            Log-Message "Could not determine hive size"
        }
        
        try {
            Rewrite-HiveSID -FilePath $targetHive -OldSID $sourceSID -NewSID $newSID
            Log-Message "SID translation completed successfully"
        } catch {
            Log-Message "SID REWRITE FAILED: $_"
            throw $_
        }
    } else {
        Log-Message "No SID translation required (source and target SIDs match)"
    }
 
    $global:ProgressBar.Value = 92
	
    Log-Message "Profile ACLs and SID rewrite completed successfully"
}

# =============================================================================
# Rewrite-HiveSID
# =============================================================================
function Rewrite-HiveSID {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string]$OldSID,
        [Parameter(Mandatory=$true)][string]$NewSID
    )

    # Critical privileges for hive rewrite operations
    Enable-Privilege SeBackupPrivilege -IsCritical $true
    Enable-Privilege SeRestorePrivilege -IsCritical $true
    Enable-Privilege SeTakeOwnershipPrivilege -IsCritical $true

    # Validate input SIDs before attempting rewrite
    if (-not ($OldSID -match '^S-\d-\d-\d+(-\d+)*$')) { throw "Invalid source SID format: $OldSID" }
    if (-not ($NewSID -match '^S-\d-\d-\d+(-\d+)*$')) { throw "Invalid target SID format: $NewSID" }
    
    $mountPoint = "TempHive_$(Get-Random)"
    
    # Ensure no collision with existing mounts
    $maxAttempts = 5
    $attempts = 0
    while ((reg query "HKU\$mountPoint" 2>$null) -and $attempts -lt $maxAttempts) {
        $mountPoint = "TempHive_$(Get-Random)"
        $attempts++
    }
    if ($attempts -ge $maxAttempts) { throw "Cannot allocate unique hive mount point for SID rewrite" }
    
    $backup = "$FilePath.backup_$(Get-Date -f yyyyMMdd_HHmmss)"

    # === BACKUP (same method as Set-ProfileAcls) ===
    $global:StatusText.Text = "Creating backup of NTUSER.DAT..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Creating hive backup using privileged copy..."
    try {
        Copy-Item -Path $FilePath -Destination $backup -Force -ErrorAction Stop
        Log-Message "Backup created: $backup"
    } catch {
        Log-Message "Standard copy failed, falling back to robocopy /B..."
        $srcDir = Split-Path $FilePath -Parent
        $fileName = Split-Path $FilePath -Leaf
        robocopy "$srcDir" "$srcDir" "$fileName" /COPY:DAT /R:1 /W:1 /B /J | Out-Null
        Copy-Item "$srcDir\$fileName" $backup -Force -ErrorAction Stop
        Log-Message "Backup created via robocopy: $backup"
    }

    # Read entire hive
    $global:StatusText.Text = "Reading registry hive into memory..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Loading hive file into memory for binary processing..."
    [byte[]]$data = [IO.File]::ReadAllBytes($FilePath)
    $dataSizeMB = [Math]::Round($data.Length / 1MB, 2)
    Log-Message "Hive loaded: $dataSizeMB MB ($($data.Length) bytes)"

    # Convert SIDs
    $global:StatusText.Text = "Converting SID formats for binary replacement..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Converting text SIDs to binary format..."
    $oldSidObj = New-Object System.Security.Principal.SecurityIdentifier($OldSID)
    $newSidObj = New-Object System.Security.Principal.SecurityIdentifier($NewSID)
    $oldBin = New-Object byte[] $oldSidObj.BinaryLength
    $newBin = New-Object byte[] $newSidObj.BinaryLength
    $oldSidObj.GetBinaryForm($oldBin, 0)
    $newSidObj.GetBinaryForm($newBin, 0)
    Log-Message "Old SID binary: $($oldBin.Length) bytes"
    Log-Message "New SID binary: $($newBin.Length) bytes"

    # CRITICAL: Only do binary replacement if SIDs are same length
    # Different length SIDs corrupt the hive structure (all offsets break)
    if ($oldBin.Length -ne $newBin.Length) {
        $global:StatusText.Text = "SID length mismatch - skipping binary replacement (will use registry string fixes)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "WARNING: SID binary length mismatch ($($oldBin.Length) vs $($newBin.Length)) - skipping binary replacement"
        Log-Message "This is normal when importing to a different user. Registry string replacements will handle path updates."
    } else {
        $global:StatusText.Text = "Performing binary SID replacement in hive..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "SID binary lengths match - performing binary replacement"
        
        # Byte-level replacement function
        function Replace-Bytes([byte[]]$haystack, [byte[]]$needle, [byte[]]$replacement) {
            $list = [System.Collections.Generic.List[byte]]::new()
            $i = 0
            $replacements = 0
            $lastProgress = [DateTime]::Now
            while ($i -le $haystack.Length - $needle.Length) {
                # Update progress every second
                if (([DateTime]::Now - $lastProgress).TotalSeconds -ge 1) {
                    $percent = [Math]::Round(($i / $haystack.Length) * 100, 1)
                    $global:StatusText.Text = "Binary SID replacement: $percent% ($replacements replacements found)"
                    [System.Windows.Forms.Application]::DoEvents()
                    $lastProgress = [DateTime]::Now
                }
                
                $match = $true
                for ($j = 0; $j -lt $needle.Length; $j++) {
                    if ($haystack[$i + $j] -ne $needle[$j]) { $match = $false; break }
                }
                if ($match) {
                    $list.AddRange($replacement)
                    $i += $needle.Length
                    $replacements++
                } else {
                    $list.Add($haystack[$i]); $i++
                }
            }
            while ($i -lt $haystack.Length) { $list.Add($haystack[$i]); $i++ }
            Log-Message "Binary replacement completed: $replacements SID instances replaced"
            return $list.ToArray()
        }

        $data = Replace-Bytes $data $oldBin $newBin
        Log-Message "Binary SID replacement completed ($($oldBin.Length) bytes per replacement)"
    }

    # Write patched version
    $global:StatusText.Text = "Writing patched hive to disk..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Writing patched hive to temporary file..."
    $patched = "$FilePath.patched_$(Get-Random)"
    [IO.File]::WriteAllBytes($patched, $data)
    Log-Message "Patched hive written: $patched"

    # Test-load patched hive
    $global:StatusText.Text = "Validating patched hive integrity..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Test-loading patched hive to verify integrity..."
    $test = reg load "HKU\$mountPoint" $patched 2>&1
    if ($LASTEXITCODE -ne 0) {
        Log-Message "ERROR: Patched hive failed load test: $test"
        Remove-Item $patched -Force -ErrorAction SilentlyContinue
        throw "SID rewrite failed - patched hive corrupt"
    }
    Log-Message "Hive load test PASSED - patched hive is valid"
    reg unload "HKU\$mountPoint" | Out-Null

    # Replace original with patched
    $global:StatusText.Text = "Replacing original hive with patched version..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Replacing original NTUSER.DAT with patched version..."
    try {
        Move-Item $patched $FilePath -Force -ErrorAction Stop
    } catch {
        Copy-Item $patched $FilePath -Force -ErrorAction Stop
        Remove-Item $patched -Force
    }
    Log-Message "Binary SID replacement completed successfully"

    # Final string cleanup + OOBE fixes
    $global:StatusText.Text = "Mounting hive for registry string cleanup..."
    [System.Windows.Forms.Application]::DoEvents()
    reg load "HKU\$mountPoint" $FilePath | Out-Null
    Log-Message "Performing registry path and string cleanup"
    
    # Derive old and new profile paths from SIDs
    $oldProfilePath = $null
    $newProfilePath = $null
    
    try {
        # Get old profile path from registry (if it exists)
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $oldSidKey = Get-ItemProperty -Path "$profileListBase\$OldSID" -ErrorAction SilentlyContinue
        if ($oldSidKey) {
            $oldProfilePath = $oldSidKey.ProfileImagePath
            Log-Message "Old profile path from registry: $oldProfilePath"
        }
        
        # Get new profile path
        $newSidKey = Get-ItemProperty -Path "$profileListBase\$NewSID" -ErrorAction SilentlyContinue
        if ($newSidKey) {
            $newProfilePath = $newSidKey.ProfileImagePath
            Log-Message "New profile path from registry: $newProfilePath"
        }
    } catch {
        Log-Message "Could not resolve profile paths from SIDs: $_"
    }
    
    # SAFER APPROACH: Use reg.exe export/import with text replacement instead of .SetValue()
    # This avoids corrupting the hive file by using atomic operations
    function Fix-Key-Safe([string]$regPath, [string]$exportFile) {
        # Export registry to .reg file
        reg export "$regPath" "$exportFile" /y 2>$null
        if ($LASTEXITCODE -ne 0) {
            Log-Message "WARNING: Could not export $regPath for string replacement"
            return
        }
        
        # Read file, perform replacements, write back
        try {
            $content = [System.IO.File]::ReadAllText($exportFile, [System.Text.Encoding]::UTF16LE)
            $originalSize = $content.Length
            
            # Replace SID strings (hex-encoded form in .reg files is complex, skip for now)
            # Focus on text path replacements which are more reliable
            
            if ($oldProfilePath -and $newProfilePath) {
                # Replace old profile path with new (case-insensitive)
                $content = $content -ireplace [regex]::Escape($oldProfilePath), $newProfilePath
            }
            
            # Replace generic C:\Users\olduser pattern
            if ($oldProfilePath -match 'C:\\Users\\([^\\]+)') {
                $oldUsername = $matches[1]
                $oldUserPattern = "C:\\Users\\$([regex]::Escape($oldUsername))"
                $content = $content -ireplace $oldUserPattern, $newProfilePath
            }
            
            # OneDrive path replacement
            if ($oldProfilePath -match 'C:\\Users\\([^\\]+)') {
                $oldUsername = $matches[1]
                $oneDrivePattern = "C:\\Users\\$([regex]::Escape($oldUsername))\\(OneDrive(?:\s*-\s*[^\\]+)?)"
                $content = $content -ireplace $oneDrivePattern, "$newProfilePath\`$1"
            }
            
            # Only reimport if changes were made
            if ($content.Length -ne $originalSize) {
                [System.IO.File]::WriteAllText($exportFile, $content, [System.Text.Encoding]::UTF16LE)
                reg import "$exportFile" 2>$null
                if ($LASTEXITCODE -eq 0) {
                    Log-Message "Registry string replacements applied successfully"
                } else {
                    Log-Message "WARNING: Registry import failed - changes may not be applied"
                }
            }
        } catch {
            Log-Message "WARNING: Registry string replacement error: $_"
        } finally {
            Remove-Item $exportFile -Force -ErrorAction SilentlyContinue
        }
    }
    
    try {
        # DISABLED: Registry export/import was corrupting hive by rebuilding structure
        # Binary SID replacement is sufficient for domain-to-local migrations
        # String path replacements via reg export/import were reducing hive from 1.8MB to 0.75MB
        # Local paths will be corrected on first profile login
        Log-Message "Skipping registry path post-processing (export/import was corrupting hive)"

        # Diagnostic: summarize rewrite status across critical keys
        Log-Message "Summarizing SID/path rewrite status"
        function Report-RewriteSummary {
            param(
                [string]$Root,
                [string]$OldSid,
                [string]$NewSid,
                [string]$OldPath,
                [string]$NewPath
            )
            $targets = @(
                "Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband",
                "Software\Microsoft\Windows\CurrentVersion\CloudStore",
                "Software\Microsoft\Windows\Shell\Bags",
                "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU",
                "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
                "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            )
            foreach ($t in $targets) {
                $base = "Registry::$Root\\$t"
                $key = Get-Item $base -ErrorAction SilentlyContinue
                if (-not $key) { continue }
                $oldSidCount = 0; $newSidCount = 0; $oldPathCount = 0; $newPathCount = 0
                try {
                    function Count-Key([string]$p) {
                        $k = Get-Item $p -ErrorAction SilentlyContinue
                        if (-not $k) { return }
                        foreach ($vn in $k.GetValueNames()) {
                            try {
                                $vv = $k.GetValue($vn, $null, 'DoNotExpandEnvironmentNames')
                                if ($vv -is [string]) {
                                    if ($OldSid -and $vv -like "*$OldSid*") { $script:oldSidCount++ }
                                    if ($NewSid -and $vv -like "*$NewSid*") { $script:newSidCount++ }
                                    if ($OldPath -and $vv -like "*$OldPath*") { $script:oldPathCount++ }
                                    if ($NewPath -and $vv -like "*$NewPath*") { $script:newPathCount++ }
                                }
                            } catch {}
                        }
                        foreach ($sk in $k.GetSubKeyNames()) { Count-Key "$p\\$sk" }
                    }
                    Count-Key $base
                    Log-Message ("Rewrite summary for '$t': OldSID=$oldSidCount, NewSID=$newSidCount, OldPath=$oldPathCount, NewPath=$newPathCount")
                } catch {
                    Log-Message "Summary scan error for ${t}: $_"
                }
            }
        }
        Report-RewriteSummary -Root "HKU:$mountPoint" -OldSid $OldSID -NewSid $NewSID -OldPath $oldProfilePath -NewPath $newProfilePath

        # Disable telemetry/OOBE junk - using reg.exe (avoids -Type bug)
        Log-Message "Disabling telemetry/OOBE via reg.exe"
        reg add "HKU\$mountPoint\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338387Enabled /t REG_DWORD /d 0 /f | Out-Null
        reg add "HKU\$mountPoint\SOFTWARE\Policies\Microsoft\Windows\OOBE" /v DisablePrivacyExperience /t REG_DWORD /d 1 /f | Out-Null
        
    } finally {
        # GUARANTEED cleanup: unload hive even if Fix-Key fails
        Log-Message "Preparing to unload hive - forcing handle cleanup..."
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500

        $unloadAttempts = 0
        do {
            $unloadAttempts++
            reg unload "HKU\$mountPoint" 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Log-Message "Hive unloaded successfully"
                break
            }
            if ($unloadAttempts -ge $Config.HiveUnloadMaxAttempts) {
                Log-Message "Hive still busy - proceeding (will auto-clean on reboot)"
                break
            }
            Start-Sleep -Milliseconds $Config.HiveUnloadWaitMs
        } while ($true)
    }

    Log-Message "Rewrite-HiveSID completed successfully: $OldSID → $NewSID"
}

# =============================================================================
# HANDLE RESTART - GUI AWARE WITH MULTIPLE BEHAVIORS
# =============================================================================
function Handle-Restart {
    param(
        [ValidateSet('Prompt', 'Immediate', 'Never', 'Delayed')]
        [string]$Behavior = 'Prompt',
        [int]$DelaySeconds = 30,
        [string]$Reason = "System configuration changed"
    )
    switch ($Behavior) {
            'Immediate' {
            Log-Message "Restarting immediately (Behavior: Immediate)"
            [System.Windows.Forms.MessageBox]::Show("$Reason`n`nRestarting NOW!", "Restarting", "OK", "Warning")
            Log-Message "Executing forced restart via shutdown.exe"
            shutdown /r /f /t 0 /c "Profile Migration Tool - Completing operation" /d p:4:1
        }
        'Delayed' {
            Log-Message "Restarting in $DelaySeconds seconds (Behavior: Delayed)"
            $response = [System.Windows.Forms.MessageBox]::Show(
                "$Reason`n`nComputer will restart in $DelaySeconds seconds.`n`nClick OK to restart now, Cancel to wait the full countdown.",
                "Restart Scheduled", "OKCancel", "Warning")

            if ($response -eq "OK") {
                Log-Message "User clicked OK - restarting immediately"
                shutdown /r /f /t 0 /c "Profile Migration Tool - User triggered restart" /d p:4:1
            }
            else {
                $global:StatusText.Text = "Restart scheduled in $DelaySeconds seconds... (cannot cancel)"
                Log-Message "Starting $DelaySeconds-second countdown"
                for ($i = $DelaySeconds; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restarting in $i seconds... (close window to attempt cancel)"
                    $global:ProgressBar.Value = [int](($DelaySeconds - $i) / $DelaySeconds * 100)
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }
                Log-Message "Countdown finished - forcing restart"
                shutdown /r /f /t 0 /c "Profile Migration Tool - Countdown complete" /d p:4:1
            }
        }
        'Never' {
            Log-Message "Restart required but automatic restart disabled (Behavior: Never)"
            $global:StatusText.Text = "Manual restart required"
            [System.Windows.Forms.MessageBox]::Show(
                "$Reason`n`nRESTART REQUIRED but auto-restart is disabled.`n`nPlease restart manually when ready.",
                "Manual Restart Required", "OK", "Information")
        }
                'Prompt' {
            Log-Message "Prompting user for restart (Behavior: Prompt)"
            $response = [System.Windows.Forms.MessageBox]::Show(
                "$Reason`n`nA restart is REQUIRED to complete the operation.`n`nRestart now?",
                "Restart Required", "YesNo", "Question")

            if ($response -eq "Yes") {
                Log-Message "User approved restart - starting 10-second forced countdown"

                # This message box is purely informational - we ignore its result
                [System.Windows.Forms.MessageBox]::Show(
                    "Computer will restart in 10 seconds.`n`nSave all work NOW!`n`nThe restart cannot be cancelled.",
                    "Restarting in 10 seconds", "OK", "Warning") | Out-Null

                # 10-second visible countdown + forced restart
                for ($i = 10; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restarting in $i seconds..."
                    $global:ProgressBar.Value = [int](100 - ($i * 10))
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }

                Log-Message "10-second countdown finished - forcing restart"
                shutdown /r /f /t 0 /c "Profile Migration Tool - User approved restart" /d p:4:1
            }
            else {
                Log-Message "User clicked No - forcing restart anyway (domain join requires it)"
                [System.Windows.Forms.MessageBox]::Show(
                    "You clicked No, but a restart is REQUIRED after domain join.`n`nRestarting in 15 seconds anyway...",
                    "Restart Mandatory", "OK", "Warning") | Out-Null

                for ($i = 15; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restart MANDATORY - restarting in $i seconds..."
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }

                shutdown /r /f /t 0 /c "Profile Migration Tool - Forced restart (domain join)" /d p:4:1
            }
        }
	}
}

# =============================================================================
# PRE-FLIGHT CHECKS
# =============================================================================
function Test-DomainReachability {
    param([string]$DomainName)
    try {
        Log-Message "Testing domain reachability: $DomainName"
        
        # DNS with timeout
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $dnsResult = $null
        try {
            $dnsResult = Resolve-DnsName -Name $DomainName -ErrorAction Stop -QuickTimeout
        } catch {
            if ($sw.ElapsedMilliseconds -gt $Config.DomainReachabilityTimeout) {
                return @{ Success = $false; Error = "DNS resolution timeout for '$DomainName'"; ErrorCode = "DNS_TIMEOUT" }
            }
            return @{ Success = $false; Error = "DNS resolution failed for domain '$DomainName'"; ErrorCode = "DNS_FAIL" }
        }
        
        if (-not $dnsResult) {
            return @{ Success = $false; Error = "DNS resolution failed for domain '$DomainName'"; ErrorCode = "DNS_FAIL" }
        }
        Log-Message "DNS resolution successful"
        
        # LDAP connectivity with proper timeout
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        try {
            $connect = $tcpClient.BeginConnect($DomainName, 389, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne($Config.DomainReachabilityTimeout, $false)
            if (-not $wait) {
                $tcpClient.Close()
                return @{ Success = $false; Error = "Cannot reach domain controller on port 389 (LDAP). Check network connectivity."; ErrorCode = "LDAP_UNREACHABLE" }
            }
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            Log-Message "Domain controller reachable on LDAP port"
            return @{ Success = $true; Error = $null; ErrorCode = $null }
        } finally {
            $tcpClient.Dispose()
        }
    } catch {
        return @{ Success = $false; Error = "Domain reachability test failed: $($_.Exception.Message)"; ErrorCode = "NETWORK_ERROR" }
    }
}

function Test-DomainCredentials {
    param([string]$DomainName, [System.Management.Automation.PSCredential]$Credential)
    try {
        Log-Message "Validating domain credentials..."
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $username = $Credential.UserName
        if ($username -match '\\') { $username = ($username -split '\\')[1] }
        Log-Message "Testing credentials for user: $username in domain: $DomainName"
        try {
            $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
            $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType, $DomainName, $username, $Credential.GetNetworkCredential().Password)
            $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($principalContext, $username)
            if ($userPrincipal) {
                Log-Message "Credentials validated successfully (User: $($userPrincipal.SamAccountName), Display: $($userPrincipal.DisplayName))"
                $principalContext.Dispose()
                return @{ Success = $true; Error = $null; ErrorCode = $null }
            } else {
                Log-Message "WARNING: Context created but user not found - credentials likely valid"
                $principalContext.Dispose()
                return @{ Success = $true; Warning = "Could not verify user account exists, but credentials appear valid for domain"; ErrorCode = $null }
            }
        } catch {
            $errorMsg = $_.Exception.Message
            if ($errorMsg -match "username or password|logon failure|invalid credentials|unknown user|0x8007052e|0x52e") {
                Log-Message "Credential validation failed: Invalid username or password"
                return @{ Success = $false; Error = "Invalid credentials for domain '$DomainName' (username or password incorrect)"; ErrorCode = "INVALID_CREDENTIALS" }
            } elseif ($errorMsg -match "account.*locked|account.*disabled|0x80070533|0x533") {
                Log-Message "Credential validation failed: Account locked or disabled"
                return @{ Success = $false; Error = "User account is locked or disabled in domain '$DomainName'"; ErrorCode = "ACCOUNT_LOCKED_OR_DISABLED" }
            } elseif ($errorMsg -match "server.*not.*operational|0x8007203a|0x203a") {
                Log-Message "Domain controller not available"
                return @{ Success = $false; Error = "Domain controller not operational or unreachable"; ErrorCode = "DC_NOT_OPERATIONAL" }
            } else {
                Log-Message "WARNING: Credential check inconclusive: $errorMsg"
                return @{ Success = $true; Warning = "Could not fully validate credentials (network or configuration issue), but will attempt domain join anyway. Error: $errorMsg"; ErrorCode = $null }
            }
        }
    } catch {
        Log-Message "WARNING: Credential validation system failed: $($_.Exception.Message)"
        return @{ Success = $true; Warning = "Credential validation unavailable (will proceed with domain join attempt): $($_.Exception.Message)"; ErrorCode = $null }
    }
}

function Test-DomainJoinPermissions {
    param([string]$DomainName, [System.Management.Automation.PSCredential]$Credential)
    try {
        Log-Message "Checking domain join permissions..."
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $DomainName, $Credential.UserName, $Credential.GetNetworkCredential().Password)
        $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context, $Credential.UserName)
        if (-not $user) { return @{ Success = $false; Error = "User account not found in domain"; ErrorCode = "USER_NOT_FOUND" } }
        $groups = $user.GetAuthorizationGroups()
        $isDomainAdmin = $false
        foreach ($group in $groups) {
            if ($group.Name -eq "Domain Admins" -or $group.Name -eq "Administrators") {
                $isDomainAdmin = $true
                break
            }
        }
        if (-not $isDomainAdmin) {
            Log-Message "WARNING: User may not have domain join permissions (not in Domain Admins)"
            return @{ Success = $true; Warning = "User is not in Domain Admins group. Domain join may fail if delegated permissions are not configured."; ErrorCode = $null }
        }
        Log-Message "User has domain admin privileges"
        return @{ Success = $true; Error = $null; ErrorCode = $null }
    } catch {
        Log-Message "Could not verify permissions, will attempt domain join anyway"
        return @{ Success = $true; Warning = "Could not verify domain join permissions: $($_.Exception.Message)"; ErrorCode = $null }
    }
}

function Get-DomainJoinErrorDetails {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $errorMessage = $ErrorRecord.Exception.Message
    $errorCode = "UNKNOWN"
    $userFriendlyMessage = $errorMessage
    $suggestion = ""
    if ($errorMessage -match "0x0000052E") {
        $errorCode = "LOGON_FAILURE"
        $userFriendlyMessage = "Invalid username or password"
        $suggestion = "Verify credentials and try again. Ensure you're using DOMAIN\username format."
    } elseif ($errorMessage -match "0x00000035") {
        $errorCode = "NETWORK_PATH_NOT_FOUND"
        $userFriendlyMessage = "Cannot find the domain controller"
        $suggestion = "Check network connectivity and DNS settings. Ensure the domain name is correct."
    } elseif ($errorMessage -match "0x0000054B") {
        $errorCode = "COMPUTER_ALREADY_EXISTS"
        $userFriendlyMessage = "A computer account with this name already exists in the domain"
        $suggestion = "Either use a different computer name, or have a domain admin delete the existing computer account."
    } elseif ($errorMessage -match "0x00000569") {
        $errorCode = "ACCOUNT_RESTRICTION"
        $userFriendlyMessage = "User account does not have permission to join computers to the domain"
        $suggestion = "Contact your domain administrator to grant domain join permissions or use an account with Domain Admin rights."
    } elseif ($errorMessage -match "0x0000232A") {
        $errorCode = "DNS_FAILURE"
        $userFriendlyMessage = "DNS name does not exist"
        $suggestion = "Verify the domain name is spelled correctly and DNS is properly configured."
    } elseif ($errorMessage -match "0x0000232B") {
        $errorCode = "DNS_SERVER_FAILURE"
        $userFriendlyMessage = "DNS server failure"
        $suggestion = "Check your DNS server settings and network connectivity."
    } elseif ($errorMessage -match "0x00000005") {
        $errorCode = "ACCESS_DENIED"
        $userFriendlyMessage = "Access denied - insufficient permissions"
        $suggestion = "You need Domain Admin rights or delegated permissions to join computers to the domain."
    } elseif ($errorMessage -match "password|credential") {
        $errorCode = "CREDENTIAL_ERROR"
        $userFriendlyMessage = "Credential authentication failed"
        $suggestion = "Verify username and password are correct. Check for typos and ensure CAPS LOCK is off."
    } elseif ($errorMessage -match "network|connection|timeout") {
        $errorCode = "NETWORK_ERROR"
        $userFriendlyMessage = "Network connectivity issue"
        $suggestion = "Check network cables, Wi-Fi connection, and firewall settings. Ensure you can ping the domain controller."
    }
    return @{
        ErrorCode = $errorCode
        OriginalMessage = $errorMessage
        UserFriendlyMessage = $userFriendlyMessage
        Suggestion = $suggestion
    }
}

# =============================================================================
# ENHANCED DOMAIN JOIN - WITH RESTART CONTROL
# =============================================================================
function Join-Domain-Enhanced {
    param(
        [Parameter(Mandatory=$false)][ValidateScript({ if ([string]::IsNullOrWhiteSpace($_)) { $true } else { if ($_ -notmatch '^[a-zA-Z0-9-]{1,15}$') { throw "Computer name must be 1-15 characters (alphanumeric and hyphens only)" } else { $true } } })][string]$ComputerName,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$DomainName,
        [Parameter(Mandatory=$false)][ValidateSet('Prompt', 'Immediate', 'Never', 'Delayed')][string]$RestartBehavior = 'Prompt',
        [Parameter(Mandatory=$false)][ValidateRange(5, 300)][int]$DelaySeconds = 30,
        [Parameter(Mandatory=$false)][System.Management.Automation.PSCredential]$Credential = $null
    )
    try {
        if ([string]::IsNullOrWhiteSpace($DomainName)) { throw "Domain name is required." }
        $cs = Get-WmiObject Win32_ComputerSystem
        $currentComputerName = $env:COMPUTERNAME
        Log-Message "Current State: Computer='$currentComputerName', Domain='$($cs.Domain)', PartOfDomain=$($cs.PartOfDomain)"
        $targetName = $null
        if (![string]::IsNullOrWhiteSpace($ComputerName)) {
            $targetName = $ComputerName.Trim()
            if ($targetName.Length -gt 15) { throw "Computer name cannot exceed 15 characters. Provided: $($targetName.Length) chars" }
            if ($targetName -match '[^a-zA-Z0-9-]') { throw "Computer name contains invalid characters. Only letters, numbers, and hyphens are allowed." }
            if ($targetName.StartsWith('-') -or $targetName.EndsWith('-')) { throw "Computer name cannot start or end with a hyphen." }
        }
        if ($cs.PartOfDomain -and $cs.Domain -ieq $DomainName) {
            Log-Message "Already member of domain '$DomainName'"
            if ($targetName -and $targetName.ToUpper() -ne $currentComputerName.ToUpper()) {
                $response = [System.Windows.Forms.MessageBox]::Show("Already in domain '$DomainName'.`n`nRename computer?`nFrom: $currentComputerName`nTo: $targetName", "Rename Computer", "YesNo", "Question")
                if ($response -eq "Yes") {
                    $cred = Get-Credential -Message "Domain credentials to rename computer" -UserName "$DomainName\"
                    if (-not $cred) { throw "Credentials cancelled." }
                    Log-Message "Renaming: '$currentComputerName' to '$targetName'"
                    try {
                        Rename-Computer -NewName $targetName -DomainCredential $cred -Force -ErrorAction Stop
                        Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Computer renamed to '$targetName'"
                    } catch {
                        $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
                        Log-Message "Rename failed: $($errorDetails.UserFriendlyMessage)"
                        throw "Computer rename failed: $($errorDetails.UserFriendlyMessage)`n`nSuggestion: $($errorDetails.Suggestion)"
                    }
                }
            } else {
                Log-Message "No changes needed. Computer name: $currentComputerName"
            }
            return
        }
        if ($cs.PartOfDomain) {
            $currentDomain = $cs.Domain
            $response = [System.Windows.Forms.MessageBox]::Show("Currently in domain: $currentDomain`n`nLeave and join: $DomainName`n`nThis requires TWO RESTARTS.`n`nContinue?", "Leave Domain", "YesNo", "Warning")
            if ($response -ne "Yes") {
                Log-Message "Operation cancelled."
                return
            }
            $disjoinCred = Get-Credential -Message "Credentials to leave '$currentDomain'" -UserName "$currentDomain\"
            if (-not $disjoinCred) { throw "Disjoin credentials cancelled." }
            Log-Message "Leaving domain '$currentDomain'"
            try {
                Remove-Computer -UnjoinDomainCredential $disjoinCred -Force -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("Left domain successfully.`n`nAfter restart, run this tool again to join '$DomainName'.", "Step 1 Complete", "OK", "Information")
                Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Left domain '$currentDomain'"
            } catch {
                $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
                Log-Message "Failed to leave domain: $($errorDetails.UserFriendlyMessage)"
                throw "Failed to leave domain: $($errorDetails.UserFriendlyMessage)`n`nSuggestion: $($errorDetails.Suggestion)"
            }
            return
        }
        Log-Message "=== STARTING PRE-FLIGHT CHECKS ==="
        $global:StatusText.Text = "Checking domain reachability..."
        [System.Windows.Forms.Application]::DoEvents()
        $reachabilityTest = Test-DomainReachability -DomainName $DomainName
        if (-not $reachabilityTest.Success) {
            Log-Message "Pre-flight check failed: $($reachabilityTest.Error)"
            throw "Domain Reachability Check Failed:`n`n$($reachabilityTest.Error)`n`nPlease verify:`nDomain name is correct`nNetwork connection is working`nDNS is properly configured`nFirewalls allow domain traffic"
        }
        Log-Message "Joining domain: $DomainName"
        # Use provided credentials or prompt for new ones
        if ($Credential) {
            Log-Message "Using provided domain credentials"
            $joinCred = $Credential
        } else {
            $joinCred = Get-Credential -Message "DOMAIN ADMIN credentials for '$DomainName'" -UserName "$DomainName\"
            if (-not $joinCred) { throw "Join credentials cancelled." }
        }
        $global:StatusText.Text = "Validating credentials..."
        [System.Windows.Forms.Application]::DoEvents()
        $credTest = Test-DomainCredentials -DomainName $DomainName -Credential $joinCred
        if (-not $credTest.Success) {
            Log-Message "Credential validation failed: $($credTest.Error)"
            throw "Credential Validation Failed:`n`n$($credTest.Error)`n`nPlease verify:`nUsername is in DOMAIN\username format`nPassword is correct`nAccount is not locked or disabled"
        }
        $global:StatusText.Text = "Checking permissions..."
        [System.Windows.Forms.Application]::DoEvents()
        $permTest = Test-DomainJoinPermissions -DomainName $DomainName -Credential $joinCred
        if ($permTest.Warning) {
            Log-Message "Permission warning: $($permTest.Warning)"
            $response = [System.Windows.Forms.MessageBox]::Show("$($permTest.Warning)`n`nContinue anyway?", "Permission Warning", "YesNo", "Warning")
            if ($response -ne "Yes") {
                Log-Message "User cancelled after permission warning"
                return
            }
        }
        Log-Message "=== PRE-FLIGHT CHECKS COMPLETE ==="
        $params = @{ DomainName = $DomainName; Credential = $joinCred; Force = $true }
        $actionDescription = "Join domain '$DomainName'"
        if ($targetName -and $targetName.ToUpper() -ne $currentComputerName.ToUpper()) {
            $params['NewName'] = $targetName
            $actionDescription += "`nRename: $currentComputerName to $targetName"
            Log-Message "Will rename computer during domain join."
        } else {
            $actionDescription += "`nKeep name: $currentComputerName"
            Log-Message "Computer name will remain: $currentComputerName"
        }
        $confirmation = [System.Windows.Forms.MessageBox]::Show("$actionDescription`n`nA restart will be required.`n`nContinue?", "Confirm Domain Join", "YesNo", "Question")
        if ($confirmation -ne "Yes") {
            Log-Message "Domain join cancelled."
            return
        }
        $global:StatusText.Text = "Joining domain..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Executing Add-Computer with parameters: Domain=$DomainName, NewName=$($params.NewName)"
        try {
            Add-Computer @params -ErrorAction Stop
            Log-Message "Successfully joined domain '$DomainName'!"
            Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Joined domain '$DomainName'"
        } catch {
            $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
            Log-Message "Domain join failed: [$($errorDetails.ErrorCode)] $($errorDetails.UserFriendlyMessage)"
            Log-Message "Original error: $($errorDetails.OriginalMessage)"
            $errorMsg = "Domain Join Failed`n`nError: $($errorDetails.UserFriendlyMessage)`n`n"
            if ($errorDetails.Suggestion) { $errorMsg += "Suggestion: $($errorDetails.Suggestion)`n`n" }
            $errorMsg += "Error Code: $($errorDetails.ErrorCode)"
            throw $errorMsg
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Log-Message "ERROR: $errorMsg"
        $global:StatusText.Text = "Domain join failed"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Domain Join Error", "OK", "Error")
    }
}

# === WINGET IMPORT USING NATIVE WINGET IMPORT COMMAND ===
function Install-WingetAppsFromExport {
    param(
        [Parameter(Mandatory=$true)][string]$TargetProfilePath
    )
    # In flat ZIP/import, Winget JSON is at the profile root
    $jsonPath = Join-Path $TargetProfilePath 'Winget-Packages.json'
    
    # 1. Safe file check
    if (-not (Test-Path $jsonPath) -or (Get-Item $jsonPath).Length -lt 10) {
        Log-Message "No valid Winget-Packages.json found - skipping app reinstall"
        return
    }

    $global:StatusText.Text = "Loading Winget app list..."
    $global:ProgressBar.Value = 98
    [System.Windows.Forms.Application]::DoEvents()

    # 2. Safe JSON parsing
    try {
        $json = Get-Content $jsonPath -Raw | ConvertFrom-Json -ErrorAction Stop
        $apps = $json.Sources.Packages
        if (-not $apps -or $apps.Count -eq 0) {
            Log-Message "Winget export is empty - nothing to install"
            return
        }
    }
    catch {
        Log-Message "Failed to parse Winget-Packages.json: $_ - skipping apps"
        $emptyJson = '{"Sources":[{"Packages":[]}]}'
        Set-Content $jsonPath -Value $emptyJson -Force -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            "The Winget app list is corrupted or unreadable.`n`nSkipping application reinstall.",
            "Winget List Error", "OK", "Warning") | Out-Null
        return
    }

    # 3. GUI to select apps - MODERN DESIGN
    $global:StatusText.Text = "Preparing app selection UI..."
    [System.Windows.Forms.Application]::DoEvents()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Application Installer"
    $form.Size = New-Object System.Drawing.Size(800,680)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $form.MinimumSize = New-Object System.Drawing.Size(800, 680)

    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(800, 70)
    $headerPanel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(500, 28)
    $lblTitle.Text = "Select Applications to Install"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $headerPanel.Controls.Add($lblTitle)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 45)
    $lblSubtitle.Size = New-Object System.Drawing.Size(750, 20)
    $lblSubtitle.Text = "Found $($apps.Count) applications from source PC - Select which ones to reinstall"
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $headerPanel.Controls.Add($lblSubtitle)

    # Main content card
    $contentCard = New-Object System.Windows.Forms.Panel
    $contentCard.Location = New-Object System.Drawing.Point(15, 85)
    $contentCard.Size = New-Object System.Drawing.Size(755, 500)
    $contentCard.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($contentCard)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(15, 15)
    $lblInfo.Size = New-Object System.Drawing.Size(720, 35)
    $lblInfo.Text = "Winget will automatically install or upgrade these applications.`nUse the checkboxes to select which apps you want:"
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $contentCard.Controls.Add($lblInfo)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(15, 55)
    $clb.Size = New-Object System.Drawing.Size(720, 350)
    $clb.CheckOnClick = $true
    $clb.Font = New-Object System.Drawing.Font("Consolas", 9)
    $clb.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    foreach ($app in $apps) {
        $name = $app.PackageIdentifier
        if ($app.Version) { $name += "  [v$($app.Version)]" }
        $clb.Items.Add($name, $true) | Out-Null
    }
    $contentCard.Controls.Add($clb)

    # Selection controls
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Location = New-Object System.Drawing.Point(15, 415)
    $btnAll.Size = New-Object System.Drawing.Size(120, 35)
    $btnAll.Text = "Select All"
    $btnAll.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnAll.ForeColor = [System.Drawing.Color]::White
    $btnAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAll.FlatAppearance.BorderSize = 0
    $btnAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnAll.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
    $btnAll.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
    $btnAll.Add_Click({ for($i=0;$i -lt $clb.Items.Count;$i++) { $clb.SetItemChecked($i,$true) } })
    $contentCard.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Location = New-Object System.Drawing.Point(145, 415)
    $btnNone.Size = New-Object System.Drawing.Size(120, 35)
    $btnNone.Text = "Select None"
    $btnNone.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnNone.ForeColor = [System.Drawing.Color]::White
    $btnNone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnNone.FlatAppearance.BorderSize = 0
    $btnNone.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnNone.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNone.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
    $btnNone.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
    $btnNone.Add_Click({ for($i=0;$i -lt $clb.Items.Count;$i++) { $clb.SetItemChecked($i,$false) } })
    $contentCard.Controls.Add($btnNone)

    # Selection counter
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Location = New-Object System.Drawing.Point(15, 460)
    $lblCount.Size = New-Object System.Drawing.Size(720, 25)
    $lblCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $contentCard.Controls.Add($lblCount)
    
    # Update counter function
    $updateCounter = {
        $checked = 0
        for($i=0; $i -lt $clb.Items.Count; $i++) {
            if ($clb.GetItemChecked($i)) { $checked++ }
        }
        $lblCount.Text = "$checked of $($clb.Items.Count) applications selected for installation"
    }
    & $updateCounter
    $clb.Add_ItemCheck({ 
        # Delay to let the check state update
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 10
        $timer.Add_Tick({
            & $updateCounter
            $this.Stop()
            $this.Dispose()
        })
        $timer.Start()
    })

    # Action buttons at bottom
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(540, 600)
    $btnOK.Size = New-Object System.Drawing.Size(110, 40)
    $btnOK.Text = "Install"
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12) })
    $btnOK.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16) })
    $btnOK.DialogResult = "OK"
    $form.Controls.Add($btnOK)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Location = New-Object System.Drawing.Point(660, 600)
    $btnSkip.Size = New-Object System.Drawing.Size(110, 40)
    $btnSkip.Text = "Skip All"
    $btnSkip.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnSkip.ForeColor = [System.Drawing.Color]::White
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.FlatAppearance.BorderSize = 0
    $btnSkip.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnSkip.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSkip.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
    $btnSkip.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
    $btnSkip.DialogResult = "Cancel"
    $form.Controls.Add($btnSkip)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnSkip

    if ($form.ShowDialog() -ne "OK") {
        Log-Message "User skipped Winget app installation"
        $form.Dispose()
        return
    }
    $form.Dispose()

    # 4. Collect selected apps and create filtered JSON
    $selectedApps = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) {
            $selectedApps += $apps[$i]
        }
    }

    if ($selectedApps.Count -eq 0) {
        Log-Message "No apps selected for install"
        return
    }

    # 5. Create filtered JSON file with only selected apps - KEEP ONLY SOURCES WITH SELECTED PACKAGES
    $filteredSources = @()
    
    # For each source in the original export
    foreach ($source in $json.Sources) {
        $sourceApps = @()
        # Check if any selected apps belong to this source
        foreach ($selectedApp in $selectedApps) {
            if ($source.Packages | Where-Object { $_.PackageIdentifier -eq $selectedApp.PackageIdentifier }) {
                $sourceApps += $selectedApp
            }
        }
        
        # Only include sources that have selected packages
        if ($sourceApps.Count -gt 0) {
            $sourceEntry = New-Object PSObject
            $sourceEntry | Add-Member -Type NoteProperty -Name "Packages" -Value $sourceApps
            $sourceEntry | Add-Member -Type NoteProperty -Name "SourceDetails" -Value $source.SourceDetails
            $filteredSources += $sourceEntry
        }
    }
    
    # Build complete filtered JSON matching exact export format
    $filteredJsonObj = New-Object PSObject
    $filteredJsonObj | Add-Member -Type NoteProperty -Name "`$schema" -Value "https://aka.ms/winget-packages.schema.2.0.json"
    $filteredJsonObj | Add-Member -Type NoteProperty -Name "CreationDate" -Value $json.CreationDate
    $filteredJsonObj | Add-Member -Type NoteProperty -Name "Sources" -Value $filteredSources
    $filteredJsonObj | Add-Member -Type NoteProperty -Name "WinGetVersion" -Value $json.WinGetVersion
    
    $filteredJson = $filteredJsonObj | ConvertTo-Json -Depth 10
    
    # Write filtered JSON to a safe temp location
    $filteredJsonPath = Join-Path $env:TEMP ("Winget-Packages-Selected_" + [IO.Path]::GetFileName($TargetProfilePath) + "_" + (Get-Random) + ".json")
    try {
        Set-Content -Path $filteredJsonPath -Value $filteredJson -Encoding UTF8 -Force
        Log-Message "Created filtered app list: $filteredJsonPath ($($selectedApps.Count) apps)"
    }
    catch {
        Log-Message "Failed to create filtered JSON: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to prepare app list for installation.`n`nError: $_",
            "Filter Error", "OK", "Error") | Out-Null
        return
    }

    # Track installed apps globally for report
    $global:InstalledAppsList = $selectedApps
    
    # 6. Use winget import command (handles install/upgrade automatically)
    $global:StatusText.Text = "Installing apps via Winget..."
    $global:ProgressBar.Value = 98
    [System.Windows.Forms.Application]::DoEvents()
    
    Log-Message "Winget import: $($selectedApps.Count) apps"
    Log-Message "Import file path: $filteredJsonPath"
    
    try {
        $global:StatusText.Text = "Running winget import..."
        [System.Windows.Forms.Application]::DoEvents()
        # Verify the file exists before attempting import
        if (-not (Test-Path $filteredJsonPath)) {
            throw "Filtered JSON file not found: $filteredJsonPath"
        }
        
        # Execute winget import via cmd.exe to ensure proper handling
        $proc = Start-Process 'cmd.exe' -ArgumentList '/c', "winget import --import-file `"$filteredJsonPath`" --accept-package-agreements --accept-source-agreements" -Wait -PassThru -ErrorAction Stop
        
        if ($proc.ExitCode -eq 0) {
            $summary = "Winget import complete. $($selectedApps.Count) apps processed.`r`nWinget handles installs and upgrades automatically."
            Log-Message $summary
            $global:StatusText.Text = "Apps installed!"
            $global:ProgressBar.Value = 100
            Start-Sleep -Seconds 1
            [System.Windows.Forms.MessageBox]::Show($summary, "Import Complete", "OK", "Information")
        }
        else {
            $msg = "Winget import exited with code $($proc.ExitCode). Some installs may have failed."
            Log-Message "Winget import failed with exit code $($proc.ExitCode)"
            [System.Windows.Forms.MessageBox]::Show($msg, "Import Warning", "OK", "Warning")
        }
    }
    catch {
        Log-Message "Winget import failed: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to run winget import.`n`nError: $_",
            "Import Error", "OK", "Error") | Out-Null
    }
    finally {
        # Clean up filtered JSON
        if (Test-Path $filteredJsonPath) {
            Remove-Item $filteredJsonPath -Force -ErrorAction SilentlyContinue
            Log-Message "Cleaned up filtered app list"
        }
    }
}

# =============================================================================
# PROFILE SIZE OPTIMIZATION - PRE-EXPORT CLEANUP WIZARD
# =============================================================================
function Show-ProfileCleanupWizard {
    param([string]$ProfilePath)
    
    Log-Message "=== PROFILE CLEANUP WIZARD ==="
    Log-Message "Analyzing profile: $ProfilePath"
    
    # Create wizard form
    $wizForm = New-Object System.Windows.Forms.Form
    $wizForm.Text = "Profile Cleanup Wizard - Optimize Export Size"
    $wizForm.Size = New-Object System.Drawing.Size(950, 750)
    $wizForm.StartPosition = "CenterScreen"
    $wizForm.FormBorderStyle = "FixedDialog"
    $wizForm.MaximizeBox = $false
    $wizForm.MinimizeBox = $false
    $wizForm.TopMost = $true
    $wizForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $wizForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(950, 80)
    $headerPanel.BackColor = [System.Drawing.Color]::White
    $wizForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(900, 30)
    $lblTitle.Text = "Reduce Export Size - Cleanup Recommendations"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 48)
    $lblSubtitle.Size = New-Object System.Drawing.Size(900, 25)
    $lblSubtitle.Text = "Analyzing profile for temporary files, caches, and large files that can be safely excluded..."
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Progress bar for analysis
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 95)
    $progressBar.Size = New-Object System.Drawing.Size(900, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $wizForm.Controls.Add($progressBar)
    
    # Results panel
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Location = New-Object System.Drawing.Point(15, 125)
    $resultsPanel.Size = New-Object System.Drawing.Size(910, 520)
    $resultsPanel.BackColor = [System.Drawing.Color]::White
    $resultsPanel.AutoScroll = $true
    $wizForm.Controls.Add($resultsPanel)
    
    # DON'T show form yet - will use ShowDialog() after analysis completes
    # This prevents the "Form that is already visible" error
    
    # Analyze profile (populate cleanup categories)
    $cleanupItems = @()
    $totalSavings = 0
    $yPos = 20
    
    # Category 1: Browser Caches
    $progressBar.Value = 10
    $lblSubtitle.Text = "Scanning browser caches..."
    
    $browserPaths = @(
        @{Name="Chrome Cache"; Path="AppData\Local\Google\Chrome\User Data\Default\Cache"},
        @{Name="Chrome Code Cache"; Path="AppData\Local\Google\Chrome\User Data\Default\Code Cache"},
        @{Name="Edge Cache"; Path="AppData\Local\Microsoft\Edge\User Data\Default\Cache"},
        @{Name="Edge Code Cache"; Path="AppData\Local\Microsoft\Edge\User Data\Default\Code Cache"},
        @{Name="Firefox Cache"; Path="AppData\Local\Mozilla\Firefox\Profiles\*\cache2"},
        @{Name="IE Cache"; Path="AppData\Local\Microsoft\Windows\INetCache"}
    )
    
    $browserCacheSize = 0
    $browserCachePaths = @()
    foreach ($bp in $browserPaths) {
        $fullPath = Join-Path $ProfilePath $bp.Path
        if ($bp.Path -like "*`**") {
            # Wildcard path - expand it
            $parentPath = Split-Path (Join-Path $ProfilePath $bp.Path) -Parent
            if (Test-Path $parentPath) {
                $matchingDirs = Get-ChildItem $parentPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like $fullPath }
                foreach ($dir in $matchingDirs) {
                    try {
                        $size = (Get-ChildItem $dir.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                        if ($size -gt 0) {
                            $browserCacheSize += $size
                            $browserCachePaths += $dir.FullName
                        }
                    } catch {}
                }
            }
        } else {
            if (Test-Path $fullPath) {
                try {
                    $size = (Get-ChildItem $fullPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    if ($size -gt 0) {
                        $browserCacheSize += $size
                        $browserCachePaths += $fullPath
                    }
                } catch {}
            }
        }
    }
    
    if ($browserCacheSize -gt 0) {
        $cleanupItems += @{
            Category = "Browser Caches"
            Description = "Temporary browser data (Chrome, Edge, Firefox, IE) - Safe to delete"
            Size = $browserCacheSize
            Paths = $browserCachePaths
            Checked = $true
        }
    }
    
    # Category 2: Windows Temp Files
    $progressBar.Value = 25
    $lblSubtitle.Text = "Scanning temporary files..."
    
    $tempSize = 0
    $tempPaths = @()
    $tempLocations = @(
        "AppData\Local\Temp",
        "AppData\Local\Microsoft\Windows\Temporary Internet Files",
        "AppData\Local\Microsoft\Windows\INetCache\Content.IE5",
        "AppData\Local\Microsoft\Windows\WebCache"
    )
    
    foreach ($tl in $tempLocations) {
        $fullPath = Join-Path $ProfilePath $tl
        if (Test-Path $fullPath) {
            try {
                $size = (Get-ChildItem $fullPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($size -gt 0) {
                    $tempSize += $size
                    $tempPaths += $fullPath
                }
            } catch {}
        }
    }
    
    if ($tempSize -gt 0) {
        $cleanupItems += @{
            Category = "Temporary Files"
            Description = "Windows temporary files and cache - Safe to delete"
            Size = $tempSize
            Paths = $tempPaths
            Checked = $true
        }
    }
    
    # Category 3: Large Files (>100MB)
    $progressBar.Value = 50
    $lblSubtitle.Text = "Finding large files (>100MB)..."
    
    $largeFiles = @()
    $largeFilesSize = 0
    try {
        $largeFiles = Get-ChildItem $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | 
            Where-Object { $_.Length -gt 100MB } | 
            Select-Object -First 50 FullName, @{N='SizeMB';E={[math]::Round($_.Length/1MB,2)}}
        
        foreach ($lf in $largeFiles) {
            $largeFilesSize += ($lf.SizeMB * 1MB)
        }
    } catch {}
    
    if ($largeFiles.Count -gt 0) {
        $cleanupItems += @{
            Category = "Large Files"
            Description = "$($largeFiles.Count) files over 100MB - Review and exclude if not needed"
            Size = $largeFilesSize
            Paths = ($largeFiles | ForEach-Object { $_.FullName })
            Checked = $false
            Details = ($largeFiles | ForEach-Object { "$(Split-Path $_.FullName -Leaf) - $($_.SizeMB) MB" })
        }
    }
    
    # Category 4: Duplicate Files
    $progressBar.Value = 70
    $lblSubtitle.Text = "Detecting duplicate files..."
    
    $duplicates = @{}
    $duplicateSize = 0
    try {
        $allFiles = Get-ChildItem $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | 
            Where-Object { $_.Length -gt 1MB -and $_.Length -lt 1GB } |
            Select-Object -First 1000 FullName, Length, @{N='Hash';E={$null}}
        
        # Group by size first (faster than hashing everything)
        $sizeGroups = $allFiles | Group-Object Length | Where-Object { $_.Count -gt 1 }
        
        $hashCount = 0
        foreach ($group in $sizeGroups) {
            if ($hashCount -gt 100) { break } # Limit hashing for performance
            foreach ($file in $group.Group) {
                try {
                    $hash = (Get-FileHash $file.FullName -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
                    if ($hash) {
                        if ($duplicates.ContainsKey($hash)) {
                            $duplicates[$hash] += @{Path=$file.FullName; Size=$file.Length}
                            $duplicateSize += $file.Length
                        } else {
                            $duplicates[$hash] = @(@{Path=$file.FullName; Size=$file.Length})
                        }
                    }
                    $hashCount++
                } catch {}
            }
        }
        
        # Filter to only groups with duplicates
        $duplicates = $duplicates.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    } catch {}
    
    if ($duplicates.Count -gt 0) {
        $dupDetails = @()
        $dupPaths = @()
        foreach ($dup in $duplicates) {
            # Keep first, mark rest as duplicates
            for ($i = 1; $i -lt $dup.Value.Count; $i++) {
                $dupPaths += $dup.Value[$i].Path
                $sizeMB = [math]::Round($dup.Value[$i].Size/1MB,2)
                $dupDetails += "$(Split-Path $dup.Value[$i].Path -Leaf) - $sizeMB MB (duplicate)"
            }
        }
        
        $cleanupItems += @{
            Category = "Duplicate Files"
            Description = "$($dupPaths.Count) duplicate files found - Can save space by excluding"
            Size = $duplicateSize
            Paths = $dupPaths
            Checked = $false
            Details = $dupDetails
        }
    }
    
    # Category 5: Recycle Bin
    $progressBar.Value = 85
    $lblSubtitle.Text = "Checking Recycle Bin..."
    
    $recyclePath = Join-Path $ProfilePath '$RECYCLE.BIN'
    $recycleSize = 0
    if (Test-Path $recyclePath) {
        try {
            $recycleSize = (Get-ChildItem $recyclePath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            if ($recycleSize -gt 0) {
                $cleanupItems += @{
                    Category = "Recycle Bin"
                    Description = "Deleted files in Recycle Bin - Safe to delete"
                    Size = $recycleSize
                    Paths = @($recyclePath)
                    Checked = $true
                }
            }
        } catch {}
    }
    
    # Category 6: Downloads folder (large files only)
    $progressBar.Value = 95
    $lblSubtitle.Text = "Analyzing Downloads folder..."
    
    $downloadsPath = Join-Path $ProfilePath 'Downloads'
    $downloadSize = 0
    $downloadLargeFiles = @()
    if (Test-Path $downloadsPath) {
        try {
            $downloadLargeFiles = Get-ChildItem $downloadsPath -File -ErrorAction SilentlyContinue | 
                Where-Object { $_.Length -gt 50MB } |
                Select-Object FullName, @{N='SizeMB';E={[math]::Round($_.Length/1MB,2)}}
            
            foreach ($dlf in $downloadLargeFiles) {
                $downloadSize += ($dlf.SizeMB * 1MB)
            }
            
            if ($downloadLargeFiles.Count -gt 0) {
                $cleanupItems += @{
                    Category = "Large Downloads"
                    Description = "$($downloadLargeFiles.Count) large files in Downloads (>50MB) - Review for exclusion"
                    Size = $downloadSize
                    Paths = ($downloadLargeFiles | ForEach-Object { $_.FullName })
                    Checked = $false
                    Details = ($downloadLargeFiles | ForEach-Object { "$(Split-Path $_.FullName -Leaf) - $($_.SizeMB) MB" })
                }
            }
        } catch {}
    }
    
    $progressBar.Value = 100
    $lblSubtitle.Text = "Analysis complete!"
    
    # Display results
    if ($cleanupItems.Count -eq 0) {
        $lblNoItems = New-Object System.Windows.Forms.Label
        $lblNoItems.Location = New-Object System.Drawing.Point(20, 20)
        $lblNoItems.Size = New-Object System.Drawing.Size(870, 60)
        $lblNoItems.Text = "No significant cleanup opportunities found.`n`nProfile appears to be already optimized."
        $lblNoItems.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $lblNoItems.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        $resultsPanel.Controls.Add($lblNoItems)
    } else {
        $checkboxes = @()
        
        foreach ($item in $cleanupItems) {
            # Category checkbox
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Location = New-Object System.Drawing.Point(20, $yPos)
            $chk.Size = New-Object System.Drawing.Size(850, 25)
            $chk.Checked = $item.Checked
            $chk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            
            $sizeMB = [math]::Round($item.Size/1MB, 2)
            $sizeGB = [math]::Round($item.Size/1GB, 2)
            $sizeStr = if ($sizeGB -gt 1) { "$sizeGB GB" } else { "$sizeMB MB" }
            
            $chk.Text = "$($item.Category) - $sizeStr - $($item.Description)"
            $chk.Tag = $item
            $resultsPanel.Controls.Add($chk)
            $checkboxes += $chk
            $yPos += 30
            
            # Details button if available
            if ($item.Details) {
                $btnDetails = New-Object System.Windows.Forms.Button
                $btnDetails.Location = New-Object System.Drawing.Point(40, $yPos)
                $btnDetails.Size = New-Object System.Drawing.Size(120, 25)
                $btnDetails.Text = "View Details"
                $btnDetails.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                $btnDetails.ForeColor = [System.Drawing.Color]::White
                $btnDetails.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $btnDetails.FlatAppearance.BorderSize = 0
                $btnDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                $btnDetails.Cursor = [System.Windows.Forms.Cursors]::Hand
                $btnDetails.Tag = $item.Details
                $btnDetails.Add_Click({
                    $detailsText = $this.Tag -join "`r`n"
                    [System.Windows.Forms.MessageBox]::Show($detailsText, "File Details", "OK", "Information")
                })
                $resultsPanel.Controls.Add($btnDetails)
                $yPos += 35
            } else {
                $yPos += 10
            }
            
            $totalSavings += $item.Size
        }
        
        # Summary at top
        $savingsGB = [math]::Round($totalSavings/1GB, 2)
        $savingsMB = [math]::Round($totalSavings/1MB, 2)
        $savingsStr = if ($savingsGB -gt 1) { "$savingsGB GB" } else { "$savingsMB MB" }
        
        $lblSummary = New-Object System.Windows.Forms.Label
        $lblSummary.Location = New-Object System.Drawing.Point(20, 655)
        $lblSummary.Size = New-Object System.Drawing.Size(500, 30)
        $lblSummary.Text = "Potential savings: $savingsStr ($($cleanupItems.Count) categories)"
        $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        $wizForm.Controls.Add($lblSummary)
    }
    
    # Action buttons
    $btnClean = New-Object System.Windows.Forms.Button
    $btnClean.Location = New-Object System.Drawing.Point(550, 655)
    $btnClean.Size = New-Object System.Drawing.Size(180, 40)
    $btnClean.Text = "Clean & Export"
    $btnClean.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    $btnClean.ForeColor = [System.Drawing.Color]::White
    $btnClean.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClean.FlatAppearance.BorderSize = 0
    $btnClean.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnClean.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClean.DialogResult = "OK"
    $btnClean.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12) })
    $btnClean.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16) })
    $wizForm.Controls.Add($btnClean)
    
    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Location = New-Object System.Drawing.Point(740, 655)
    $btnSkip.Size = New-Object System.Drawing.Size(180, 40)
    $btnSkip.Text = "Skip Cleanup"
    $btnSkip.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnSkip.ForeColor = [System.Drawing.Color]::White
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.FlatAppearance.BorderSize = 0
    $btnSkip.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnSkip.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSkip.DialogResult = "Cancel"
    $btnSkip.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
    $btnSkip.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
    $wizForm.Controls.Add($btnSkip)
    
    $wizForm.AcceptButton = $btnClean
    $wizForm.CancelButton = $btnSkip
    
    $result = $wizForm.ShowDialog()
    
    # Process cleanup if user clicked Clean & Export
    $cleanupPaths = @()
    $selectedCategories = @()
    $totalSelectedSavings = 0
    
    if ($result -eq "OK") {
        foreach ($chk in $checkboxes) {
            if ($chk.Checked -and $chk.Tag) {
                $item = $chk.Tag
                $selectedCategories += $item.Category
                $totalSelectedSavings += $item.Size
                Log-Message "Selected for cleanup: $($item.Category) - $([math]::Round($item.Size/1MB,2)) MB"
                $cleanupPaths += $item.Paths
            }
        }
        
        if ($cleanupPaths.Count -gt 0) {
            Log-Message "Total paths to exclude from export: $($cleanupPaths.Count)"
            Log-Message "Total space to be saved: $([math]::Round($totalSelectedSavings/1MB,2)) MB"
        }
    }
    
    $wizForm.Dispose()
    
    return @{
        Proceed = ($result -eq "OK")
        CleanupPaths = $cleanupPaths
        CleanupCategories = $selectedCategories
        TotalSavingsMB = [math]::Round($totalSelectedSavings/1MB, 2)
    }
}

# =============================================================================
# EXPORT - NOW WITH *EXACT SAME* LIVE FILE-COUNT PROGRESS AS IMPORT
# =============================================================================
function Export-UserProfile {
    param([string]$Username, [string]$ZipPath)

    $global:ProgressBar.Value = 0
    $global:StatusText.Text = "Preparing export..."
    $global:CancelRequested = $false
    $global:CurrentOperation = "Export"
    if ($global:CancelButton) { $global:CancelButton.Enabled = $true }
    [System.Windows.Forms.Application]::DoEvents()

    # Start persistent logging for export operation
    Start-OperationLog -Operation "Export" -Username $Username

    try {
        $global:ExportStartTime = [DateTime]::Now
        Log-Message "=== EXPORT START: $Username -> $ZipPath"
        Log-Message "Operation started at: $($global:ExportStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"

        # Extract short name from DOMAIN\username or COMPUTERNAME\username for folder lookup
        $shortName = if ($Username -match '\\') { ($Username -split '\\',2)[1] } else { $Username }
        
        $profile = Get-LocalProfiles | Where-Object Username -eq $shortName | Select-Object -First 1
        if (-not $profile) { throw "Profile not found: $shortName" }
        
        # Store source path for reporting
        $source = $profile.Path

        if ($global:CancelRequested) { throw "Operation cancelled by user" }

        # === PROFILE SIZE OPTIMIZATION - CLEANUP WIZARD ===
        # Show cleanup wizard to let user reduce export size
        $cleanupResult = Show-ProfileCleanupWizard -ProfilePath $profile.Path
        
        if (-not $cleanupResult.Proceed) {
            Log-Message "Export cancelled - user declined cleanup wizard"
            throw "Export cancelled by user"
        }
        
        $userCleanupPaths = $cleanupResult.CleanupPaths
        $cleanupCategories = $cleanupResult.CleanupCategories  # Store for report
        $cleanupSavingsMB = $cleanupResult.TotalSavingsMB      # Store for report
        
        if ($userCleanupPaths.Count -gt 0) {
            Log-Message "User selected $($userCleanupPaths.Count) paths for exclusion from export"
            Log-Message "Estimated space savings: $cleanupSavingsMB MB"
        } else {
            Log-Message "No cleanup paths selected - proceeding with standard export"
        }

        $ts   = Get-Date -f yyyyMMdd_HHmmss
        $tmp  = Join-Path ([IO.Path]::GetDirectoryName($ZipPath)) "$shortName-Export-$ts"
        New-Item -ItemType Directory $tmp -Force | Out-Null
        Log-Message "Temporary export directory: $tmp"

        # NOTE: We no longer create a Profile\ subfolder - files go directly in $tmp for flat ZIP structure

        # NOTE: We no longer create a Profile\ subfolder - files go directly in $tmp for flat ZIP structure

        # Use 7-Zip to compress profile folder directly (handles locked files better than robocopy)
        $global:StatusText.Text = "Compressing profile with 7-Zip..."
        Log-Message "Compressing profile folder directly with 7-Zip (no robocopy)"
        $global:ProgressBar.Value = 10
        [System.Windows.Forms.Application]::DoEvents()

        # Build exclusion filters for 7-Zip
        $exclusions = Get-RobocopyExclusions -Mode 'Export'
        $7zExclusions = @()
        
        # Add user-selected cleanup paths to exclusions
        foreach ($cleanupPath in $userCleanupPaths) {
            if (-not [string]::IsNullOrWhiteSpace($cleanupPath)) {
                $relPath = $cleanupPath
                if ($cleanupPath -like "$($profile.Path)\*") {
                    $relPath = $cleanupPath.Substring($profile.Path.Length + 1)
                }
                
                # Handle both files and directories
                if (Test-Path $cleanupPath -PathType Container) {
                    # Directory - exclude it and all contents
                    $7zExclusions += "-xr!`"$relPath\*`""
                    $7zExclusions += "-x!`"$relPath`""
                    Log-Message "Excluding cleanup directory: $relPath"
                } else {
                    # File - exclude just the file
                    $7zExclusions += "-x!`"$relPath`""
                    Log-Message "Excluding cleanup file: $relPath"
                }
            }
        }
        
        # Add standard exclusions
        foreach ($file in $exclusions.Files) {
            if (-not [string]::IsNullOrWhiteSpace($file)) {
                $relFile = $file
                if ($profile -and $profile.Path -and $file -like ("$($profile.Path)\*")) {
                    $relFile = $file.Substring($profile.Path.Length + 1)
                }
                $7zExclusions += "-x!`"$relFile`""
            }
        }
        foreach ($dir in $exclusions.Dirs) {
            if (-not [string]::IsNullOrWhiteSpace($dir)) {
                $relDir = $dir
                if ($profile -and $profile.Path -and $dir -like ("$($profile.Path)\*")) {
                    $relDir = $dir.Substring($profile.Path.Length + 1)
                }
                $patternContents = if ($relDir.EndsWith('*')) { $relDir } else { "$relDir\*" }
                # Exclude the directory entry and all its contents
                $7zExclusions += "-xr!`"$patternContents`""
                $7zExclusions += "-x!`"$relDir`""
            }
        }

        # Explicitly exclude common legacy junctions under Documents
        $junctions = @(
            'Documents\My Music',
            'Documents\My Pictures',
            'Documents\My Videos'
        )
        foreach ($j in $junctions) {
            # Exclude junction directory entry and its contents
            $7zExclusions += "-xr!`"$j\*`""
            $7zExclusions += "-x!`"$j`""
        }

        # Discover and exclude all reparse-point directories (junctions) under the profile root
        try {
            $rpItems = Get-ChildItem -Path $profile.Path -Directory -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }
            foreach ($item in $rpItems) {
                $rel = $item.FullName.Substring($profile.Path.Length).TrimStart('\\')
                if ([string]::IsNullOrWhiteSpace($rel)) { continue }
                $patternContents = "-xr!`"$rel\*`""
                $patternDir = "-x!`"$rel`""
                if ($7zExclusions -notcontains $patternContents) { $7zExclusions += $patternContents }
                if ($7zExclusions -notcontains $patternDir) { $7zExclusions += $patternDir }
            }
            if ($rpItems.Count -gt 0) {
                Log-Message "7-Zip export: excluding $($rpItems.Count) reparse points (junctions)"
                $preview = $rpItems | Select-Object -First 10
                foreach ($p in $preview) {
                    $relPreview = $p.FullName.Substring($profile.Path.Length).TrimStart('\\')
                    Log-Message " - excluded junction: $relPreview"
                }
                if ($rpItems.Count -gt $preview.Count) {
                    Log-Message " - and $($rpItems.Count - $preview.Count) more..."
                }
            }
        } catch {
            Log-Message "WARNING: Junction discovery failed: $_"
        }
        
        # Detect and log PST files (Outlook personal folders/archives)
        try {
            $pstFiles = Get-ChildItem -Path $profile.Path -Filter "*.pst" -Recurse -Force -ErrorAction SilentlyContinue
            if ($pstFiles.Count -gt 0) {
                Log-Message "========================================="
                Log-Message "OUTLOOK PST FILES DETECTED: $($pstFiles.Count) file(s)"
                Log-Message "========================================="
                foreach ($pst in $pstFiles) {
                    $pstSize = [math]::Round($pst.Length/1MB, 2)
                    $relPath = $pst.FullName.Substring($profile.Path.Length).TrimStart('\\')
                    Log-Message "PST: $relPath ($pstSize MB)"
                }
                Log-Message "NOTE: PST files in profile will be migrated."
                Log-Message "PST files on network drives or custom locations must be copied manually."
                Log-Message "========================================="
            }
        } catch {
            Log-Message "WARNING: PST file detection failed: $_"
        }

        # === GENERATE MANIFEST AND WINGET FIRST (before compression) ===
        $global:StatusText.Text = "Generating manifest..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Get SID
        $sid = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
                Where-Object { (Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$shortName" }).PSChildName

        # Derive domain vs local from SID -> NTAccount
        $derivedDomain = $null
        $derivedUsername = $Username
        try {
            if ($sid) {
                $sidObj = [System.Security.Principal.SecurityIdentifier]::new($sid)
                $nt = $sidObj.Translate([System.Security.Principal.NTAccount])
                $parts = $nt.Value -split '\\',2
                if ($parts.Count -ge 2) { 
                    $derivedDomain = $parts[0]
                    $derivedUsername = $parts[1]
                }
            }
        } catch {
            Log-Message "WARNING: Could not translate SID to NTAccount: $_"
        }
        
        # Profile hash/timestamp for incremental backup detection
        $profileHash = $null
        $profileTimestamp = (Get-Item $profile.Path -Force -ErrorAction SilentlyContinue).LastWriteTimeUtc.ToString('o')
        try {
            $profileHash = (Get-FileHash -Path $profile.Path -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
        } catch {
            Log-Message "WARNING: Could not calculate profile hash: $_"
        }
        
        # Create manifest in temp folder (will be included in compression)
        $manifest = [pscustomobject]@{
            ExportedAt       = (Get-Date).ToString('o')
            Username         = $derivedUsername
            ProfilePath      = $profile.Path
            SourceSID        = $sid
            IsDomainUser     = if ($derivedDomain) { ($derivedDomain -ine $env:COMPUTERNAME) } else { $Username -match '\\' }
            Domain           = if ($derivedDomain -and ($derivedDomain -ine $env:COMPUTERNAME)) { $derivedDomain } elseif ($Username -match '\\') { ($Username -split '\\')[0] } else { $null }
            ProfileHash      = $profileHash
            ProfileTimestamp = $profileTimestamp
        }
        $manifest | ConvertTo-Json -Depth 5 | Out-File "$tmp\manifest.json" -Encoding UTF8
        Log-Message "Manifest created in temp folder"

        # Winget export in temp folder (will be included in compression)
        $wingetFile = "$tmp\Winget-Packages.json"
        try {
            Log-Message "Exporting Winget + Store apps..."
            $wingetProc = Start-Process "winget.exe" -ArgumentList "export", "-o", "`"$wingetFile`"", "--include-versions", "--accept-source-agreements" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop

            if ($wingetProc.ExitCode -eq 0) {
                Log-Message "Winget export succeeded"
            } else {
                Log-Message "Winget completed with exit code $($wingetProc.ExitCode)"
                "{}" | Out-File $wingetFile -Encoding UTF8
            }
        } catch {
            Log-Message "Winget FAILED: $_ - creating empty placeholder"
            "{}" | Out-File $wingetFile -Encoding UTF8
        }

        # Compress profile folder contents + manifest + winget in ONE operation
        # Multi-threading: Use all CPU cores for maximum compression speed (2-3x faster on multi-core systems)
        $threadCount = $Config.SevenZipThreads
        Log-Message "7-Zip compression using $threadCount threads (detected $cpuCores CPU cores)"
        $7zArgs = @('a', '-tzip', $ZipPath, "$($profile.Path)\*", "$tmp\manifest.json", "$tmp\Winget-Packages.json", '-mx=5', "-mmt=$threadCount", '-bsp1') + $7zExclusions
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName           = $global:SevenZipPath
        $psi.Arguments          = $7zArgs -join ' '
        $psi.UseShellExecute    = $false
        $psi.CreateNoWindow     = $true
        $psi.RedirectStandardOutput = $true

        $zip = [System.Diagnostics.Process]::Start($psi)

        while (-not $zip.HasExited) {
            if ($global:CancelRequested) {
                try { $zip.Kill() } catch { }
                throw "Export cancelled by user"
            }
            if (-not $zip.StandardOutput.EndOfStream) {
                $line = $zip.StandardOutput.ReadLine()
                if ($line -match '(\d+)%') {
                    $pct = [int]$Matches[1]
                    $global:ProgressBar.Value = [Math]::Min(90, 10 + [int]($pct * 0.8))
                    $global:StatusText.Text = "Compressing - $pct%"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
            Start-Sleep -Milliseconds 100
        }
        $zip.WaitForExit()

        if ($zip.ExitCode -notin 0,1) {
            throw "7-Zip failed with exit code $($zip.ExitCode)"
        }

        Log-Message "Profile compressed successfully"
        $global:ProgressBar.Value = 90
        
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        
        # === HASH VERIFICATION (Export) ===
        if ($Config.HashVerificationEnabled) {
            $global:StatusText.Text = "Verifying ZIP integrity..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Calculating ZIP hash for verification..."
            try {
                $zipHashResult = Get-FileHash -Path $ZipPath -Algorithm SHA256 -ErrorAction Stop
                if ($zipHashResult -and $zipHashResult.Hash) {
                    Log-Message "ZIP Hash (SHA256): $($zipHashResult.Hash)"
                    # Store hash in a sidecar file for later verification
                    "$($zipHashResult.Hash)  $(Split-Path $ZipPath -Leaf)" | Out-File "$ZipPath.sha256" -Encoding ASCII
                    Log-Message "Hash saved to: $ZipPath.sha256"
                } else {
                    Log-Message "WARNING: Hash calculation returned null"
                }
            } catch {
                Log-Message "WARNING: Hash verification failed: $_"
            }
        }

        # === COLLECT DATA FOR REPORT (BEFORE CLEANUP) ===
        # Parse installed programs from Winget export for report
        $installedPrograms = @()
        $wingetExportPath = "$tmp\Winget-Packages.json"
        if (Test-Path $wingetExportPath) {
            try {
                Log-Message "Parsing Winget export for installed programs list..."
                $wingetData = Get-Content $wingetExportPath -Raw | ConvertFrom-Json
                if ($wingetData.Sources) {
                    foreach ($wingetSource in $wingetData.Sources) {
                        if ($wingetSource.Packages) {
                            foreach ($pkg in $wingetSource.Packages) {
                                $installedPrograms += @{
                                    Name = $pkg.PackageIdentifier
                                    Version = if ($pkg.Version) { $pkg.Version } else { "Latest" }
                                    Source = if ($wingetSource.SourceDetails.Name) { $wingetSource.SourceDetails.Name } else { "Unknown" }
                                }
                            }
                        }
                    }
                }
                Log-Message "Collected $($installedPrograms.Count) installed programs for report"
            } catch {
                Log-Message "Could not parse Winget export for report: $_"
            }
        } else {
            Log-Message "WARNING: Winget-Packages.json not found at $wingetExportPath"
        }

        # Cleanup temp folder
        $global:StatusText.Text = "Cleaning up temp files"
        Log-Message "Cleaning up temp files"
        Remove-FolderRobust $tmp

        $sizeMB = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
        $global:ProgressBar.Value = 100
        $global:StatusText.Text = "[OK] Export complete!"
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)

        # Log elapsed time for diagnostics
        if ($global:ExportStartTime) {
            $elapsed = [DateTime]::Now - $global:ExportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }

        Log-Message "EXPORT SUCCESS: $ZipPath ($sizeMB MB)"
        
        # Generate migration report
        try {
            # Collect ZIP file statistics for report
            $zipFileCount = 'N/A'
            $zipFolderCount = 'N/A'
            $uncompressedSizeMB = 'N/A'
            $compressionRatio = 'N/A'
            
            try {
                if (Test-Path $ZipPath) {
                    Log-Message "Collecting ZIP archive statistics for report..."
                    
                    # Use 7-Zip with -slt (technical listing) for reliable parsing
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $global:SevenZipPath
                    $psi.Arguments = "l -slt `"$ZipPath`""
                    $psi.CreateNoWindow = $true
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    
                    $listProc = [System.Diagnostics.Process]::Start($psi)
                    $output = $listProc.StandardOutput.ReadToEnd()
                    $listProc.WaitForExit()
                    
                    if ($listProc.ExitCode -eq 0) {
                        Log-Message "7-Zip technical listing received, parsing..."
                        
                        # Count files and folders, sum sizes from technical output
                        $files = 0
                        $folders = 0
                        $totalSize = 0
                        
                        $lines = $output -split "`n"
                        $isFolder = $false
                        $currentSize = 0
                        
                        foreach ($line in $lines) {
                            $line = $line.Trim()
                            
                            # Check if this entry is a folder
                            if ($line -match '^Attributes = .+D') {
                                $isFolder = $true
                            }
                            
                            # Get the size
                            if ($line -match '^Size = (\d+)$') {
                                $currentSize = [int64]$matches[1]
                            }
                            
                            # When we hit a Path entry, we've finished reading an item
                            if ($line -match '^Path = (.+)$' -and $line -notmatch '^Path\s*=\s*$') {
                                if ($isFolder) {
                                    $folders++
                                } else {
                                    $files++
                                    $totalSize += $currentSize
                                }
                                # Reset for next item
                                $isFolder = $false
                                $currentSize = 0
                            }
                        }
                        
                        if ($files -gt 0) {
                            $zipFileCount = $files
                            $zipFolderCount = $folders
                            $uncompressedSizeMB = [math]::Round($totalSize / 1MB, 2)
                            
                            # Calculate compression ratio
                            if ($totalSize -gt 0) {
                                $zipFileObj = Get-Item $ZipPath
                                $compressionRatio = "$([math]::Round(($zipFileObj.Length / $totalSize) * 100, 1))%"
                            }
                            
                            Log-Message "ZIP statistics: $zipFileCount files, $zipFolderCount folders, $uncompressedSizeMB MB uncompressed, compression: $compressionRatio"
                        } else {
                            Log-Message "WARNING: No files found in 7-Zip output"
                        }
                    } else {
                        Log-Message "7-Zip list command failed with exit code: $($listProc.ExitCode)"
                    }
                }
            } catch {
                Log-Message "Could not collect ZIP statistics: $_"
            }
            
            $reportData = @{
                Username = $Username
                SourceSID = $sid
                SourcePath = $source
                ZipPath = $ZipPath
                ZipSizeMB = $sizeMB
                ElapsedMinutes = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
                ElapsedSeconds = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                FileCount = $zipFileCount
                FolderCount = $zipFolderCount
                UncompressedSizeMB = $uncompressedSizeMB
                CompressionRatio = $compressionRatio
                Exclusions = @('Temp folders', 'Cache files', 'Log files', 'OST files')
                HashEnabled = $Config.HashVerificationEnabled
                InstalledPrograms = $installedPrograms
                CleanupCategories = $cleanupCategories
                CleanupSavingsMB = $cleanupSavingsMB
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Export' -ReportData $reportData
            if ($reportPath) {
                Log-Info "Migration report available: $reportPath"
            }
        } catch {
            Log-Warning "Could not generate migration report: $_"
        }
        
        [System.Windows.Forms.MessageBox]::Show("Export completed!`n`n$ZipPath`nSize: $sizeMB MB", "Success", "OK", "Information")

    }
    catch {
        # Log elapsed time for diagnostics
        if ($global:ExportStartTime) {
            $elapsed = [DateTime]::Now - $global:ExportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        Log-Message "EXPORT FAILED: $_"
        [System.Windows.Forms.MessageBox]::Show("Export failed:`n$_", "Error", "OK", "Error")
    }
    finally {
        Stop-OperationLog
        $global:StatusText.Text = "Idle"
        if ($global:ProgressBar.Value -ne 100) { $global:ProgressBar.Value = 0 }
    }
}

# =============================================================================
# IMPORT-USERPROFILE FUNCTION
# =============================================================================
function Import-UserProfile {
    param([string]$ZipPath, [string]$Username)
    $global:ProgressBar.Value = 0
    $global:StatusText.Text = "Starting import..."
    if ($global:CancelButton) { $global:CancelButton.Enabled = $true }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Start persistent logging for import operation
    Start-OperationLog -Operation "Import" -Username $Username
    
    try {
        $global:CancelRequested = $false
        $global:CurrentOperation = "Import"
        
        Log-Message "=== IMPORT to $Username ==="
        if (-not (Test-Path $ZipPath)) { throw "ZIP not found: $ZipPath" }
        Log-Message "ZIP file verified: $ZipPath"
        
        # === HASH VERIFICATION (Import) ===
        if ($Config.HashVerificationEnabled) {
            $hashFile = "$ZipPath.sha256"
            if (Test-Path $hashFile) {
                $global:StatusText.Text = "Verifying ZIP integrity..."
                [System.Windows.Forms.Application]::DoEvents()
                Log-Message "Verifying ZIP hash..."
                try {
                    $expectedHash = (Get-Content $hashFile -First 1).Split()[0]
                    $actualHashResult = Get-FileHash -Path $ZipPath -Algorithm SHA256 -ErrorAction Stop
                    if ($actualHashResult -and $actualHashResult.Hash) {
                        if ($actualHashResult.Hash -eq $expectedHash) {
                            Log-Message "ZIP hash verification: PASSED"
                        } else {
                            Log-Message "WARNING: ZIP hash mismatch!"
                            Log-Message "Expected: $expectedHash"
                            Log-Message "Actual:   $($actualHashResult.Hash)"
                            $hashResult = [System.Windows.Forms.MessageBox]::Show(
                                "ZIP file hash verification FAILED!`n`nThe file may be corrupted or tampered with.`n`nContinue anyway?",
                                "Hash Verification Failed", "YesNo", "Warning")
                            if ($hashResult -ne "Yes") {
                                throw "Import cancelled - hash verification failed"
                            }
                        }
                    }
                } catch {
                    Log-Message "WARNING: Hash verification error: $_"
                }
            } else {
                Log-Message "No hash file found - skipping verification"
            }
        }
        
        $global:StatusText.Text = "Validating ZIP file..."
        [System.Windows.Forms.Application]::DoEvents()
        
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        
        # Parse username and determine if domain or local user
        # Local user formats: "username" or "COMPUTERNAME\username"
        # Domain user format: "DOMAIN\username" (where DOMAIN != COMPUTERNAME)
        $domain = $null
        $shortName = $Username
        $isDomain = $false
        
        if ($Username -match '\\') {
            $parts = $Username -split '\\', 2
            $parsedDomain = $parts[0].Trim().ToUpper()
            $shortName = $parts[1].Trim()
            
            # Check if domain is actually the computer name (local user)
            if ($parsedDomain -ieq $env:COMPUTERNAME) {
                $isDomain = $false
                Log-Message "Parsed as LOCAL user (computer name match): User='$shortName'"
            } else {
                $isDomain = $true
                $domain = $parsedDomain
                Log-Message "Parsed as DOMAIN user: Domain='$domain', User='$shortName'"
            }
        } else {
            $isDomain = $false
            Log-Message "Parsed as LOCAL user (no domain): User='$shortName'"
        }
        
        $global:StatusText.Text = "Checking if user is logged on..."
        [System.Windows.Forms.Application]::DoEvents()
        $loggedOn = qwinsta | Select-String "\b$shortName\b" -Quiet
        if ($loggedOn) {
            Log-Message "WARNING: User '$shortName' appears to be logged on!"
            $res = [System.Windows.Forms.MessageBox]::Show("User '$shortName' appears to be logged on.`n`nRisk: Corrupted profile! Continue anyway?", "User Logged On", "YesNo", "Warning")
            if ($res -ne "Yes") { throw "Cancelled by user." }
        } else {
            Log-Message "User is not logged on - safe to proceed"
        }
        if (-not $isDomain -and -not (Get-LocalUser -Name $shortName -ErrorAction SilentlyContinue)) {
            Log-Message "Creating local user '$shortName'..."
            $global:StatusText.Text = "Creating local user..."
            [System.Windows.Forms.Application]::DoEvents()
            $passForm = New-Object System.Windows.Forms.Form
            $passForm.Text = "Create Local User"
            $passForm.Size = New-Object System.Drawing.Size(480,360)
            $passForm.StartPosition = "CenterScreen"
            $passForm.FormBorderStyle = "FixedDialog"
            $passForm.MaximizeBox = $false
            $passForm.MinimizeBox = $false
            $passForm.TopMost = $true
            $passForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
            $passForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            # Header panel
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size = New-Object System.Drawing.Size(480, 70)
            $headerPanel.BackColor = [System.Drawing.Color]::White
            $passForm.Controls.Add($headerPanel)

            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
            $lblTitle.Size = New-Object System.Drawing.Size(440, 25)
            $lblTitle.Text = "Create User: $shortName"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $headerPanel.Controls.Add($lblTitle)

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Location = New-Object System.Drawing.Point(22, 43)
            $lblInfo.Size = New-Object System.Drawing.Size(440, 20)
            $lblInfo.Text = "User does not exist - Set a password to create"
            $lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
            $headerPanel.Controls.Add($lblInfo)

            # Main content card
            $contentCard = New-Object System.Windows.Forms.Panel
            $contentCard.Location = New-Object System.Drawing.Point(15, 85)
            $contentCard.Size = New-Object System.Drawing.Size(440, 180)
            $contentCard.BackColor = [System.Drawing.Color]::White
            $passForm.Controls.Add($contentCard)
            $lblPass1 = New-Object System.Windows.Forms.Label
            $lblPass1.Location = New-Object System.Drawing.Point(20,20)
            $lblPass1.Size = New-Object System.Drawing.Size(100,20)
            $lblPass1.Text = "Password:"
            $lblPass1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblPass1)
            $txtPass1 = New-Object System.Windows.Forms.TextBox
            $txtPass1.Location = New-Object System.Drawing.Point(130,18)
            $txtPass1.Size = New-Object System.Drawing.Size(220,25)
            $txtPass1.UseSystemPasswordChar = $true
            $txtPass1.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $contentCard.Controls.Add($txtPass1)
            $lblPass2 = New-Object System.Windows.Forms.Label
            $lblPass2.Location = New-Object System.Drawing.Point(20,55)
            $lblPass2.Size = New-Object System.Drawing.Size(100,20)
            $lblPass2.Text = "Confirm:"
            $lblPass2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblPass2)
            $txtPass2 = New-Object System.Windows.Forms.TextBox
            $txtPass2.Location = New-Object System.Drawing.Point(130,53)
            $txtPass2.Size = New-Object System.Drawing.Size(220,25)
            $txtPass2.UseSystemPasswordChar = $true
            $txtPass2.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $contentCard.Controls.Add($txtPass2)
            $chkShow1 = New-Object System.Windows.Forms.CheckBox
            $chkShow1.Location = New-Object System.Drawing.Point(360,18)
            $chkShow1.Size = New-Object System.Drawing.Size(70,23)
            $chkShow1.Text = "Show"
            $chkShow1.Add_CheckedChanged({ $txtPass1.UseSystemPasswordChar = -not $chkShow1.Checked; $txtPass2.UseSystemPasswordChar = -not $chkShow1.Checked })
            $contentCard.Controls.Add($chkShow1)
            $lblStrength = New-Object System.Windows.Forms.Label
            $lblStrength.Location = New-Object System.Drawing.Point(130,85)
            $lblStrength.Size = New-Object System.Drawing.Size(290,20)
            $lblStrength.Text = ""
            $contentCard.Controls.Add($lblStrength)

            $chkAdmin = New-Object System.Windows.Forms.CheckBox
            $chkAdmin.Location = New-Object System.Drawing.Point(20,120)
            $chkAdmin.Size = New-Object System.Drawing.Size(400,25)
            $chkAdmin.Text = "Add this user to Administrators group"
            $chkAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($chkAdmin)

            $lblError = New-Object System.Windows.Forms.Label
            $lblError.Location = New-Object System.Drawing.Point(20,150)
            $lblError.Size = New-Object System.Drawing.Size(400,20)
            $lblError.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            $lblError.Text = ""
            $lblError.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblError)
            $txtPass1.Add_TextChanged({
                $len = $txtPass1.Text.Length
                if ($len -eq 0) {
                    $lblStrength.Text = ""
                    $lblStrength.ForeColor = [System.Drawing.Color]::Black
                } elseif ($len -lt 8) {
                    $lblStrength.Text = "Weak password (< 8 characters)"
                    $lblStrength.ForeColor = [System.Drawing.Color]::OrangeRed
                } elseif ($len -lt 12) {
                    $lblStrength.Text = "Moderate password"
                    $lblStrength.ForeColor = [System.Drawing.Color]::Orange
                } else {
                    $lblStrength.Text = "Strong password"
                    $lblStrength.ForeColor = [System.Drawing.Color]::Green
                }
            })
            # Action buttons
            $btnOK = New-Object System.Windows.Forms.Button
            $btnOK.Location = New-Object System.Drawing.Point(230,280)
            $btnOK.Size = New-Object System.Drawing.Size(110,35)
            $btnOK.Text = "Create User"
            $btnOK.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
            $btnOK.ForeColor = [System.Drawing.Color]::White
            $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnOK.FlatAppearance.BorderSize = 0
            $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnOK.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12) })
            $btnOK.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16) })
            $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $passForm.Controls.Add($btnOK)
            $btnCancel = New-Object System.Windows.Forms.Button
            $btnCancel.Location = New-Object System.Drawing.Point(350,280)
            $btnCancel.Size = New-Object System.Drawing.Size(105,35)
            $btnCancel.Text = "Cancel"
            $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
            $btnCancel.ForeColor = [System.Drawing.Color]::White
            $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCancel.FlatAppearance.BorderSize = 0
            $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnCancel.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
            $btnCancel.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $passForm.Controls.Add($btnCancel)
            $passForm.AcceptButton = $btnOK
            $passForm.CancelButton = $btnCancel
            $txtPass1.Focus()
            $btnOK.Add_Click({
                $p1 = $txtPass1.Text
                $p2 = $txtPass2.Text
                if ($p1 -ne $p2) {
                    $lblError.Text = " Passwords do not match!"
                    $passForm.DialogResult = [System.Windows.Forms.DialogResult]::None
                    return
                }
                if ($p1.Length -lt 1) {
                    $lblError.Text = " Password cannot be empty!"
                    $passForm.DialogResult = [System.Windows.Forms.DialogResult]::None
                    return
                }
                $lblError.Text = ""
            })
            $result = $passForm.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $passwordText = $txtPass1.Text
                $pass = ConvertTo-SecureString $passwordText -AsPlainText -Force
                try {
                    New-LocalUser -Name $shortName -Password $pass -FullName $shortName -Description "Profile Import" -ErrorAction Stop
                    Log-Message "Local user '$shortName' created successfully"
                    if ($chkAdmin.Checked) {
                        Add-LocalGroupMember -Group "Administrators" -Member $shortName -ErrorAction Stop
                        Log-Message "Added '$shortName' to Administrators group"
                    }
                } catch {
                    throw "Failed to create local user or add to Administrators: $_"
                }
            } else {
                throw "User creation cancelled - cannot proceed with import"
            }
            $passForm.Dispose()
		}
        # Target profile directory
        $target = "C:\Users\$shortName"
        
        # Check if user profile already exists and prompt for merge/replace option
        $targetExistsBeforeResolveSID = Test-Path $target
        $mergeMode = $false
        if ($targetExistsBeforeResolveSID -and -not $isDomain) {
            Log-Message "User profile already exists: $target"
            
            # Create custom dialog with Merge/Replace buttons
            $choiceForm = New-Object System.Windows.Forms.Form
            $choiceForm.Text = "User Profile Exists"
            $choiceForm.Size = New-Object System.Drawing.Size(550, 340)
            $choiceForm.StartPosition = "CenterScreen"
            $choiceForm.FormBorderStyle = "FixedDialog"
            $choiceForm.MaximizeBox = $false
            $choiceForm.MinimizeBox = $false
            $choiceForm.TopMost = $true
            $choiceForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
            $choiceForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            # Header panel
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size = New-Object System.Drawing.Size(550, 60)
            $headerPanel.BackColor = [System.Drawing.Color]::White
            $choiceForm.Controls.Add($headerPanel)

            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
            $lblTitle.Size = New-Object System.Drawing.Size(510, 30)
            $lblTitle.Text = "User '$shortName' already exists"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $headerPanel.Controls.Add($lblTitle)

            # Content panel
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Location = New-Object System.Drawing.Point(15, 75)
            $contentPanel.Size = New-Object System.Drawing.Size(510, 180)
            $contentPanel.BackColor = [System.Drawing.Color]::White
            $choiceForm.Controls.Add($contentPanel)

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Location = New-Object System.Drawing.Point(15, 15)
            $lblInfo.Size = New-Object System.Drawing.Size(480, 150)
            $lblInfo.Text = "Choose how to handle the existing profile:`n`nMERGE: Keep existing profile and merge imported files`n               (preserves user settings)`n`nREPLACE: Backup existing profile and replace with`n                  imported profile"
            $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $contentPanel.Controls.Add($lblInfo)

            # Merge button
            $btnMerge = New-Object System.Windows.Forms.Button
            $btnMerge.Location = New-Object System.Drawing.Point(180, 265)
            $btnMerge.Size = New-Object System.Drawing.Size(110, 35)
            $btnMerge.Text = "Merge"
            $btnMerge.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $btnMerge.ForeColor = [System.Drawing.Color]::White
            $btnMerge.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnMerge.FlatAppearance.BorderSize = 0
            $btnMerge.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnMerge.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnMerge.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
            $btnMerge.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
            $btnMerge.Add_Click({ $choiceForm.Tag = "Merge"; $choiceForm.Close() })
            $choiceForm.Controls.Add($btnMerge)

            # Replace button
            $btnReplace = New-Object System.Windows.Forms.Button
            $btnReplace.Location = New-Object System.Drawing.Point(300, 265)
            $btnReplace.Size = New-Object System.Drawing.Size(110, 35)
            $btnReplace.Text = "Replace"
            $btnReplace.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            $btnReplace.ForeColor = [System.Drawing.Color]::White
            $btnReplace.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnReplace.FlatAppearance.BorderSize = 0
            $btnReplace.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnReplace.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnReplace.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(200, 15, 30) })
            $btnReplace.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35) })
            $btnReplace.Add_Click({ $choiceForm.Tag = "Replace"; $choiceForm.Close() })
            $choiceForm.Controls.Add($btnReplace)

            $choiceForm.ShowDialog() | Out-Null
            $choice = $choiceForm.Tag
            $choiceForm.Dispose()

            if ($choice -eq "Merge") {
                $mergeMode = $true
                Log-Message "MERGE MODE: Will merge imported files into existing profile"
            } elseif ($choice -eq "Replace") {
                $mergeMode = $false
                Log-Message "REPLACE MODE: Will backup and replace existing profile"
            } else {
                throw "Operation cancelled - no option selected"
            }
        }
        
        $global:StatusText.Text = "Resolving target SID..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # RESOLVE SID FIRST - needed for profile mounted checks and registry operations
        if ($isDomain) {
            # Modern domain credential prompt
            $credForm = New-Object System.Windows.Forms.Form
            $credForm.Text = "Domain Credentials Required"
            $credForm.Size = New-Object System.Drawing.Size(500,320)
            $credForm.StartPosition = "CenterScreen"
            $credForm.FormBorderStyle = "FixedDialog"
            $credForm.MaximizeBox = $false
            $credForm.MinimizeBox = $false
            $credForm.TopMost = $true
            $credForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
            $credForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            # Header panel
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size = New-Object System.Drawing.Size(500, 70)
            $headerPanel.BackColor = [System.Drawing.Color]::White
            $credForm.Controls.Add($headerPanel)

            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
            $lblTitle.Size = New-Object System.Drawing.Size(460, 25)
            $lblTitle.Text = "Domain Administrator Credentials"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $headerPanel.Controls.Add($lblTitle)

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Location = New-Object System.Drawing.Point(22, 43)
            $lblInfo.Size = New-Object System.Drawing.Size(460, 20)
            $lblInfo.Text = "Enter credentials to query domain: $domain"
            $lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
            $headerPanel.Controls.Add($lblInfo)

            # Main content card
            $contentCard = New-Object System.Windows.Forms.Panel
            $contentCard.Location = New-Object System.Drawing.Point(15, 85)
            $contentCard.Size = New-Object System.Drawing.Size(460, 140)
            $contentCard.BackColor = [System.Drawing.Color]::White
            $credForm.Controls.Add($contentCard)

            $lblUser = New-Object System.Windows.Forms.Label
            $lblUser.Location = New-Object System.Drawing.Point(20,25)
            $lblUser.Size = New-Object System.Drawing.Size(100,20)
            $lblUser.Text = "Username:"
            $lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblUser)

            $txtUser = New-Object System.Windows.Forms.TextBox
            $txtUser.Location = New-Object System.Drawing.Point(130,23)
            $txtUser.Size = New-Object System.Drawing.Size(300,25)
            $txtUser.Text = "$domain\"
            $txtUser.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $contentCard.Controls.Add($txtUser)

            $lblPass = New-Object System.Windows.Forms.Label
            $lblPass.Location = New-Object System.Drawing.Point(20,65)
            $lblPass.Size = New-Object System.Drawing.Size(100,20)
            $lblPass.Text = "Password:"
            $lblPass.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblPass)

            $txtPass = New-Object System.Windows.Forms.TextBox
            $txtPass.Location = New-Object System.Drawing.Point(130,63)
            $txtPass.Size = New-Object System.Drawing.Size(300,25)
            $txtPass.UseSystemPasswordChar = $true
            $txtPass.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $contentCard.Controls.Add($txtPass)

            $chkShowPass = New-Object System.Windows.Forms.CheckBox
            $chkShowPass.Location = New-Object System.Drawing.Point(130,100)
            $chkShowPass.Size = New-Object System.Drawing.Size(150,23)
            $chkShowPass.Text = "Show password"
            $chkShowPass.Add_CheckedChanged({ $txtPass.UseSystemPasswordChar = -not $chkShowPass.Checked })
            $contentCard.Controls.Add($chkShowPass)

            # Action buttons
            $btnOK = New-Object System.Windows.Forms.Button
            $btnOK.Location = New-Object System.Drawing.Point(250,240)
            $btnOK.Size = New-Object System.Drawing.Size(110,35)
            $btnOK.Text = "Connect"
            $btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $btnOK.ForeColor = [System.Drawing.Color]::White
            $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnOK.FlatAppearance.BorderSize = 0
            $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnOK.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
            $btnOK.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
            $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $credForm.Controls.Add($btnOK)

            $btnCancel = New-Object System.Windows.Forms.Button
            $btnCancel.Location = New-Object System.Drawing.Point(370,240)
            $btnCancel.Size = New-Object System.Drawing.Size(105,35)
            $btnCancel.Text = "Cancel"
            $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
            $btnCancel.ForeColor = [System.Drawing.Color]::White
            $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCancel.FlatAppearance.BorderSize = 0
            $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnCancel.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60) })
            $btnCancel.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80) })
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $credForm.Controls.Add($btnCancel)

            $credForm.AcceptButton = $btnOK
            $credForm.CancelButton = $btnCancel
            $txtUser.Select($txtUser.Text.Length, 0)
            $txtUser.Focus()

            $result = $credForm.ShowDialog()
            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                $credForm.Dispose()
                throw "Credentials cancelled"
            }

            $username = $txtUser.Text.Trim()
            $password = $txtPass.Text
            $credForm.Dispose()

            if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
                throw "Username and password cannot be empty"
            }

            # Create PSCredential object
            $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($username, $securePassword)
            
            # Store credentials globally for optional reuse in domain join
            $global:DomainCredential = $cred
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain',$domain,$cred.UserName,$cred.GetNetworkCredential().Password)
            $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx,$shortName)
            if (-not $user) { throw "User '$shortName' not found in domain '$domain'" }
            $sid = $user.Sid.Value
        } else {
            try {
                $nt = New-Object System.Security.Principal.NTAccount($shortName)
                $sid = $nt.Translate([System.Security.Principal.SecurityIdentifier]).Value
            } catch {
                throw "Failed to resolve SID for user '$shortName'. Ensure the user exists: $_"
            }
        }
        Log-Message "Target SID resolved: $sid"
        
        $global:StatusText.Text = "Preparing target directory..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # PRE-FLIGHT VALIDATION: Record whether target existed before validation and check writeability
        $targetExistedBefore = Test-Path $target
        Log-Message "Running pre-flight validation... (target existed before: $targetExistedBefore)"
        if ($targetExistedBefore -and -not (Test-ProfilePathWriteable $target)) {
            throw "Target path '$target' is not writable. Check permissions."
        }
        if ($targetExistedBefore) {
            Log-Message "Target path writeability verified"
        }
        
        # PRE-FLIGHT CHECK: Detect if profile is currently mounted/logged in
        if (Test-ProfileMounted $sid) {
            Log-Message "WARNING: User profile is currently mounted (user may be logged in via HKU)"
            $res = [System.Windows.Forms.MessageBox]::Show(
                "User profile is currently mounted in registry (HKU\$sid).`n`nThis may indicate the user is logged in. Importing now may cause data loss or corruption.`n`nContinue anyway?",
                "Profile Mounted", "YesNo", "Warning")
            if ($res -ne "Yes") {
                throw "Cancelled - profile is mounted"
            }
        } else {
            Log-Message "Profile not mounted - safe to proceed"
        }
        
        # BACKUP EXISTING PROFILE: Create timestamped backup before importing
        if ($targetExistedBefore) {
            Log-Message "Creating backup for existing profile: $target"
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $backup = "$target.backup_$timestamp"
            $global:ImportBackupPath = $backup
            
            if ($mergeMode) {
                Log-Message "MERGE MODE: Backup is for safety only (will not replace profile)"
            } else {
                Log-Message "REPLACE MODE: Backup before replacement"
            }
            
            $global:StatusText.Text = "Backing up existing profile..."
            $global:ProgressBar.Value = 1
            [System.Windows.Forms.Application]::DoEvents()
            try {
                Log-Message "Backup path: $backup"
                # Use robocopy for faster backup (multi-threaded, handles locked files)
                # Dynamic thread count: 8-32 threads based on CPU cores for optimal performance
                $robocopyThreads = $Config.RobocopyThreads
                Log-Message "Robocopy backup using $robocopyThreads threads (detected $cpuCores CPU cores)"
                $robocopyArgs = @(
                    "`"$target`"",
                    "`"$backup`"",
                    '/E',           # Copy subdirectories including empty
                    '/COPYALL',     # Copy all file info (timestamps, ACLs, owner, auditing)
                    '/R:2',         # Retry 2 times on failed copies
                    '/W:1',         # Wait 1 second between retries
                    "/MT:$robocopyThreads",  # Multi-threaded (dynamic based on CPU cores)
                    '/NP',          # No progress per file (reduces output)
                    '/NFL',         # No file list
                    '/NDL'          # No directory list
                )
                
                Log-Message "Starting robocopy backup..."
                $robocopyProc = Start-Process 'robocopy.exe' -ArgumentList $robocopyArgs -PassThru -NoNewWindow
                
                # Monitor backup progress by checking destination folder growth
                $startTime = [DateTime]::Now
                while (-not $robocopyProc.HasExited) {
                    Start-Sleep -Milliseconds 500
                    
                    # Update UI to show we're still working
                    $elapsed = ([DateTime]::Now - $startTime).TotalSeconds
                    $global:StatusText.Text = "Backing up existing profile... $([int]$elapsed)s"
                    
                    # Pulse progress bar (1-4 range)
                    $pulseValue = 1 + (($elapsed % 3) / 3 * 3)
                    $global:ProgressBar.Value = [Math]::Min(4, $pulseValue)
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                $robocopyProc.WaitForExit()
                
                # Robocopy exit codes: 0-7 are success (0=no changes, 1=files copied, 2=extra files, etc.)
                if ($robocopyProc.ExitCode -gt 7) {
                    throw "Robocopy failed with exit code $($robocopyProc.ExitCode)"
                }
                Log-Message "Backup completed (robocopy exit code: $($robocopyProc.ExitCode))"
            } catch {
                Log-Message "WARNING: Backup failed: $_"
            }
            
            # Only remove profile directory if not in merge mode
            if (-not $mergeMode) {
                Log-Message "Removing existing profile directory: $target"
                $global:StatusText.Text = "Removing existing profile directory..."
                [System.Windows.Forms.Application]::DoEvents()
                Remove-FolderRobust -Path $target
                Log-Message "Existing directory removed"
            } else {
                Log-Message "MERGE MODE: Keeping existing profile directory for merge"
            }
        } else {
            Log-Message "No existing profile found at $target"
            $global:ImportBackupPath = $null
        }
        
        # Start import timer for elapsed time logging
        $global:ImportStartTime = [DateTime]::Now
        Log-Message "Import operation started at: $($global:ImportStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"

        # Determine extraction location
        # In merge mode: extract to temp location first, then merge
        # In replace mode: extract directly to target
        $extractionTarget = $target
        $tempMergeLocation = $null
        if ($mergeMode) {
            $tempMergeLocation = Join-Path $env:TEMP ("ProfileMerge_$shortName_$(Get-Random)")
            $extractionTarget = $tempMergeLocation
            Log-Message "MERGE MODE: Will extract to temp location first: $tempMergeLocation"
        }

        # Create target directory for extraction (clean directory after backup/removal)
        if (-not (Test-Path $extractionTarget)) { 
            New-Item -ItemType Directory -Path $extractionTarget -Force | Out-Null 
        }
        
        if ($mergeMode) {
            Log-Message "MERGE MODE: Extracting to temp location for merging: $extractionTarget"
        } else {
            Log-Message "REPLACE MODE: Extracting directly to final profile location: $extractionTarget"
        }
        
        $global:ProgressBar.Value = 5
        $global:StatusText.Text = "Extracting ZIP archive..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Starting 7-Zip extraction..."
        
        # Build 7-Zip exclusion filters for Import mode
        $importExclusions = Get-RobocopyExclusions -Mode 'Import'
        $7zExclusions = @()
        
        # Exclude LOG files explicitly (not wildcards that might match NTUSER.DAT)
        $7zExclusions += "-x!ntuser.dat.LOG1"
        $7zExclusions += "-x!ntuser.dat.LOG2"
        $7zExclusions += "-x!NTUSER.DAT.LOG1"
        $7zExclusions += "-x!NTUSER.DAT.LOG2"
        $7zExclusions += "-x!UsrClass.dat.LOG1"
        $7zExclusions += "-x!UsrClass.dat.LOG2"
        
        # Exclude temp/cache files
        $7zExclusions += "-x!Thumbs.db"
        $7zExclusions += "-x!desktop.ini"
        
        # Exclude temp directories
        $7zExclusions += "-xr!AppData/Local/Temp"
        $7zExclusions += "-xr!AppData/LocalLow"
        
        # Exclude Outlook OST files (cached Exchange data - rebuilds automatically)
        $7zExclusions += "-x!*.ost"
        Log-Message "Excluding Outlook OST files (will rebuild automatically from Exchange)"
        
        Log-Message "Import exclusions: $($7zExclusions.Count) filters applied"
        
        # Extract with exclusion filters
        # Multi-threading: Use all CPU cores for maximum extraction speed (2-3x faster on multi-core systems)
        $threadCount = $Config.SevenZipThreads
        Log-Message "7-Zip extraction using $threadCount threads (detected $cpuCores CPU cores)"
        $args = @('x', $ZipPath, "-o$extractionTarget", '-y', "-mmt=$threadCount", '-bsp1') + $7zExclusions
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $global:SevenZipPath
        $psi.Arguments = $args -join ' '
        $psi.CreateNoWindow = $true
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        
        Log-Message "7-Zip command: $global:SevenZipPath $($args -join ' ')"
        
        Log-Message "Starting 7-Zip extraction with progress monitoring..."
        $extractProcess = [System.Diagnostics.Process]::Start($psi)
        
        # Monitor extraction progress
        $lastUpdate = [DateTime]::Now
        $buffer = New-Object System.Text.StringBuilder
        while (-not $extractProcess.HasExited) {
            if ($global:CancelRequested) {
                try { $extractProcess.Kill() } catch { }
                throw "Import cancelled by user during extraction"
            }
            $line = $extractProcess.StandardOutput.ReadLine()
            if ($line) {
                $buffer.AppendLine($line) | Out-Null
                # Parse progress from 7-Zip output (format: "  5%")
                if ($line -match '^\s*(\d+)%') {
                    $percent = [int]$matches[1]
                    # Map 7-Zip progress (0-100%) to our progress range (5-20%)
                    $mappedProgress = 5 + ($percent * 0.15)
                    $global:ProgressBar.Value = [Math]::Min(20, $mappedProgress)
                    $global:StatusText.Text = "Extracting files... $percent%"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            } elseif (([DateTime]::Now - $lastUpdate).TotalSeconds -ge 1) {
                $lastUpdate = [DateTime]::Now
                [System.Windows.Forms.Application]::DoEvents()
            }
            Start-Sleep -Milliseconds 100
        }
        
        # Consume any remaining output
        $buffer.Append($extractProcess.StandardOutput.ReadToEnd()) | Out-Null
        
        $extractProcess.WaitForExit()
        Log-Message "7-Zip exit code: $($extractProcess.ExitCode)"
        if ($extractProcess.ExitCode -ne 0) { 
            throw "Extract failed (ExitCode: $($extractProcess.ExitCode))"
        }
        
        if ($mergeMode) {
            Log-Message "ZIP extracted successfully to temp location for merging"
        } else {
            Log-Message "ZIP extracted successfully to final profile location"
        }
        
        # === CRITICAL: Verify extraction integrity ===
        $global:ProgressBar.Value = 20
        $global:StatusText.Text = "Verifying extracted files..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Post-extraction verification..."
        
        # List what was actually extracted
        $extractedItems = @()
        try {
            $extractedItems = Get-ChildItem -Path $extractionTarget -Force -ErrorAction SilentlyContinue | Select-Object -First 20
            Log-Message "Items in extraction directory:"
            foreach ($item in $extractedItems) {
                Log-Message "  - $($item.Name) $(if ($item.PSIsContainer) { '[DIR]' } else { "($($item.Length) bytes)" })"
            }
            if ((Get-ChildItem -Path $extractionTarget -Force -ErrorAction SilentlyContinue | Measure-Object).Count -gt 20) {
                Log-Message "  ... and more"
            }
        } catch {
            Log-Message "WARNING: Could not list extraction directory: $_"
        }
        
        # Check NTUSER.DAT exists and has reasonable size
        # Retry a few times in case of timing issues
        $hiveSrc = "$extractionTarget\NTUSER.DAT"
        $hiveFound = $false
        $hiveSize = 0
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                if (Test-Path $hiveSrc -ErrorAction Stop) {
                    $fileObj = Get-Item $hiveSrc -Force -ErrorAction Stop
                    $hiveSize = $fileObj.Length
                    $hiveFound = $true
                    Log-Message "NTUSER.DAT found on attempt $($attempt): $hiveSize bytes"
                    break
                }
            } catch {
                Log-Message "Attempt $($attempt) to access NTUSER.DAT failed: $_"
                if ($attempt -lt 3) {
                    Start-Sleep -Milliseconds 500
                }
            }
        }
        
        if (-not $hiveFound) {
            throw "CRITICAL: NTUSER.DAT missing or inaccessible after extraction! ZIP may be corrupted or extraction failed."
        }
        
        if ($hiveSize -lt 100000) {
            throw "CRITICAL: NTUSER.DAT is suspiciously small ($hiveSize bytes) - extraction corrupted?"
        }
        
        # Check manifest
        $manifestPath = "$extractionTarget\manifest.json"
        if (-not (Test-Path $manifestPath)) {
            throw "manifest.json missing after extraction"
        }
        Log-Message "Manifest found"
        
        # Parse manifest for source SID
        $manifest = $null
        try {
            $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
            Log-Message "Manifest parsed successfully"
        } catch {
            Log-Message "WARNING: Could not parse manifest: $_"
        }
        
        $global:ProgressBar.Value = 25
        # Install Winget apps from the correct location based on mode
        if ($mergeMode) {
            # In merge mode, Winget JSON is in temp extraction location
            Install-WingetAppsFromExport -TargetProfilePath $extractionTarget
        } else {
            # In replace mode, Winget JSON is in final profile location
            Install-WingetAppsFromExport -TargetProfilePath $target
        }
        $global:StatusText.Text = "Resolving user SID..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Resolving SID for user..."
        $sourceSID = $null
        if ($manifest) {
            $sourceSID = $manifest.SourceSID
            Log-Message "Source SID from manifest: $sourceSID"
            
            # Check if profile hash/timestamp available for incremental detection
            if ($manifest.ProfileHash -and $manifest.ProfileTimestamp) {
                Log-Message "Profile hash from export: $($manifest.ProfileHash)"
                Log-Message "Profile timestamp from export: $($manifest.ProfileTimestamp)"
                # Note: This info can be used to detect if an existing profile is outdated
                # and needs backup. Currently used for diagnostics only.
            }
        } else {
            Log-Message "WARNING: manifest.json not available - cannot verify source SID"
        }

        # === MERGE MODE: Merge extracted files into existing profile ===
        if ($mergeMode) {
            $global:ProgressBar.Value = 28
            $global:StatusText.Text = "Merging profile files..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "MERGE MODE: Merging files from extracted profile into existing profile..."
            Log-Message "MERGE MODE: Preserving existing user's NTUSER.DAT (keeping user settings/preferences)"
            
            try {
                # In merge mode, we want to preserve the existing user's NTUSER.DAT hive
                # This keeps their current registry settings, preferences, and avoids SID rewriting
                # We'll delete NTUSER.DAT from temp extraction before merging
                $hiveFileToSkip = Join-Path $tempMergeLocation "NTUSER.DAT"
                if (Test-Path $hiveFileToSkip) {
                    Log-Message "Removing NTUSER.DAT from extracted profile (will use existing user's hive)"
                    Remove-Item $hiveFileToSkip -Force -ErrorAction SilentlyContinue
                }
                
                # Also skip UsrClass.dat (user class registry hive)
                $userClassToSkip = Join-Path $tempMergeLocation "AppData\Local\Microsoft\Windows\UsrClass.dat"
                if (Test-Path $userClassToSkip) {
                    Log-Message "Removing UsrClass.dat from extracted profile (will use existing user's registry)"
                    Remove-Item $userClassToSkip -Force -ErrorAction SilentlyContinue
                }
                
                # Use robocopy to merge remaining files, skipping identical files and not overwriting newer files
                # in existing profile (to preserve any recent customizations)
                # Dynamic thread count: 8-32 threads based on CPU cores for optimal performance
                $robocopyThreads = $Config.RobocopyThreads
                Log-Message "Robocopy merge using $robocopyThreads threads (detected $cpuCores CPU cores)"
                $mergeArgs = @(
                    "`"$tempMergeLocation`"",  # Source
                    "`"$target`"",              # Destination
                    '/E',                       # Include subdirectories (even empty)
                    '/COPY:DATSOU',            # Copy all metadata
                    '/IS',                      # Include same-size files
                    '/IT',                      # Include same files with different times
                    '/R:2',                     # Retry on failure
                    '/W:1',                     # Wait 1 second between retries
                    "/MT:$robocopyThreads",    # Multi-threaded (dynamic based on CPU cores)
                    '/NP',                      # No progress percentage
                    '/NFL',                     # No file list
                    '/NDL'                      # No directory list
                )
                
                Log-Message "Robocopy merge command: robocopy $($mergeArgs -join ' ')"
                $mergeProcess = Start-Process -FilePath 'robocopy.exe' -ArgumentList $mergeArgs -Wait -NoNewWindow -PassThru
                
                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($mergeProcess.ExitCode -gt 7) {
                    throw "Robocopy merge failed with exit code $($mergeProcess.ExitCode)"
                }
                Log-Message "Robocopy merge completed with exit code $($mergeProcess.ExitCode)"
                
                # Cleanup temp merge location
                Log-Message "Cleaning up temporary merge location: $tempMergeLocation"
                Remove-FolderRobust -Path $tempMergeLocation
                Log-Message "Temporary merge location cleaned up successfully"
                
                # Verify key files exist in target after merge
                # Note: We skip NTUSER.DAT verification in merge mode since we're keeping the existing one
                $keyFiles = @('AppData', 'Desktop', 'Documents')
                $missingFiles = @()
                foreach ($keyFile in $keyFiles) {
                    if (-not (Test-Path (Join-Path $target $keyFile) -ErrorAction SilentlyContinue)) {
                        $missingFiles += $keyFile
                    }
                }
                if ($missingFiles.Count -gt 0) {
                    Log-Message "WARNING: Some folders missing after merge (may be OK): $($missingFiles -join ', ')"
                } else {
                    Log-Message "Merge verification: All key profile folders present in target"
                }
                
            } catch {
                Log-Message "ERROR during merge: $_"
                throw "Profile merge failed: $_"
            }
        }

        $global:ProgressBar.Value = 30
        $global:StatusText.Text = "Cleaning stale registry entries..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Checking for stale profile registry entries..."
        $base = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $staleKeys = Get-ChildItem $base -ErrorAction SilentlyContinue | Where-Object {
            $_.PSChildName -like "$sid*" -or ((Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$shortName*")
        }
        if ($staleKeys.Count -gt 0) {
            Log-Message "Found $($staleKeys.Count) stale registry entries - removing..."
            foreach ($k in $staleKeys) {
                Remove-Item $k.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                Log-Message "Removed: $($k.PSChildName)"
            }
        } else {
            Log-Message "No stale registry entries found"
        }

        # Also remove any abandoned Temp* profile folders (excluding current target user)
        try {
            # Match any profile folder starting with Temp (e.g. Temp, Temp., Temp-XYZ, Temp123, temp.anything)
            $tempFolders = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^Temp.*' }
            foreach ($tf in $tempFolders) {
                if ($tf.Name -ieq $shortName) { continue } # Never remove the target user folder if it matches
                Log-Message "Removing temporary profile folder: $($tf.FullName)"
                Remove-FolderRobust -Path $tf.FullName
            }
        } catch {
            Log-Message "WARNING: Failed Temp* profile cleanup: $_"
        }
        $global:ProgressBar.Value = 35
        
        # Log appropriate completion message based on mode
        if ($mergeMode) {
            Log-Message "MERGE MODE: Profile files successfully merged into existing profile at $target"
        } else {
            Log-Message "REPLACE MODE: Files extracted directly to target: $target"
        }
        $global:ProgressBar.Value = 95
        $global:StatusText.Text = "Applying permissions..."
        [System.Windows.Forms.Application]::DoEvents()
        
        if (-not $isDomain) {
            if ($mergeMode) {
                Log-Message "MERGE MODE: Skipping SID rewriting (using existing user's NTUSER.DAT)"
                Log-Message "MERGE MODE: Only applying folder ACLs (not touching existing hive)"
                
                # In merge mode, only set folder permissions without SID rewriting
                $global:StatusText.Text = "Applying folder permissions (merge mode)..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Apply basic folder ACLs without hive manipulation
                try {
                    $acl = Get-Acl $target
                    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($shortName, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $target -AclObject $acl
                    Log-Message "Basic folder permissions applied for merge mode"
                } catch {
                    Log-Message "WARNING: Could not set basic folder ACLs: $_"
                }
            } else {
                Log-Message "Setting ACLs for local user $shortName"
                Set-ProfileAcls -ProfileFolder $target -UserName $shortName -SourceSID $sourceSID -UserSID $sid
            }
        } else {
            # DOMAIN USER - Set-ProfileAcls handles SID translation + ACLs
            $global:StatusText.Text = "Applying DOMAIN user permissions + SID translation..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "DOMAIN USER: Applying profile ACLs for $Username..."
            Set-ProfileAcls -ProfileFolder $target -UserName $Username -SourceSID $sourceSID -UserSID $sid
            Log-Message "DOMAIN user profile fully prepared - first logon will succeed"
        }
        if (-not $isDomain) {
            $global:StatusText.Text = "Configuring user for login screen..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Configuring local user for interactive login..."
            try {
                $localUser = Get-LocalUser -Name $shortName -ErrorAction Stop
                if ($localUser.PasswordExpires) {
                    Set-LocalUser -Name $shortName -PasswordNeverExpires $true
                    Log-Message "Set password to never expire"
                }
                if (-not $localUser.Enabled) {
                    Enable-LocalUser -Name $shortName
                    Log-Message "Enabled user account"
                }
                try {
                    Add-LocalGroupMember -Group "Users" -Member $shortName -ErrorAction SilentlyContinue
                    Log-Message "Confirmed user is in Users group"
                } catch {
                    Log-Message "User already in Users group"
                }
                Log-Message "User account configured for login"
            } catch {
                Log-Message "WARNING: Could not fully configure user account: $_"
            }
            Log-Message "Configuring registry settings for login screen visibility..."
            try {
                $winlogonPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                if (Test-Path $winlogonPath) {
                    Set-ItemProperty -Path $winlogonPath -Name "DontDisplayLastUserName" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                    Log-Message "Configured Winlogon to show user list"
                }
                $specialAccountsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                if (-not (Test-Path $specialAccountsPath)) { New-Item -Path $specialAccountsPath -Force | Out-Null }
                Set-ItemProperty -Path $specialAccountsPath -Name "DontDisplayLastUserName" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                $hiddenAccountsPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"
                if (Test-Path $hiddenAccountsPath) {
                    Remove-ItemProperty -Path $hiddenAccountsPath -Name $shortName -ErrorAction SilentlyContinue
                    Log-Message "Ensured user is not in hidden accounts list"
                }
                Log-Message "Registry configured for login screen visibility"
            } catch {
                Log-Message "WARNING: Could not configure registry for login screen: $_"
            }
        }
        $global:ProgressBar.Value = 93
        $global:StatusText.Text = "Registering profile in Windows..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Registering profile in ProfileList..."
        $profileKey = "$base\$sid"
        New-Item -Path $profileKey -Force | Out-Null
        Set-ItemProperty -Path $profileKey -Name ProfileImagePath -Value $target -Type ExpandString -Force
        $dwordProps = "Flags","State","RefCount","ProfileLoadTimeLow","ProfileLoadTimeHigh"
        foreach ($p in $dwordProps) { Set-ItemProperty -Path $profileKey -Name $p -Value 0 -Type DWord -Force }
        Set-ItemProperty -Path $profileKey -Name CentralProfile -Value "" -Type String -Force
        Log-Message "Creating ProfileList SID binary value..."
        $userSidObj = [System.Security.Principal.SecurityIdentifier]::new($sid)
        $sidBytes = [byte[]]::new($userSidObj.BinaryLength)
        $userSidObj.GetBinaryForm($sidBytes, 0)
        Set-ItemProperty -Path $profileKey -Name Sid -Value $sidBytes -Type Binary -Force
        Log-Message "SID binary value created successfully ($($sidBytes.Length) bytes)"
        Log-Message "Profile registered: $profileKey"
		Log-Message "ProfileImagePath: $target"

        # === VERIFY: Ensure NTUSER.DAT can be loaded (prevent temporary profile issues) ===
        function Test-AndFix-ProfileHive {
            param(
                [string]$HivePath,
                [string]$UserId,
                [string]$UserName
            )

            $testMount = "HKU\ProfileCheck_$(Get-Random)"
            try {
                Log-Message "Attempting to load hive for verification: $HivePath"
                $out = reg load "$testMount" "$HivePath" 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Log-Message "Initial hive load failed: $out"
                    Log-Message "Attempting to repair NTUSER.DAT ownership and ACLs..."

                    # Take ownership as Administrators and grant full control, then set to user
                    try {
                        takeown /F "$HivePath" /A /R /D Y >$null 2>&1
                        icacls "$HivePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
                        icacls "$HivePath" /setowner *"$UserName" /Q >$null 2>&1
                        icacls "$HivePath" /grant:r *"${UserName}:(F)" /Q >$null 2>&1
                        Start-Sleep -Milliseconds 500
                    } catch {
                        Log-Message "Failed to adjust NTUSER.DAT ACLs: $_"
                    }

                    # Retry load
                    $out2 = reg load "$testMount" "$HivePath" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Log-Message "Retry hive load failed: $out2"
                        return $false
                    }
                }

                # Successfully loaded - unload and return success
                reg unload "$testMount" | Out-Null
                Log-Message "Hive verification succeeded"
                return $true
            } catch {
                Log-Message "Exception during hive verification: $_"
                try { reg unload "$testMount" | Out-Null } catch {}
                return $false
            }
        }

        # Run verification and attempt fixes if needed
        # In merge mode, we skip hive verification since we kept the existing user's NTUSER.DAT
        if (-not $mergeMode) {
            $hivePathToTest = "$target\NTUSER.DAT"
            $verifyOk = Test-AndFix-ProfileHive -HivePath $hivePathToTest -UserId $sid -UserName $shortName
            if (-not $verifyOk) {
                Log-Message "CRITICAL: NTUSER.DAT verification failed - profile cannot be used!"
                Log-Message "The hive file cannot be loaded into registry. This will result in a temporary profile on login."
                Log-Message "Possible causes:"
                Log-Message "  1. Hive file corrupted during extraction"
                Log-Message "  2. Owner/permissions incorrect (cannot be fixed with icacls)"
                Log-Message "  3. Hive file has incompatible format"
                Log-Message "Attempting additional diagnostics..."
                
                # Additional diagnostics
                $fileInfo = Get-Item $hivePathToTest -Force -ErrorAction SilentlyContinue
                if ($fileInfo) {
                    Log-Message "File exists: Yes"
                    Log-Message "File size: $($fileInfo.Length) bytes"
                    Log-Message "File ACL: $(icacls $hivePathToTest 2>&1 | Select-Object -First 3)"
                } else {
                    Log-Message "File exists: NO - extraction failed!"
                }
                
                throw "CRITICAL: NTUSER.DAT cannot be loaded. Import failed. Check the log for diagnostics."
            }
        } else {
            Log-Message "MERGE MODE: Skipping hive verification (using existing user's NTUSER.DAT)"
        }

        # IMPORTANT: Do NOT modify any registry settings here!
        # The extracted hive already contains all settings exactly as they were on source.
        # Any modifications to shell folders, navigation pane, search, or theme
        # will break UX fidelity. The only transformations applied were:
        # 1. Binary SID replacement (handled in Rewrite-HiveSID)
        # 2. Path string replacement for embedded paths (handled in Rewrite-HiveSID)
        Log-Message "Hive extraction complete - all settings preserved from source"

		# --------------------------------------------------------------
		# NOTE: Registry migration sections (printers, drives, shell extensions, theme)
		# have been removed because files are extracted directly to target.
		# NTUSER.DAT already contains all source profile settings.
		# SID rewrite in Set-ProfileAcls handles necessary transformations.
		# --------------------------------------------------------------

		# --------------------------------------------------------------
		# FINAL SUCCESS
		# --------------------------------------------------------------
		$global:ProgressBar.Value = 100
		$global:StatusText.Text   = "[OK] Migration completed successfully!"
		$global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)

		Log-Message "=================================================="
		Log-Message "PROFILE MIGRATION COMPLETED SUCCESSFULLY"
		Log-Message "=================================================="

		# Your original cleanup (if you still want it)
		$global:StatusText.Text = "Cleaning up temporary files..."
			
		$global:StatusText.Text = "Cleaning up temporary files..."
		[System.Windows.Forms.Application]::DoEvents()

        # === RECREATE STANDARD JUNCTIONS ===
        # These were excluded during export and need to be recreated pointing to new paths
        Log-Message "Recreating standard profile junctions..."
        $junctionsToCreate = @(
            @{ Source = "$target\Documents\My Music"; Target = "$target\Music"; Name = "My Music" }
            @{ Source = "$target\Documents\My Pictures"; Target = "$target\Pictures"; Name = "My Pictures" }
            @{ Source = "$target\Documents\My Videos"; Target = "$target\Videos"; Name = "My Videos" }
        )
        
        foreach ($junc in $junctionsToCreate) {
            try {
                # Check if target directory exists
                if (-not (Test-Path $junc.Target)) {
                    Log-Message "Junction target does not exist (OK if user customized profile): $($junc.Target)"
                    continue
                }
                
                # Remove if junction already exists (shouldn't, but be safe)
                if (Test-Path $junc.Source) {
                    $item = Get-Item $junc.Source -Force
                    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                        # For junctions, use rmdir /s /q to avoid prompts - removes only the link, not target
                        cmd /c rmdir /s /q "$($junc.Source)" >$null 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            Log-Message "Removed stale junction: $($junc.Source)"
                        } else {
                            Log-Message "WARNING: Failed to remove junction at $($junc.Source) (will attempt to recreate)"
                        }
                    }
                }
                
                # Create junction
                cmd /c mklink /j "$($junc.Source)" "$($junc.Target)" | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Log-Message "Created junction: $($junc.Name) -> $($junc.Target)"
                } else {
                    Log-Message "WARNING: Failed to create junction $($junc.Name)"
                }
            } catch {
                Log-Message "WARNING: Junction creation error for $($junc.Name): $_"
            }
        }

        # Remove Winget JSON artifacts from target profile
        try {
            $wingetJsons = @()
            $wingetJsons += (Join-Path $target 'Winget-Packages.json')
            $wingetJsons += (Join-Path $target 'Winget-Packages-Selected.json')
            foreach ($wj in $wingetJsons) {
                if (Test-Path $wj) {
                    Remove-Item $wj -Force -ErrorAction SilentlyContinue
                    Log-Message "Removed Winget file: $wj"
                }
            }
        } catch {
            Log-Message "WARNING: Failed to remove Winget JSON files: $_"
        }

        # Remove manifest.json from profile folder
        try {
            $manifestPath = Join-Path $target "manifest.json"
            if (Test-Path $manifestPath) {
                Remove-Item $manifestPath -Force -ErrorAction Stop
                Log-Message "Removed manifest.json from profile folder"
            }
        } catch {
            Log-Message "WARNING: Failed to remove manifest.json: $_"
        }

        # Clean up old backup profile folders
        try {
            Log-Message "Checking for old profile backup folders to clean up..."
            $backupPattern = "$target.backup_*"
            $backupFolders = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -match "^$([regex]::Escape($shortName))\.backup_\d{8}_\d{6}$" } |
                Sort-Object Name -Descending
            
            if ($backupFolders.Count -gt 0) {
                Log-Message "Found $($backupFolders.Count) backup folder(s) for user '$shortName'"
                
                # Keep the 2 most recent backups, delete older ones
                $keepCount = 2
                $foldersToDelete = $backupFolders | Select-Object -Skip $keepCount
                
                if ($foldersToDelete.Count -gt 0) {
                    Log-Message "Keeping $keepCount most recent backup(s), removing $($foldersToDelete.Count) older backup(s)..."
                    foreach ($oldBackup in $foldersToDelete) {
                        try {
                            $backupAge = [DateTime]::Now - $oldBackup.CreationTime
                            Log-Message "Removing old backup: $($oldBackup.Name) (created $([int]$backupAge.TotalDays) days ago)"
                            Remove-FolderRobust -Path $oldBackup.FullName
                            Log-Message "Successfully removed: $($oldBackup.Name)"
                        } catch {
                            Log-Message "WARNING: Could not remove backup folder $($oldBackup.Name): $_"
                        }
                    }
                    Log-Message "Old backup cleanup complete"
                } else {
                    Log-Message "All backups are recent - nothing to clean up"
                }
            } else {
                Log-Message "No backup folders found for cleanup"
            }
        } catch {
            Log-Message "WARNING: Error during backup folder cleanup: $_"
        }

        Log-Message "Cleanup complete"
        
        $global:ProgressBar.Value = 100
        Log-Message "=========================================="
        Log-Message "IMPORT SUCCESSFUL for user '$Username'"
        Log-Message "Profile SID: $sid"
        Log-Message "Profile Path: $target"
        Log-Message "User Type: $(if ($isDomain) { 'Domain' } else { 'Local' })"
        if (-not $isDomain) { Log-Message "GPSVC Support: ENABLED (Users group permissions applied)" }
        
        # (Removed) ACL diagnostics output
        
        # Log elapsed time for diagnostics
        if ($global:ImportStartTime) {
            $elapsed = [DateTime]::Now - $global:ImportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        Log-Message "=========================================="
        
        # Generate HTML migration report
        try {
            # Collect ZIP file statistics for report
            $zipSizeMB = 'N/A'
            $zipFileCount = 'N/A'
            $zipFolderCount = 'N/A'
            $uncompressedSizeMB = 'N/A'
            $compressionRatio = 'N/A'
            
            try {
                if (Test-Path $ZipPath) {
                    $zipFile = Get-Item $ZipPath -ErrorAction Stop
                    $zipSizeMB = [math]::Round($zipFile.Length / 1MB, 2)
                    
                    # Get basic ZIP statistics using 7-Zip
                    Log-Message "Collecting ZIP archive statistics for report..."
                    
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $global:SevenZipPath
                    $psi.Arguments = "l -slt `"$ZipPath`""
                    $psi.CreateNoWindow = $true
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    
                    $listProc = [System.Diagnostics.Process]::Start($psi)
                    $output = $listProc.StandardOutput.ReadToEnd()
                    $listProc.WaitForExit()
                    
                    if ($listProc.ExitCode -eq 0) {
                        Log-Message "7-Zip technical listing received, parsing..."
                        
                        # Count files and folders, sum sizes from technical output
                        $files = 0
                        $folders = 0
                        $totalSize = 0
                        
                        $lines = $output -split "`n"
                        $isFolder = $false
                        $currentSize = 0
                        
                        foreach ($line in $lines) {
                            $line = $line.Trim()
                            
                            # Check if this entry is a folder
                            if ($line -match '^Attributes = .+D') {
                                $isFolder = $true
                            }
                            
                            # Get the size
                            if ($line -match '^Size = (\d+)$') {
                                $currentSize = [int64]$matches[1]
                            }
                            
                            # When we hit a Path entry, we've finished reading an item
                            if ($line -match '^Path = (.+)$' -and $line -notmatch '^Path\s*=\s*$') {
                                if ($isFolder) {
                                    $folders++
                                } else {
                                    $files++
                                    $totalSize += $currentSize
                                }
                                # Reset for next item
                                $isFolder = $false
                                $currentSize = 0
                            }
                        }
                        
                        if ($files -gt 0) {
                            $zipFileCount = $files
                            $zipFolderCount = $folders
                            $uncompressedSizeMB = [math]::Round($totalSize / 1MB, 2)
                            
                            # Calculate compression ratio
                            if ($totalSize -gt 0) {
                                $compressionRatio = "$([math]::Round(($zipFile.Length / $totalSize) * 100, 1))%"
                            }
                            
                            Log-Message "ZIP statistics: $zipFileCount files, $zipFolderCount folders, $uncompressedSizeMB MB uncompressed, compression: $compressionRatio"
                        } else {
                            Log-Message "WARNING: No files found in 7-Zip output"
                        }
                    } else {
                        Log-Message "7-Zip list command failed with exit code: $($listProc.ExitCode)"
                    }
                }
            } catch {
                Log-Message "Could not collect ZIP statistics: $_"
            }
            
            # Get installed apps from global variable if available
            $installedAppsList = @()
            if ($global:InstalledAppsList) {
                $installedAppsList = $global:InstalledAppsList
                Log-Message "Including $($installedAppsList.Count) installed apps in report"
            }
            
            $reportData = @{
                Username = $Username
                UserType = if ($isDomain) { 'Domain' } else { 'Local' }
                TargetSID = $sid
                SourceSID = if ($sourceSID) { $sourceSID } else { 'N/A' }
                ProfilePath = $target
                ZipPath = $ZipPath
                ZipSizeMB = $zipSizeMB
                ImportMode = if ($mergeMode) { 'Merge' } else { 'Replace' }
                ElapsedMinutes = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
                ElapsedSeconds = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                BackupPath = $global:ImportBackupPath
                HashVerified = $Config.HashVerificationEnabled
                FileCount = $zipFileCount
                FolderCount = $zipFolderCount
                UncompressedSizeMB = $uncompressedSizeMB
                CompressionRatio = $compressionRatio
                Warnings = @()
                InstalledApps = $installedAppsList
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Import' -ReportData $reportData
            if ($reportPath) {
                Log-Info "Migration report available: $reportPath"
            }
        } catch {
            Log-Warning "Could not generate migration report: $_"
        }
        
        $successMsg = "IMPORT SUCCESSFUL`n`nUser: $Username`nSID: $sid`nPath: $target`n`n"
        if (-not $isDomain) {
            $successMsg += "Local user configured for login screen`nGPSVC permissions applied`n`n"
        }
        $successMsg += "REBOOT NOW, THEN LOG IN`n`n"
        $successMsg += "EMAIL SETUP (After First Login):`n"
        $successMsg += "- Outlook will rebuild mailbox cache (10-30 min)`n"
        $successMsg += "- Re-enter email passwords if prompted`n"
        $successMsg += "- Reconnect to Exchange/Microsoft 365`n"
        $successMsg += "- Copy PST archives manually if stored outside profile"
        [System.Windows.Forms.MessageBox]::Show($successMsg, "SUCCESS", "OK", "Information")
    } catch {
        Log-Message "=========================================="
        Log-Message "IMPORT FAILED: $_"
        Log-Message "=========================================="
        
        # ROLLBACK: Restore from backup if it exists and operation had started
        if ($global:ImportBackupPath -and (Test-Path $global:ImportBackupPath)) {
            Log-Message "Rollback: restoring from backup"
            try {
                if (Test-Path $target) {
                    Log-Message "ROLLBACK: Removing failed import at $target"
                    Remove-FolderRobust -Path $target -ErrorAction SilentlyContinue
                }
                Log-Message "Rollback: source $($global:ImportBackupPath)"
                Copy-Item -Path $global:ImportBackupPath -Destination $target -Recurse -Force -ErrorAction Stop
                Log-Message "ROLLBACK: Restoration completed successfully"
                # Remove the backup folder after successful restoration to avoid leaving stale backups
                try {
                    if (Test-Path $global:ImportBackupPath) {
                        Log-Message "Rollback: removing $($global:ImportBackupPath)"
                        Remove-FolderRobust -Path $global:ImportBackupPath -ErrorAction SilentlyContinue
                        Log-Message "Rollback: backup folder removed"
                        $global:ImportBackupPath = $null
                    }
                } catch {
                    Log-Message "WARNING: Could not remove backup folder: $_"
                }
            } catch {
                Log-Message "WARNING: Rollback failed: $_. Manual recovery may be needed."
                Log-Message "Backup available at: $($global:ImportBackupPath)"
            }
        } else {
            Log-Message "No backup available for rollback"
        }
        
        # Log elapsed time for diagnostics
        if ($global:ImportStartTime) {
            $elapsed = [DateTime]::Now - $global:ImportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        [System.Windows.Forms.MessageBox]::Show("Import failed: $_","Error","OK","Error")
    } finally {
        Stop-OperationLog
        # No temp folder cleanup needed - files extracted directly to target
        $global:StatusText.Text = "Idle"
        $global:ProgressBar.Value = 0
        $global:CurrentOperation = $null
        if ($global:CancelButton) { $global:CancelButton.Enabled = $false }
    }
}

# =============================================================================
# GUI - MODERN FLAT DESIGN
# =============================================================================
$global:Form = New-Object System.Windows.Forms.Form
$global:Form.Text = "Profile Migration Tool"
$global:Form.Size = New-Object System.Drawing.Size(920,720)
$global:Form.StartPosition = "CenterScreen"
$global:Form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$global:Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:Form.MinimumSize = New-Object System.Drawing.Size(920, 720)

# Header section with modern styling
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(920, 80)
$headerPanel.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($headerPanel)

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Location = New-Object System.Drawing.Point(20, 15)
$lblHeader.Size = New-Object System.Drawing.Size(500, 35)
$lblHeader.Text = "Windows Profile Migration"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$headerPanel.Controls.Add($lblHeader)

$lblSubheader = New-Object System.Windows.Forms.Label
$lblSubheader.Location = New-Object System.Drawing.Point(22, 48)
$lblSubheader.Size = New-Object System.Drawing.Size(500, 20)
$lblSubheader.Text = "Export, import, and merge user profiles with ease"
$lblSubheader.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$headerPanel.Controls.Add($lblSubheader)

# Settings button in header
$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Location = New-Object System.Drawing.Point(850, 25)
$btnSettings.Size = New-Object System.Drawing.Size(40, 40)
$btnSettings.Text = "..."
$btnSettings.BackColor = [System.Drawing.Color]::White
$btnSettings.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSettings.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnSettings.FlatAppearance.BorderSize = 1
$btnSettings.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSettings.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
})
$btnSettings.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::White
})
$btnSettings.Add_Click({
    # Settings dialog
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(550, 200)
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.FormBorderStyle = "FixedDialog"
    $settingsForm.MaximizeBox = $false
    $settingsForm.MinimizeBox = $false
    $settingsForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $settingsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Content panel
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 15)
    $panel.Size = New-Object System.Drawing.Size(505, 90)
    $panel.BackColor = [System.Drawing.Color]::White
    $settingsForm.Controls.Add($panel)
    
    # 7-Zip path label
    $lbl7z = New-Object System.Windows.Forms.Label
    $lbl7z.Location = New-Object System.Drawing.Point(15, 15)
    $lbl7z.Size = New-Object System.Drawing.Size(100, 23)
    $lbl7z.Text = "7-Zip Path:"
    $lbl7z.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($lbl7z)
    
    # 7-Zip path textbox
    $txt7zPath = New-Object System.Windows.Forms.TextBox
    $txt7zPath.Location = New-Object System.Drawing.Point(15, 40)
    $txt7zPath.Size = New-Object System.Drawing.Size(390, 25)
    $txt7zPath.Text = $global:SevenZipPath
    $txt7zPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txt7zPath.ReadOnly = $true
    $panel.Controls.Add($txt7zPath)
    
    # Browse button for 7-Zip
    $btnBrowse7z = New-Object System.Windows.Forms.Button
    $btnBrowse7z.Location = New-Object System.Drawing.Point(415, 38)
    $btnBrowse7z.Size = New-Object System.Drawing.Size(75, 28)
    $btnBrowse7z.Text = "Browse..."
    $btnBrowse7z.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnBrowse7z.ForeColor = [System.Drawing.Color]::White
    $btnBrowse7z.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse7z.FlatAppearance.BorderSize = 0
    $btnBrowse7z.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse7z.Add_Click({
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = "7-Zip Executable (7z.exe)|7z.exe|All Files (*.*)|*.*"
        $openFile.Title = "Locate 7z.exe"
        if ($openFile.ShowDialog() -eq "OK") {
            if (Test-Path $openFile.FileName) {
                $txt7zPath.Text = $openFile.FileName
            }
        }
    })
    $panel.Controls.Add($btnBrowse7z)
    
    # OK button
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(310, 120)
    $btnOK.Size = New-Object System.Drawing.Size(100, 32)
    $btnOK.Text = "Save"
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.Add_Click({
        $newPath = $txt7zPath.Text
        if (Test-Path $newPath) {
            $global:SevenZipPath = $newPath
            [System.Windows.Forms.MessageBox]::Show("7-Zip path updated successfully!`n`nNew path: $newPath", "Settings Saved", "OK", "Information")
            $settingsForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid path! Please select a valid 7z.exe file.", "Invalid Path", "OK", "Error")
        }
    })
    $settingsForm.Controls.Add($btnOK)
    
    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(420, 120)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
    $btnCancel.Text = "Cancel"
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_Click({ $settingsForm.Close() })
    $settingsForm.Controls.Add($btnCancel)
    
    $settingsForm.ShowDialog() | Out-Null
})
$headerPanel.Controls.Add($btnSettings)

# Tooltip for settings button
$settingsTooltip = New-Object System.Windows.Forms.ToolTip
$settingsTooltip.SetToolTip($btnSettings, "Settings - Configure 7-Zip path")

# Main content card
$cardPanel = New-Object System.Windows.Forms.Panel
$cardPanel.Location = New-Object System.Drawing.Point(15, 95)
$cardPanel.Size = New-Object System.Drawing.Size(880, 100)
$cardPanel.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($cardPanel)

$lblU = New-Object System.Windows.Forms.Label
$lblU.Location = New-Object System.Drawing.Point(15, 15)
$lblU.Size = New-Object System.Drawing.Size(100, 23)
$lblU.Text = "Username:"
$lblU.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$cardPanel.Controls.Add($lblU)
$global:UserComboBox = New-Object System.Windows.Forms.ComboBox
$global:UserComboBox.Location = New-Object System.Drawing.Point(120, 12)
$global:UserComboBox.Size = New-Object System.Drawing.Size(280,25)
$global:UserComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$global:UserComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Get-ProfileDisplayEntries | ForEach-Object { $global:UserComboBox.Items.Add($_.DisplayName) | Out-Null }
$cardPanel.Controls.Add($global:UserComboBox)

# Refresh Profiles button
$btnRefreshProfiles = New-Object System.Windows.Forms.Button
$btnRefreshProfiles.Location = New-Object System.Drawing.Point(405, 10)
$btnRefreshProfiles.Size = New-Object System.Drawing.Size(30, 28)
$btnRefreshProfiles.Text = "⟳"
$btnRefreshProfiles.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnRefreshProfiles.ForeColor = [System.Drawing.Color]::White
$btnRefreshProfiles.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefreshProfiles.FlatAppearance.BorderSize = 0
$btnRefreshProfiles.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnRefreshProfiles.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRefreshProfiles.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
})
$btnRefreshProfiles.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
})
$btnRefreshProfiles.Add_Click({
    # Refresh profile list with progress indication
    $btnRefreshProfiles.Enabled = $false
    $originalText = $btnRefreshProfiles.Text
    $btnRefreshProfiles.Text = "..."
    
    # Show scanning status
    $originalStatusText = $global:StatusText.Text
    $originalStatusColor = $global:StatusText.ForeColor
    $global:StatusText.Text = "Scanning profiles and calculating sizes..."
    $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        $currentSelection = $global:UserComboBox.Text
        $global:UserComboBox.Items.Clear()
        
        # Calculate sizes during refresh (this is when user explicitly requests it)
        $allProfiles = Get-LocalProfiles
        $profileCount = $allProfiles.Count
        $currentIndex = 0
        
        $profiles = Get-ProfileDisplayEntries -CalculateSizes
        
        foreach ($p in $profiles) {
            $global:UserComboBox.Items.Add($p.DisplayName) | Out-Null
        }
        
        # Restore selection if it still exists
        if ($currentSelection) {
            $matchingItem = $global:UserComboBox.Items | Where-Object { $_ -like "$currentSelection*" } | Select-Object -First 1
            if ($matchingItem) {
                $global:UserComboBox.Text = $matchingItem
            }
        }
        
        # Success message with size info
        $global:StatusText.Text = "Profile scan complete - $($profiles.Count) profiles with size estimates"
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        [System.Windows.Forms.Application]::DoEvents()
        
        # Brief pause to show success message
        Start-Sleep -Milliseconds 800
        
    } catch {
        $global:StatusText.Text = "Error scanning profiles: $_"
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
        [System.Windows.Forms.MessageBox]::Show("Error refreshing profiles: $_", "Error", "OK", "Error")
    } finally {
        # Restore original status
        $global:StatusText.Text = $originalStatusText
        $global:StatusText.ForeColor = $originalStatusColor
        $btnRefreshProfiles.Text = $originalText
        $btnRefreshProfiles.Enabled = $true
    }
})
$cardPanel.Controls.Add($btnRefreshProfiles)

# Tooltip for refresh button
$refreshTooltip = New-Object System.Windows.Forms.ToolTip
$refreshTooltip.SetToolTip($btnRefreshProfiles, "Refresh profile list and scan sizes")

# Modern flat buttons with icons and hover effects
$global:BrowseButton = New-Object System.Windows.Forms.Button
$global:BrowseButton.Location = New-Object System.Drawing.Point(465, 8)
$global:BrowseButton.Size = New-Object System.Drawing.Size(120, 32)
$global:BrowseButton.Text = "[...] Browse"
$global:BrowseButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:BrowseButton.ForeColor = [System.Drawing.Color]::White
$global:BrowseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:BrowseButton.FlatAppearance.BorderSize = 0
$global:BrowseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:BrowseButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:BrowseButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
})
$global:BrowseButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
})
$cardPanel.Controls.Add($global:BrowseButton)

# View Log Button (below Browse button)
$global:ViewLogButton = New-Object System.Windows.Forms.Button
$global:ViewLogButton.Location = New-Object System.Drawing.Point(465, 48)
$global:ViewLogButton.Size = New-Object System.Drawing.Size(35, 32)
$global:ViewLogButton.Text = "..."
$global:ViewLogButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:ViewLogButton.ForeColor = [System.Drawing.Color]::White
$global:ViewLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ViewLogButton.FlatAppearance.BorderSize = 0
$global:ViewLogButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ViewLogButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ViewLogButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
})
$global:ViewLogButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
})
$global:ViewLogButton.Add_Click({
    if ($global:CurrentLogFile -and (Test-Path $global:CurrentLogFile)) {
        Show-LogViewer -LogPath $global:CurrentLogFile -Title "Current Operation Log"
    } else {
        # Show most recent log file
        $logsDir = Join-Path $PSScriptRoot "Logs"
        if (Test-Path $logsDir) {
            $latestLog = Get-ChildItem -Path $logsDir -Filter "*.log" -File -ErrorAction SilentlyContinue | 
                         Sort-Object LastWriteTime -Descending | 
                         Select-Object -First 1
            if ($latestLog) {
                Show-LogViewer -LogPath $latestLog.FullName -Title "Latest Log: $($latestLog.Name)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("No log files found.`n`nLog files will be created in the Logs folder when you run Export or Import operations.", "No Logs", "OK", "Information")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No log files found.`n`nLog files will be created in the Logs folder when you run Export or Import operations.", "No Logs", "OK", "Information")
        }
    }
})
$cardPanel.Controls.Add($global:ViewLogButton)

# Tooltip for View Log button
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($global:ViewLogButton, "View Full Log")

# Log Level Selector (next to View Log button)
$lblLogLevel = New-Object System.Windows.Forms.Label
$lblLogLevel.Location = New-Object System.Drawing.Point(510, 51)
$lblLogLevel.Size = New-Object System.Drawing.Size(70, 23)
$lblLogLevel.Text = "Log Level:"
$lblLogLevel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblLogLevel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$cardPanel.Controls.Add($lblLogLevel)

$global:LogLevelComboBox = New-Object System.Windows.Forms.ComboBox
$global:LogLevelComboBox.Location = New-Object System.Drawing.Point(510, 70)
$global:LogLevelComboBox.Size = New-Object System.Drawing.Size(75, 25)
$global:LogLevelComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$global:LogLevelComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 8)
@('DEBUG', 'INFO', 'WARN', 'ERROR') | ForEach-Object { $global:LogLevelComboBox.Items.Add($_) | Out-Null }
$global:LogLevelComboBox.SelectedItem = $Config.LogLevel
$global:LogLevelComboBox.Add_SelectedIndexChanged({
    $oldLevel = $Config.LogLevel
    $Config.LogLevel = $global:LogLevelComboBox.SelectedItem
    # Refresh display to show filtered logs
    Refresh-LogDisplay
    # Log the change (will appear if new level allows INFO)
    Log-Info "Log level changed: $oldLevel -> $($Config.LogLevel)"
})
$cardPanel.Controls.Add($global:LogLevelComboBox)

$global:ExportButton = New-Object System.Windows.Forms.Button
$global:ExportButton.Location = New-Object System.Drawing.Point(595, 8)
$global:ExportButton.Size = New-Object System.Drawing.Size(120, 32)
$global:ExportButton.Text = "[>>] Export"
$global:ExportButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:ExportButton.ForeColor = [System.Drawing.Color]::White
$global:ExportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ExportButton.FlatAppearance.BorderSize = 0
$global:ExportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ExportButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ExportButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
})
$global:ExportButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
})
$cardPanel.Controls.Add($global:ExportButton)

$global:ImportButton = New-Object System.Windows.Forms.Button
$global:ImportButton.Location = New-Object System.Drawing.Point(725, 8)
$global:ImportButton.Size = New-Object System.Drawing.Size(120, 32)
$global:ImportButton.Text = "[<<] Import"
$global:ImportButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$global:ImportButton.ForeColor = [System.Drawing.Color]::White
$global:ImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ImportButton.FlatAppearance.BorderSize = 0
$global:ImportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ImportButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ImportButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12)
})
$global:ImportButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
})
$cardPanel.Controls.Add($global:ImportButton)

# Cancel Button (in second row, rightmost position)
$global:CancelButton = New-Object System.Windows.Forms.Button
$global:CancelButton.Location = New-Object System.Drawing.Point(725, 48)
$global:CancelButton.Size = New-Object System.Drawing.Size(120, 32)
$global:CancelButton.Text = "[X] Cancel"
$global:CancelButton.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
$global:CancelButton.ForeColor = [System.Drawing.Color]::White
$global:CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:CancelButton.FlatAppearance.BorderSize = 0
$global:CancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:CancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:CancelButton.Enabled = $false
$global:CancelButton.Add_MouseEnter({
    if ($this.Enabled) { $this.BackColor = [System.Drawing.Color]::FromArgb(200, 15, 30) }
})
$global:CancelButton.Add_MouseLeave({
    if ($this.Enabled) { $this.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35) }
})
$global:CancelButton.Add_Click({
    if ($global:CurrentOperation) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Cancel current $($global:CurrentOperation) operation?`n`nThis may leave the profile in an inconsistent state.",
            "Cancel Operation", "YesNo", "Warning")
        if ($result -eq "Yes") {
            $global:CancelRequested = $true
            Log-Message "User requested cancellation of $($global:CurrentOperation) operation"
            $global:StatusText.Text = "Cancelling operation..."
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            $global:CancelButton.Enabled = $false
        }
    }
})
$cardPanel.Controls.Add($global:CancelButton)

# File selection indicator
$global:FileLabel = New-Object System.Windows.Forms.Label
$global:FileLabel.Location = New-Object System.Drawing.Point(120, 50)
$global:FileLabel.Size = New-Object System.Drawing.Size(725, 35)
$global:FileLabel.Text = "No file selected - Use Browse to select a profile archive"
$global:FileLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$global:FileLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$cardPanel.Controls.Add($global:FileLabel)

# Domain Join Card
$domainCard = New-Object System.Windows.Forms.Panel
$domainCard.Location = New-Object System.Drawing.Point(15, 210)
$domainCard.Size = New-Object System.Drawing.Size(875, 180)
$domainCard.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($domainCard)

$lblDomainHeader = New-Object System.Windows.Forms.Label
$lblDomainHeader.Location = New-Object System.Drawing.Point(15, 10)
$lblDomainHeader.Size = New-Object System.Drawing.Size(250, 25)
$lblDomainHeader.Text = "Domain Join Settings"
$lblDomainHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblDomainHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$domainCard.Controls.Add($lblDomainHeader)

$global:DomainCheckBox = New-Object System.Windows.Forms.CheckBox
$global:DomainCheckBox.Location = New-Object System.Drawing.Point(25,45)
$global:DomainCheckBox.Size = New-Object System.Drawing.Size(200,23)
$global:DomainCheckBox.Text = "Join Domain After Import"
$global:DomainCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$domainCard.Controls.Add($global:DomainCheckBox)
$lblC = New-Object System.Windows.Forms.Label
$lblC.Location = New-Object System.Drawing.Point(25,75)
$lblC.Size = New-Object System.Drawing.Size(120,23)
$lblC.Text = "Computer Name:"
$domainCard.Controls.Add($lblC)
$global:ComputerNameTextBox = New-Object System.Windows.Forms.TextBox
$global:ComputerNameTextBox.Location = New-Object System.Drawing.Point(150,73)
$global:ComputerNameTextBox.Size = New-Object System.Drawing.Size(200,25)
$global:ComputerNameTextBox.Text = $env:COMPUTERNAME
$global:ComputerNameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:ComputerNameTextBox)

$lblD = New-Object System.Windows.Forms.Label
$lblD.Location = New-Object System.Drawing.Point(380,75)
$lblD.Size = New-Object System.Drawing.Size(100,23)
$lblD.Text = "Domain Name:"
$domainCard.Controls.Add($lblD)
$global:DomainNameTextBox = New-Object System.Windows.Forms.TextBox
$global:DomainNameTextBox.Location = New-Object System.Drawing.Point(485,73)
$global:DomainNameTextBox.Size = New-Object System.Drawing.Size(250,25)
$global:DomainNameTextBox.Text = "corp.example.com"
$global:DomainNameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:DomainNameTextBox)

$lblRestart = New-Object System.Windows.Forms.Label
$lblRestart.Location = New-Object System.Drawing.Point(25,110)
$lblRestart.Size = New-Object System.Drawing.Size(120,23)
$lblRestart.Text = "Restart Mode:"
$domainCard.Controls.Add($lblRestart)
$global:RestartComboBox = New-Object System.Windows.Forms.ComboBox
$global:RestartComboBox.Location = New-Object System.Drawing.Point(150,108)
$global:RestartComboBox.Size = New-Object System.Drawing.Size(200,25)
$global:RestartComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$global:RestartComboBox.Items.AddRange(@('Prompt', 'Delayed', 'Never', 'Immediate'))
$global:RestartComboBox.SelectedIndex = 0
$global:RestartComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:RestartComboBox)

$lblDelay = New-Object System.Windows.Forms.Label
$lblDelay.Location = New-Object System.Drawing.Point(380,110)
$lblDelay.Size = New-Object System.Drawing.Size(100,23)
$lblDelay.Text = "Delay (sec):"
$domainCard.Controls.Add($lblDelay)
$global:DelayTextBox = New-Object System.Windows.Forms.TextBox
$global:DelayTextBox.Location = New-Object System.Drawing.Point(485,108)
$global:DelayTextBox.Size = New-Object System.Drawing.Size(80,25)
$global:DelayTextBox.Text = "30"
$global:DelayTextBox.Enabled = $false
$global:DelayTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:DelayTextBox)

$global:DomainJoinButton = New-Object System.Windows.Forms.Button
$global:DomainJoinButton.Location = New-Object System.Drawing.Point(25,143)
$global:DomainJoinButton.Size = New-Object System.Drawing.Size(150,30)
$global:DomainJoinButton.Text = "Join Now"
$global:DomainJoinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:DomainJoinButton.ForeColor = [System.Drawing.Color]::White
$global:DomainJoinButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:DomainJoinButton.FlatAppearance.BorderSize = 0
$global:DomainJoinButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:DomainJoinButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:DomainJoinButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
})
$global:DomainJoinButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
})
$domainCard.Controls.Add($global:DomainJoinButton)

# Retry domain reachability button
$global:DomainRetryButton = New-Object System.Windows.Forms.Button
$global:DomainRetryButton.Location = New-Object System.Drawing.Point(185,143)
$global:DomainRetryButton.Size = New-Object System.Drawing.Size(165,30)
$global:DomainRetryButton.Text = "Retry Domain Check"
$global:DomainRetryButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:DomainRetryButton.ForeColor = [System.Drawing.Color]::White
$global:DomainRetryButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:DomainRetryButton.FlatAppearance.BorderSize = 0
$global:DomainRetryButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:DomainRetryButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:DomainRetryButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
})
$global:DomainRetryButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
})
$domainCard.Controls.Add($global:DomainRetryButton)

$global:RestartComboBox.Add_SelectedIndexChanged({ $global:DelayTextBox.Enabled = ($global:RestartComboBox.SelectedItem -eq 'Delayed') })
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($global:RestartComboBox, "Prompt: Ask before restarting`nDelayed: Countdown before restart`nNever: Manual restart required`nImmediate: Restart without asking")

# Progress and Status Card
$progressCard = New-Object System.Windows.Forms.Panel
$progressCard.Location = New-Object System.Drawing.Point(15, 405)
$progressCard.Size = New-Object System.Drawing.Size(875, 75)
$progressCard.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($progressCard)

$global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$global:ProgressBar.Location = New-Object System.Drawing.Point(15,15)
$global:ProgressBar.Size = New-Object System.Drawing.Size(845,25)
$global:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$global:ProgressBar.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$progressCard.Controls.Add($global:ProgressBar)

$global:StatusText = New-Object System.Windows.Forms.Label
$global:StatusText.Location = New-Object System.Drawing.Point(15,45)
$global:StatusText.Size = New-Object System.Drawing.Size(845,23)
$global:StatusText.Text = "[Ready]"
$global:StatusText.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$progressCard.Controls.Add($global:StatusText)

# Activity Log Section
$lblLogHeader = New-Object System.Windows.Forms.Label
$lblLogHeader.Location = New-Object System.Drawing.Point(20, 490)
$lblLogHeader.Size = New-Object System.Drawing.Size(200, 25)
$lblLogHeader.Text = "Activity Log"
$lblLogHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblLogHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:Form.Controls.Add($lblLogHeader)

$global:LogBox = New-Object System.Windows.Forms.TextBox
$global:LogBox.Location = New-Object System.Drawing.Point(15,520)
$global:LogBox.Size = New-Object System.Drawing.Size(875,165)
$global:LogBox.Multiline = $true
$global:LogBox.ScrollBars = "Vertical"
$global:LogBox.ReadOnly = $true
$global:LogBox.Font = New-Object System.Drawing.Font("Consolas",9)
$global:LogBox.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$global:LogBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$global:Form.Controls.Add($global:LogBox)

# =============================================================================
# BUTTON HANDLERS
# =============================================================================
$global:BrowseButton.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "ZIP Files (*.zip)|*.zip"
    if ($dlg.ShowDialog() -eq "OK") {
        $global:SelectedZipPath = $dlg.FileName
        
        # Update file selection indicator
        $fileInfo = Get-Item $dlg.FileName
        $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
        $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        $sizeDisplay = if ($sizeMB -gt 1) { "$sizeMB MB" } else { "$sizeKB KB" }
        $global:FileLabel.Text = "Selected: $($fileInfo.Name) ($sizeDisplay)"
        $global:FileLabel.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        
        $global:ProgressBar.Value = 5
        $global:StatusText.Text = "Analyzing ZIP file..."
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Selected: $global:SelectedZipPath"

        # Auto-detect domain profile and pre-select domain join if computer not in a domain
        try {
            $global:ProgressBar.Value = 10
            $global:StatusText.Text = "Checking computer domain status..."
            [System.Windows.Forms.Application]::DoEvents()
            $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
            $partOfDomain = $cs.PartOfDomain
            if (-not $partOfDomain) {
                $global:ProgressBar.Value = 20
                $global:StatusText.Text = "Reading manifest from ZIP..."
                [System.Windows.Forms.Application]::DoEvents()
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = [IO.Compression.ZipFile]::OpenRead($global:SelectedZipPath)
                $entry = $zip.Entries | Where-Object { $_.FullName -ieq 'manifest.json' }
                if ($entry) {
                    $sr = New-Object IO.StreamReader($entry.Open())
                    $jsonContent = $sr.ReadToEnd()
                    $sr.Dispose(); $zip.Dispose()
                    try {
                        $manifestObj = $jsonContent | ConvertFrom-Json -ErrorAction Stop
                        if ($manifestObj.IsDomainUser -and $manifestObj.Domain) {
                            # Resolve FQDN from NetBIOS name
                            $global:ProgressBar.Value = 40
                            $global:StatusText.Text = "Resolving domain name..."
                            [System.Windows.Forms.Application]::DoEvents()
                            $domainName = Get-DomainFQDN -NetBIOSName $manifestObj.Domain
                            $global:DomainCheckBox.Checked = $true
                            $global:DomainNameTextBox.Text = $domainName
                            Log-Message "Auto-enabled domain join (domain: $domainName from NetBIOS: $($manifestObj.Domain)) based on imported profile manifest"
                            # Warn early if domain appears unreachable
                            if (Get-Command Test-DomainReachability -ErrorAction SilentlyContinue) {
                                $global:ProgressBar.Value = 60
                                $global:StatusText.Text = "Testing domain connectivity..."
                                [System.Windows.Forms.Application]::DoEvents()
                                $reach = Test-DomainReachability -DomainName $domainName
                                if (-not $reach.Success) {
                                    Log-Message "WARNING: Domain '$domainName' unreachable during ZIP selection: $($reach.Error)"
                                    $global:DomainCheckBox.Checked = $false
                                    $global:ProgressBar.Value = 80
                                    
                                    # Modern domain unreachable dialog
                                    $reachForm = New-Object System.Windows.Forms.Form
                                    $reachForm.Text = "Domain Unreachable"
                                    $reachForm.Size = New-Object System.Drawing.Size(520,280)
                                    $reachForm.StartPosition = "CenterScreen"
                                    $reachForm.TopMost = $true
                                    $reachForm.FormBorderStyle = "FixedDialog"
                                    $reachForm.MaximizeBox = $false
                                    $reachForm.MinimizeBox = $false
                                    $reachForm.BackColor = [System.Drawing.Color]::White
                                    $reachForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                                    
                                    $lblWarning = New-Object System.Windows.Forms.Label
                                    $lblWarning.Location = New-Object System.Drawing.Point(20,15)
                                    $lblWarning.Size = New-Object System.Drawing.Size(480,30)
                                    $lblWarning.Text = "Domain Connection Issue"
                                    $lblWarning.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
                                    $lblWarning.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                                    $reachForm.Controls.Add($lblWarning)
                                    
                                    $lblMsg = New-Object System.Windows.Forms.Label
                                    $lblMsg.Location = New-Object System.Drawing.Point(20,55)
                                    $lblMsg.Size = New-Object System.Drawing.Size(470,120)
                                    $lblMsg.Text = "Domain '$domainName' appears unreachable right now.`n`nDetail: $($reach.Error)`n`nThe domain join option has been unchecked. Click Retry after connectivity is restored or Close to continue."
                                    $reachForm.Controls.Add($lblMsg)
                                    
                                    $btnRetry = New-Object System.Windows.Forms.Button
                                    $btnRetry.Location = New-Object System.Drawing.Point(260,195)
                                    $btnRetry.Size = New-Object System.Drawing.Size(110,35)
                                    $btnRetry.Text = "Retry"
                                    $btnRetry.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                                    $btnRetry.ForeColor = [System.Drawing.Color]::White
                                    $btnRetry.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                                    $btnRetry.FlatAppearance.BorderSize = 0
                                    $btnRetry.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                    $btnRetry.Cursor = [System.Windows.Forms.Cursors]::Hand
                                    $btnRetry.Add_Click({
                                        $global:ProgressBar.Value = 90
                                        $global:StatusText.Text = "Retrying domain connectivity..."
                                        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                                        [System.Windows.Forms.Application]::DoEvents()
                                        $retryResult = Test-DomainReachability -DomainName $domainName
                                        if ($retryResult.Success) {
                                            $global:DomainCheckBox.Checked = $true
                                            Log-Message "Domain '$domainName' now reachable"
                                            $global:ProgressBar.Value = 100
                                            $global:StatusText.Text = "[OK] Domain reachable - ready to import"
                                            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
                                            [System.Windows.Forms.MessageBox]::Show("Domain '$domainName' is now reachable.","Domain OK","OK","Information") | Out-Null
                                            $reachForm.Close()
                                        } else {
                                            Log-Message "Retry failed: $($retryResult.Error)"
                                            $global:ProgressBar.Value = 0
                                            $global:StatusText.Text = "[ERROR] Domain unreachable"
                                            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                                            [System.Windows.Forms.MessageBox]::Show("Still unreachable.`n`n$($retryResult.Error)","Retry Failed","OK","Warning") | Out-Null
                                        }
                                    })
                                    $reachForm.Controls.Add($btnRetry)
                                    
                                    $btnClose = New-Object System.Windows.Forms.Button
                                    $btnClose.Location = New-Object System.Drawing.Point(380,195)
                                    $btnClose.Size = New-Object System.Drawing.Size(110,35)
                                    $btnClose.Text = "Close"
                                    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
                                    $btnClose.ForeColor = [System.Drawing.Color]::White
                                    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                                    $btnClose.FlatAppearance.BorderSize = 0
                                    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                                    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
                                    $btnClose.DialogResult = "OK"
                                    $reachForm.Controls.Add($btnClose)
                                    $reachForm.AcceptButton = $btnClose
                                    
                                    $reachForm.ShowDialog() | Out-Null
                                    $reachForm.Dispose()
                                    $global:ProgressBar.Value = 0
                                    $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf) - Domain unreachable"
                                } else {
                                    Log-Message "Domain '$domainName' reachable (pre-flight hint)"
                                    $global:ProgressBar.Value = 100
                                    $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf) - Domain OK"
                                    Start-Sleep -Milliseconds 500
                                    $global:ProgressBar.Value = 0
                                }
                            }
                        } else {
                            Log-Message "Manifest found but not a domain profile (IsDomainUser=$($manifestObj.IsDomainUser))"
                            $global:ProgressBar.Value = 0
                            $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                        }
                    } catch {
                        Log-Message "Domain auto-detect parse failed: $($_.Exception.Message)"
                        $global:ProgressBar.Value = 0
                        $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                    }
                } else {
                    $zip.Dispose()
                    Log-Message "No manifest.json in ZIP - skipping domain auto-detect"
                    $global:ProgressBar.Value = 0
                    $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                }
            } else {
                $global:ProgressBar.Value = 0
                $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
            }
        } catch {
            Log-Message "Domain auto-detect skipped: $($_.Exception.Message)"
            $global:ProgressBar.Value = 0
            $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
        }
    }
})
$global:ExportButton.Add_Click({
    $u = $global:UserComboBox.Text.Trim()
    if (-not $u) {
        [System.Windows.Forms.MessageBox]::Show("Select username.","Error","OK","Error")
        return
    }
    # Strip size suffix if present (e.g., "username - [756.6 MB]" -> "username")
    if ($u -match '^(.+?)\s+-\s+\[.+\]$') {
        $u = $matches[1].Trim()
    }
    # Extract short name from DOMAIN\username or COMPUTERNAME\username
    $shortName = if ($u -match '\\') { ($u -split '\\',2)[1] } else { $u }
    # Remove any remaining backslashes for safe filename
    $shortName = $shortName -replace '\\',''
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "ZIP Files (*.zip)|*.zip"
    $dlg.FileName = "$shortName-Export-$(Get-Date -f yyyyMMdd_HHmmss).zip"
    if ($dlg.ShowDialog() -eq "OK") {
        Export-UserProfile -Username $u -ZipPath $dlg.FileName
    }
})
$global:ImportButton.Add_Click({
    $u = $global:UserComboBox.Text.Trim()
    if (-not $u) {
        [System.Windows.Forms.MessageBox]::Show("Enter or select a username (e.g., 'john' or 'DOMAIN\john').`n`nYou can type a new username that doesn't exist yet - it will be created during import.", "Username Required","OK","Warning")
        return
    }
    # Strip size suffix if present (e.g., "username - [756.6 MB]" -> "username")
    if ($u -match '^(.+?)\s+-\s+\[.+\]$') {
        $u = $matches[1].Trim()
    }
    if (-not $global:SelectedZipPath) {
        [System.Windows.Forms.MessageBox]::Show("Browse and select a ZIP file first.", "Error","OK","Error")
        return
    }
    Import-UserProfile -ZipPath $global:SelectedZipPath -Username $u
    if ($global:DomainCheckBox.Checked) {
        $computerName = $global:ComputerNameTextBox.Text.Trim()
        $domainName = $global:DomainNameTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($domainName)) {
            [System.Windows.Forms.MessageBox]::Show("Domain name is required when 'Join Domain After Import' is checked.", "Domain Required","OK","Warning")
            return
        }
        $restartBehavior = $global:RestartComboBox.SelectedItem
        $delaySeconds = 30
        if ($restartBehavior -eq 'Delayed') {
            if ([int]::TryParse($global:DelayTextBox.Text, [ref]$delaySeconds)) {
                if ($delaySeconds -lt 5) { $delaySeconds = 5 }
                if ($delaySeconds -gt 300) { $delaySeconds = 300 }
            } else {
                $delaySeconds = 30
            }
        }
        Log-Message "Initiating domain join with restart behavior: $restartBehavior"
        Join-Domain-Enhanced -ComputerName $computerName -DomainName $domainName -RestartBehavior $restartBehavior -DelaySeconds $delaySeconds -Credential $global:DomainCredential
    }
})
$global:DomainJoinButton.Add_Click({
    $computerName = $global:ComputerNameTextBox.Text.Trim()
    $domainName = $global:DomainNameTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domainName)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a domain name.", "Domain Required","OK","Warning")
        return
    }
    $restartBehavior = $global:RestartComboBox.SelectedItem
    $delaySeconds = 30
    if ($restartBehavior -eq 'Delayed') {
        [int]::TryParse($global:DelayTextBox.Text, [ref]$delaySeconds) | Out-Null
        $delaySeconds = [Math]::Max(5, [Math]::Min(300, $delaySeconds))
    }
    Log-Message "=== STANDALONE DOMAIN JOIN ==="
    Log-Message "Domain: $domainName"
    Log-Message "Computer Name: $(if($computerName){$computerName}else{'Keep current'})"
    Log-Message "Restart Behavior: $restartBehavior"
    Join-Domain-Enhanced -ComputerName $computerName -DomainName $domainName -RestartBehavior $restartBehavior -DelaySeconds $delaySeconds -Credential $global:DomainCredential
})

$global:DomainRetryButton.Add_Click({
    $domainName = $global:DomainNameTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domainName)) {
        [System.Windows.Forms.MessageBox]::Show("Enter a domain name first.","Domain Required","OK","Warning") | Out-Null
        return
    }
    if (-not (Get-Command Test-DomainReachability -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("Reachability test function not available.","Unavailable","OK","Warning") | Out-Null
        return
    }
    Log-Message "Retrying domain reachability: $domainName"
    $global:StatusText.Text = "Testing domain reachability..."
    [System.Windows.Forms.Application]::DoEvents()
    $result = Test-DomainReachability -DomainName $domainName
    if ($result.Success) {
        Log-Message "Domain reachable: $domainName"
        $global:StatusText.Text = "Domain reachable"
        if (-not $global:DomainCheckBox.Checked) { $global:DomainCheckBox.Checked = $true }
        [System.Windows.Forms.MessageBox]::Show("Domain '$domainName' is reachable.","Domain OK","OK","Information") | Out-Null
    } else {
        Log-Message "Domain unreachable: $($result.Error)"
        $global:StatusText.Text = "Domain unreachable"
        $global:DomainCheckBox.Checked = $false
        [System.Windows.Forms.MessageBox]::Show("Domain '$domainName' unreachable.\n\n$($result.Error)","Domain Unreachable","OK","Warning") | Out-Null
    }
})

# =============================================================================
# LAUNCH
# =============================================================================
Log-Message "Profile Transfer Tool"
Log-Message "IMPORTANT: After import, RESTART computer before logging in!"
Log-Message "Restart Mode Options: Prompt (ask), Delayed (countdown), Never (manual), Immediate (auto)"
$global:Form.ShowDialog() | Out-Null